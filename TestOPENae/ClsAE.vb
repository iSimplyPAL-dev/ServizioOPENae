Imports System.Data.OleDb
Imports System.Configuration
Imports System

Imports System.ComponentModel
Imports System.Data
Imports System.Text
Imports log4net

Public Class GetConnectionDB

    Public StrConnDBOpen As String
    Public MyCnString As String

    Public Sub New()
        'connessione al DB di RDB
        'AppReader = New System.Configuration.AppSettingsReader

        'MyCnString = ConfigurationSettings.AppSettings("ConnessioneACCESS").ToString()
        'MyCnString = MyCnString & CType(AppReader.GetValue("PathACCESS", GetType(String)), String)
        'MyCnString = MyCnString & CType(AppReader.GetValue("Database", GetType(String)), String)

    End Sub


    Protected Function GetConnectionACCESS(ByVal myconn As String) As OleDbConnection
        Dim MyConnection As OleDbConnection

        MyCnString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="
        MyCnString = MyCnString & "C:\RDB\DATABASE\rdb.mdb"

        MyConnection = New OleDbConnection(myconn)

        MyConnection.Open()
        GetConnectionACCESS = MyConnection
    End Function

    Protected Sub CloseConnectionACCESS(ByVal MyConn As OleDbConnection)
        MyConn.Close()
        MyConn = Nothing
    End Sub

    Public Function GetDataReaderACCESS(ByVal MySQL As String, ByVal myconn As String, ByRef StrErrore As String) As OleDbDataReader
        Dim MyCn As OleDbConnection = GetConnectionACCESS(myconn)
        Dim MyDataReader As OleDbDataReader
        Dim MyCm As New OleDbCommand(MySQL, MyCn)
        MyCm.CommandTimeout = 900

        Try
            MyDataReader = MyCm.ExecuteReader(CommandBehavior.CloseConnection)
            MyCm.Dispose()
            Return MyDataReader
        Catch ex As Exception
            StrErrore = ex.Message
        End Try

    End Function

	Public Sub RunExecuteQueryACCESS(ByVal MySQL As String, ByVal myconn As String)
		Dim MyCn As OleDbConnection = GetConnectionACCESS(myconn)
		Dim MyCm As New OleDbCommand(MySQL, MyCn)
		MyCm.CommandTimeout = 900

		Try
			MyCm.ExecuteNonQuery()
			MyCm.Dispose()
		Catch EX As Exception
			MessageBox.Show("Errore: " & EX.Message, "RUN EXECUTE QUERY", MessageBoxButtons.OK, MessageBoxIcon.Error)
			Throw EX
		Finally
			CloseConnectionACCESS(MyCn)
		End Try
	End Sub
End Class

Public Class EstrazRDB
	Private Shared Log As ILog = LogManager.GetLogger("EstrazRDB")

	Public Function PreparazioneDati(ByVal sAnnoRuolo As String, ByVal sDal As String, ByVal sCodIstat As String, ByVal SCODFISCALE As String, ByVal SDENOMINAZIONECOM As String, ByVal SCOMUNELEGALE As String, ByVal SPROVLEG As String, ByVal sConnDB As String) As Boolean
		Dim sSQL, CodImmobile, sDataInizioOccup As String
		Dim DataDal, DataAl As Date
		Dim ProgEstrazione As Integer
		Dim ProgUnivoco As Integer = 0
		Dim oDBManager As New GetConnectionDB
		Dim drResult As OleDb.OleDbDataReader
		Dim drResultRif As OleDb.OleDbDataReader
		Dim oMyTariffe() As Tariffe
		Dim oMyDatiFlusso As ObjDatiFlusso
		Dim MyArray() As String

		Try
			Try
				'pulisco la tabella dati flusso dai record non consolidati
				sSQL = "DELETE * FROM _AE_DATI_FLUSSO WHERE CONSOLIDATO =0"
				oDBManager.RunExecuteQueryACCESS(sSQL, sConnDB)
				sSQL = "DELETE [_AE_ESTRAZIONI].*, [_AE_ESTRAZIONI].PROG_ESTRAZIONE FROM _AE_ESTRAZIONI WHERE ((([_AE_ESTRAZIONI].PROG_ESTRAZIONE) Not In (SELECT PROG_ESTRAZIONE FROM _AE_DATI_FLUSSO)))"
				oDBManager.RunExecuteQueryACCESS(sSQL, sConnDB)
			Catch ex As Exception
				'creo tabella dati flusso
				sSQL = "CREATE TABLE _AE_DATI_FLUSSO ("
				sSQL += " PROG_ESTRAZIONE LONG "
				sSQL += " PROGRESSIVO DOUBLE, "
				sSQL += " ABBINATO LONG, CONSOLIDATO LONG, "
				sSQL += " COD_CONTRIB LONG, NSCHEDA TEXT(10), "
				sSQL += " CFPIVA TEXT(16), COGNOME TEXT(26), NOME TEXT(25), "
				sSQL += " DENOMINAZIONE TEXT(50), COMUNE TEXT(40), PROVINCIA TEXT(2), "
				sSQL += " TITOLO_OCCUPAZIONE TEXT(1), OCCUPANTE TEXT(1), "
				sSQL += " DI_OCCUPAZIONE TEXT(10), DF_OCCUPAZIONE TEXT(10), "
				sSQL += " DI_RIFC TEXT(10), DF_RIFC TEXT(10), "
				sSQL += " DESTINAZIONE_USO TEXT(1), TIPO_UNITA TEXT(1), "
				sSQL += " SEZIONE TEXT(3), FOGLIO TEXT(5), PARTICELLA TEXT(5), EST_PARTICELLA TEXT(4),"
				sSQL += " TIPO_PARTICELLA TEXT(1) , SUBALTERNO TEXT(4),"
				sSQL += " UBICAZIONE TEXT(30), CIVICO TEXT(6), INTERNO TEXT(2), SCALA TEXT(1), CADC TEXT(1),"
				sSQL += " DENOMINAZIONECOM TEXT(60), CODFISCALE TEXT(16), COMUNELEGALE TEXT(40), PROVLEG TEXT(2), "
				sSQL += " COMUNEAMM TEXT(20), PROV TEXT(2), COMUNECAT TEXT(20), CODICE TEXT(5), CODISTAT TEXT(6) "
				sSQL += ")"
				oDBManager.RunExecuteQueryACCESS(sSQL, sConnDB)

				'creo tabella dati estrazione
				sSQL = "CREATE TABLE _AE_ESTRAZIONI ("
				sSQL += " PROG_ESTRAZIONE DOUBLE, "
				sSQL += " DATA_CREAZIONE DATETIME, DATA_ESTRAZIONE DATETIME, ANNO_RUOLO TEXT (4))"
				oDBManager.RunExecuteQueryACCESS(sSQL, sConnDB)
			End Try

			'estrapolo il progressivo estrazione
			ProgEstrazione = ProgressivoEstrazione(sAnnoRuolo, oDBManager, sConnDB)
			oMyTariffe = PrelevaTariffe(sAnnoRuolo, oDBManager, sConnDB)
			'***  IL DAL LO DEVE PRENDERE DA OCCUPANTI QUELLO DI RUOLO E' UN DAL FITTIZIO ***
			sSQL = "SELECT DISTINCT RUOLO_TRRSU.COD_IMMOBILE, RUOLO_TRRSU.COD_OCCUPANTE, RUOLO_TRRSU.OCCUPATO_DAL, RUOLO_TRRSU.BIMESTRI, RUOLO_TRRSU.CATEGORIA, RUOLO_TRRSU.CIVICO, RIDUZIONI.COD_RIDUZIONE, ANAGRAFE.*, STRADARIO.*, IIf(Not [COD FISCALE] Is Null,[COD FISCALE],[PARTITA IVA]) AS CFPIVA, OCCUPANTI.[OCCUPATO DAL]"
			sSQL += " FROM (((RUOLO_TRRSU"
			sSQL += " INNER JOIN ANAGRAFE ON RUOLO_TRRSU.COD_OCCUPANTE = ANAGRAFE.[COD ANAGRAFICO])"
			sSQL += " LEFT JOIN STRADARIO ON RUOLO_TRRSU.COD_STRADA = STRADARIO.[COD STRADA])"
			sSQL += " LEFT JOIN RIDUZIONI ON (RUOLO_TRRSU.COD_OCCUPANTE = RIDUZIONI.[COD OCCUPANTE]) AND (RUOLO_TRRSU.ANNO = RIDUZIONI.ANNO) AND (RUOLO_TRRSU.COD_IMMOBILE = RIDUZIONI.[COD IMMOBILE]))"
			sSQL += " INNER JOIN OCCUPANTI ON (RUOLO_TRRSU.COD_IMMOBILE = OCCUPANTI.[COD IMMOBILE]) AND (RUOLO_TRRSU.COD_OCCUPANTE = OCCUPANTI.[COD OCCUPANTE])"
			sSQL += " WHERE(((RUOLO_TRRSU.ANNO) = " & sAnnoRuolo & ") AND ((OCCUPANTI.[OCCUPATO DAL]) >= #" & sDal & "#))"
			sSQL += " AND  (([RUOLO_TRRSU].[COD_IMMOBILE] & CStr([RUOLO_TRRSU].[MQ])) IN ("
			sSQL += "    SELECT COD_IMMOBILE & CSTR(MAX(MQ))"
			sSQL += "    FROM RUOLO_TRRSU"
			sSQL += "    INNER JOIN OCCUPANTI ON (RUOLO_TRRSU.COD_OCCUPANTE = OCCUPANTI.[COD OCCUPANTE]) AND (RUOLO_TRRSU.COD_IMMOBILE = OCCUPANTI.[COD IMMOBILE])"
			sSQL += "    WHERE (ANNO=" & sAnnoRuolo & ") AND (OCCUPANTI.[OCCUPATO DAL] >= #" & sDal & "#)"
			sSQL += "    GROUP BY COD_IMMOBILE"
			sSQL += " ))"
			sSQL += " ORDER BY RUOLO_TRRSU.COD_OCCUPANTE, RUOLO_TRRSU.COD_IMMOBILE, RUOLO_TRRSU.OCCUPATO_DAL"
			'********************************************************************************
			drResult = oDBManager.GetDataReaderACCESS(sSQL, sConnDB, "")
			Do While drResult.Read
				oMyDatiFlusso = New ObjDatiFlusso
				oMyDatiFlusso.CODISTAT = sCodIstat
				oMyDatiFlusso.CODFISCALE = SCODFISCALE
				oMyDatiFlusso.DENOMINAZIONECOM = SDENOMINAZIONECOM
				oMyDatiFlusso.COMUNELEGALE = SCOMUNELEGALE
				oMyDatiFlusso.PROVLEGALE = SPROVLEG

				If ControlloCampo(drResult("cod_immobile")) = "" Then				'se non c'è codice immobile non posso valorizzare tutti i dati
					oMyDatiFlusso.PROGESTRAZIONE = ProgEstrazione
					ProgUnivoco += 1
					oMyDatiFlusso.PROGUNIVOCO = ProgUnivoco
					oMyDatiFlusso.CONSOLIDATO = 0
					oMyDatiFlusso.COD_OCCUPANTE = drResult("cod_occupante")
					oMyDatiFlusso.CFPIVA = drResult("cfpiva")
					If drResult("sesso") = "G" Then
						oMyDatiFlusso.DENOMINAZIONE = Mid(ControlloCampo(drResult("cognome")), 1, 50)
						oMyDatiFlusso.CITTA = Mid(ControlloCampo(drResult("CITTA RES")), 1, 40)
						oMyDatiFlusso.PROVINCIA = Mid(ControlloCampo(drResult("PROVINCIA RES")), 1, 2)
					Else
						oMyDatiFlusso.COGNOME = Mid(ControlloCampo(drResult("cognome")), 1, 26)
						oMyDatiFlusso.NOME = Mid(ControlloCampo(drResult("nome")), 1, 25)
					End If
					oMyDatiFlusso.TITOLOOCCUPAZIONE = 4
					'controllo se l'occupante è unico occupante (da elenco single o da riduzioni) e attribuisco in base a configurazione manuale
					oMyDatiFlusso.OCCUPANTE = ControlloOccupante(oMyTariffe, drResult("categoria"), ControlloCampo(drResult("cod_riduzione")), ControlloCampo(drResult("cfpiva")), "")

					If ControlloCampo(drResult("occupato_dal")) = "" Then
						oMyDatiFlusso.DATAINIZIO = +"31/12/" & CShort(sAnnoRuolo) - 1
					Else
						oMyDatiFlusso.DATAINIZIO = ControlloCampo(drResult("occupato_dal"))
					End If
					sDataInizioOccup = drResult("occupato_dal")
					Call CalcolaFineOccupazione(drResult("Bimestri"), sDataInizioOccup)
					If CDate(sDataInizioOccup) > CDate("31/12/" & sAnnoRuolo) Then
						sDataInizioOccup = "31/12/" & sAnnoRuolo
					End If
					oMyDatiFlusso.DATAFINE = sDataInizioOccup
					oMyDatiFlusso.DESTINAZIONEUSO = 1
					oMyDatiFlusso.ABBINATO = 0
					oMyDatiFlusso.TIPOUNITA = "F"
					oMyDatiFlusso.SEZIONE = ""
					oMyDatiFlusso.FOGLIO = ""
					oMyDatiFlusso.NUMERO = ""
					oMyDatiFlusso.ESTPARTICELLA = ""
					oMyDatiFlusso.TIPOPARTICELLA = ""
					oMyDatiFlusso.SUBALTERNO = ""
					oMyDatiFlusso.UBICAZIONE = Mid(ControlloCampo(drResult("TIP STRADA")) & " " & ControlloCampo(drResult("STRADA")), 1, 30)
					oMyDatiFlusso.CIVICO = Mid(ControlloCampo(drResult("CIVICO")), 1, 6)
					oMyDatiFlusso.INTERNO = ""
					oMyDatiFlusso.SCALA = ""
					oMyDatiFlusso.CADC = "3"
					If SetDatiFlusso(oMyDatiFlusso, oDBManager, sConnDB) = False Then
						Log.Debug("EstrazRDB::errore in popolamento _AE_DATI_FLUSSO::1")
					End If
				Else

					'attraverso il codice scheda reperisco i riferimenti catastali
					DataDal = CDate("31/12/" & sAnnoRuolo)
					DataAl = CDate("01/01/" & CDbl(sAnnoRuolo) + 1)
					CodImmobile = drResult("cod_immobile")

					sSQL = "SELECT distinct Catasto_p.CodiceImmobile, Catasto_p.sezione, Catasto_p.Foglio, Catasto_p.Numero, Catasto_p.Sub"					  ', Min(Catasto_p.Dal) AS MinDiDal, Max(IIf(IsNull([CATASTO_P.AL),#" & DataDal & "#,[CATASTO_P.AL)) AS MaxDiAl"
					sSQL += " FROM Catasto_p INNER JOIN Catasto_s ON Catasto_p.CodiceCatasto = Catasto_s.CodiceCatasto"
					sSQL += " WHERE ((Catasto_p.Dal<=#" & DataDal & "# AND Catasto_p.Al Is Null) OR (Catasto_p.Dal<=#" & DataDal & "# AND Catasto_p.Al>=#" & DataAl & "#))"
					sSQL += " AND ((Catasto_s.Dal<=#" & DataDal & "# AND Catasto_s.Al Is Null) OR (Catasto_s.Dal<=#" & DataDal & "# AND Catasto_s.Al>=#" & DataAl & "#))"
					sSQL += " AND (Catasto_s.TipoRendita='RE')"
					sSQL += " AND (Catasto_p.CodiceImmobile = """ & CodImmobile & """)"
					sSQL += " GROUP BY Catasto_p.CodiceImmobile, Catasto_p.Sezione, Catasto_p.Foglio, "
					sSQL += " Catasto_p.Numero, Catasto_p.Sub"
					drResultRif = oDBManager.GetDataReaderACCESS(sSQL, sConnDB, "")
					If drResultRif.HasRows Then
						Do While drResultRif.Read
							oMyDatiFlusso = New ObjDatiFlusso
							oMyDatiFlusso.CODFISCALE = SCODFISCALE
							oMyDatiFlusso.DENOMINAZIONECOM = SDENOMINAZIONECOM
							oMyDatiFlusso.COMUNELEGALE = SCOMUNELEGALE
							oMyDatiFlusso.PROVLEGALE = SPROVLEG
							'trovato riferimenti catastali attraverso codice immobile
							oMyDatiFlusso.PROGESTRAZIONE = ProgEstrazione
							ProgUnivoco += 1
							oMyDatiFlusso.PROGUNIVOCO = ProgUnivoco
							oMyDatiFlusso.CONSOLIDATO = "0"
							oMyDatiFlusso.COD_OCCUPANTE = drResult("cod_occupante")
							oMyDatiFlusso.CFPIVA = drResult("cfpiva")
							oMyDatiFlusso.COD_IMMOBILE = drResult("cod_immobile")
							If drResult("sesso") = "G" Then
								oMyDatiFlusso.DENOMINAZIONE = Mid(ControlloCampo(drResult("cognome")), 1, 50)
								oMyDatiFlusso.CITTA = Mid(ControlloCampo(drResult("CITTA RES")), 1, 40)
								oMyDatiFlusso.PROVINCIA = Mid(ControlloCampo(drResult("PROVINCIA RES")), 1, 2)
							Else
								oMyDatiFlusso.COGNOME = Mid(ControlloCampo(drResult("cognome")), 1, 26)
								oMyDatiFlusso.NOME = Mid(ControlloCampo(drResult("nome")), 1, 25)
							End If

							'controllo se l'occupante è tra i proprietari attivi
							'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							oMyDatiFlusso.TITOLOOCCUPAZIONE = OccupanteAttivo(drResult("cod_occupante"), drResult("cod_immobile"), DataDal, DataAl, oDBManager, sConnDB)

							'controllo se l'occupante è unico occupante (da elenco single o da riduzioni) e attribuisco in base a configurazione manuale
							oMyDatiFlusso.OCCUPANTE = ControlloOccupante(oMyTariffe, drResult("categoria"), ControlloCampo(drResult("cod_riduzione")), ControlloCampo(drResult("cfpiva")), "")
							If ControlloCampo(drResult("occupato_dal")) = "" Then
								oMyDatiFlusso.DATAINIZIO = "31/12/" & CShort(sAnnoRuolo) - 1
							Else
								oMyDatiFlusso.DATAINIZIO = ControlloCampo(drResult("occupato_dal"))
							End If

							'recupero i dati di occupazione
							MyArray = Split(DatiOccupazione(oMyTariffe, ControlloCampo(drResult("categoria")), sAnnoRuolo, drResult("cod_immobile"), drResult("cod_occupante"), ControlloCampo(drResult("occupato_dal")), oDBManager, sConnDB), "-")
							sDataInizioOccup = drResult("occupato_dal")
							Call CalcolaFineOccupazione(drResult("Bimestri"), sDataInizioOccup)
							If CDate(sDataInizioOccup) > CDate("31/12/" & sAnnoRuolo) Then
								sDataInizioOccup = "31/12/" & sAnnoRuolo
							End If
							oMyDatiFlusso.DATAFINE = sDataInizioOccup

							oMyDatiFlusso.DESTINAZIONEUSO = MyArray(1)
							oMyDatiFlusso.ABBINATO = "1"
							'memorizzo i dati relativi a inizio e fine riferimento catastale principale
							oMyDatiFlusso.TIPOUNITA = "F"
							oMyDatiFlusso.SEZIONE = Mid(ControlloCampo(drResultRif("SEZIONE")), 1, 3)
							oMyDatiFlusso.FOGLIO = Mid(ControlloCampo(drResultRif("FOGLIO")), 1, 5)
							oMyDatiFlusso.NUMERO = Mid(ControlloCampo(drResultRif("numero")), 1, 5)
							oMyDatiFlusso.ESTPARTICELLA = ""
							oMyDatiFlusso.TIPOPARTICELLA = ""
							oMyDatiFlusso.SUBALTERNO = Mid(ControlloCampo(drResultRif("Sub")), 1, 4)
							oMyDatiFlusso.UBICAZIONE = Mid(ControlloCampo(drResult("TIP STRADA")) & " " & ControlloCampo(drResult("STRADA")), 1, 30)
							oMyDatiFlusso.CIVICO = Mid(ControlloCampo(drResult("CIVICO")), 1, 6)
							oMyDatiFlusso.INTERNO = ""
							oMyDatiFlusso.SCALA = ""
							If ControlloCampo(drResultRif("FOGLIO")) = "" Or ControlloCampo(drResultRif("numero")) = "" Then
								oMyDatiFlusso.CADC = "3"
							Else
								oMyDatiFlusso.CADC = "0"
							End If
							If SetDatiFlusso(oMyDatiFlusso, oDBManager, sConnDB) = False Then
								Log.Debug("EstrazRDB::errore in popolamento _AE_DATI_FLUSSO::2")
							End If
						Loop
					Else
						oMyDatiFlusso = New ObjDatiFlusso
						oMyDatiFlusso.CODFISCALE = SCODFISCALE
						oMyDatiFlusso.DENOMINAZIONECOM = SDENOMINAZIONECOM
						oMyDatiFlusso.COMUNELEGALE = SCOMUNELEGALE
						oMyDatiFlusso.PROVLEGALE = SPROVLEG
						oMyDatiFlusso.PROGESTRAZIONE = ProgEstrazione.ToString
						ProgUnivoco += 1
						oMyDatiFlusso.PROGUNIVOCO = ProgUnivoco.ToString
						oMyDatiFlusso.CONSOLIDATO = "0"
						oMyDatiFlusso.COD_OCCUPANTE = CStr(drResult("cod_occupante"))
						oMyDatiFlusso.CFPIVA = CStr(drResult("cfpiva"))
						oMyDatiFlusso.COD_IMMOBILE = CStr(drResult("cod_immobile"))
						If CStr(drResult("sesso")) = "G" Then
							oMyDatiFlusso.DENOMINAZIONE = Mid(ControlloCampo(drResult("cognome")), 1, 50)
							oMyDatiFlusso.CITTA = Mid(ControlloCampo(drResult("CITTA RES")), 1, 40)
							oMyDatiFlusso.PROVINCIA = Mid(ControlloCampo(drResult("PROVINCIA RES")), 1, 2)
						Else
							oMyDatiFlusso.COGNOME = Mid(ControlloCampo(drResult("cognome")), 1, 26)
							oMyDatiFlusso.NOME = Mid(ControlloCampo(drResult("nome")), 1, 25)
						End If
						oMyDatiFlusso.TITOLOOCCUPAZIONE = OccupanteAttivo(drResult("cod_occupante"), drResult("cod_immobile"), DataDal, DataAl, oDBManager, sConnDB)

						'controllo se l'occupante è unico occupante (da elenco single o da riduzioni) e attribuisco in base a configurazione manuale
						oMyDatiFlusso.OCCUPANTE = ControlloOccupante(oMyTariffe, drResult("categoria"), ControlloCampo(drResult("cod_riduzione")), ControlloCampo(drResult("cfpiva")), "")

						If ControlloCampo(drResult("occupato_dal")) = "" Then
							oMyDatiFlusso.DATAINIZIO = "31/12/" & CShort(sAnnoRuolo) - 1
						Else
							oMyDatiFlusso.DATAINIZIO = ControlloCampo(drResult("occupato_dal"))
						End If
						sDataInizioOccup = CStr(drResult("occupato_dal"))
						Call CalcolaFineOccupazione(drResult("Bimestri"), sDataInizioOccup)
						If CDate(sDataInizioOccup) > CDate("31/12/" & sAnnoRuolo) Then
							sDataInizioOccup = "31/12/" & sAnnoRuolo
						End If
						oMyDatiFlusso.DATAFINE = sDataInizioOccup
						MyArray = Split(DatiOccupazione(oMyTariffe, ControlloCampo(drResult("Categoria")), sAnnoRuolo, drResult("cod_immobile"), drResult("cod_occupante"), ControlloCampo(drResult("occupato_dal")), oDBManager, sConnDB), "-")
						oMyDatiFlusso.DESTINAZIONEUSO = MyArray(1)
						oMyDatiFlusso.ABBINATO = "0"
						oMyDatiFlusso.TIPOUNITA = "F"
						oMyDatiFlusso.SEZIONE = ""
						oMyDatiFlusso.FOGLIO = ""
						oMyDatiFlusso.NUMERO = ""
						oMyDatiFlusso.ESTPARTICELLA = ""
						oMyDatiFlusso.TIPOPARTICELLA = ""
						oMyDatiFlusso.SUBALTERNO = ""
						oMyDatiFlusso.UBICAZIONE = Mid(ControlloCampo(drResult("TIP STRADA")) & " " & ControlloCampo(drResult("STRADA")), 1, 30)
						oMyDatiFlusso.CIVICO = Mid(ControlloCampo(drResult("CIVICO")), 1, 6)
						oMyDatiFlusso.INTERNO = ""
						oMyDatiFlusso.SCALA = ""
						oMyDatiFlusso.CADC = "3"
						If SetDatiFlusso(oMyDatiFlusso, oDBManager, sConnDB) = False Then
							Log.Debug("EstrazRDB::errore in popolamento _AE_DATI_FLUSSO::3")
							Return False
						End If
					End If
					drResultRif.Close()
					'controllo su presenza dati catastali
				End If				'controllo su presenza codice immobile
			Loop
			Return True
		Catch err As Exception
			Log.Debug("EstrazRDB::si è verificato il seguente errore::" + err.Message)
			Return False
		Finally
			drResult.Close()
		End Try
	End Function
	Private Function ControlloCampo(ByRef Valore As Object) As String
		If IsDBNull(Valore) Then
			ControlloCampo = ""
		Else
			ControlloCampo = Valore
		End If
	End Function

	Private Function ControlloOccupante(ByVal Tariffe() As Tariffe, ByVal sCategoria As String, ByVal sCodRid As String, ByVal sCFPIVA As String, ByVal sFileSingle As String) As String
		Dim i As Object
		For i = 0 To Tariffe.GetUpperBound(0)
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Tariffe(i).sCategoria = sCategoria Then
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				ControlloOccupante = Tariffe(i).sControlloOccupante
				'se sono su nucleo familiare controllo l'attribuzione a singolo
				If ControlloOccupante = "2" Then
					'se ho il file dei single cerco l'utente da file
					If sFileSingle <> "" Then
						ControlloOccupante = CStr(OccupanteSingolo(sCFPIVA, sFileSingle))
					Else
						'controllo se l'utente ha la riduzione unico occupante
						If sCodRid = "UO" Then
							ControlloOccupante = "1"
						End If
					End If
				End If
				Exit For
			End If
		Next i
	End Function

	Private Function OccupanteSingolo(ByRef cfpiva As String, Optional ByRef FileSingle As String = "") As Short
		On Error GoTo ERRORE

		Dim nFileNum As Short
		Dim sNextLine As String

		nFileNum = FreeFile()

		FileOpen(nFileNum, FileSingle, OpenMode.Input)
		Do While Not EOF(nFileNum)
			sNextLine = LineInput(nFileNum)
			'controllo che il codice fiscale che sto esaminando sia un single o no
			If InStr(sNextLine, cfpiva) <> 0 Then
				OccupanteSingolo = CShort("1")
				Exit Do
			Else
				OccupanteSingolo = CShort("2")
			End If
		Loop

		FileClose(nFileNum)


		Exit Function

ERRORE:
		If Err.Number = 3078 Then
			OccupanteSingolo = CShort("2")
			Exit Function
		Else
			MsgBox(Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		End If
	End Function

	Private Sub CalcolaFineOccupazione(ByRef Bimestri As Short, ByRef DataF As Object)
		Dim Mesi As Short
		Dim DataI As Date

		'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DataI = CDate(DataF)

		'se data inizio corrisponde a fine mese della data inizio non faccio nulla ai mesi
		'altrimenti ai mesi ottenuti moltiplicando i bimestri per 2 toglo 1 mese
		Select Case Month(CDate(DataI))
			Case 11, 4, 6, 9
				DataI = CDate("30" & Mid(CStr(DataI), 3))
			Case 2
				DataI = CDate("28" & Mid(CStr(DataI), 3))
			Case Else
				DataI = CDate("31" & Mid(CStr(DataI), 3))
		End Select

		'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If DataI = DataF Then
			Mesi = Bimestri * 2
		Else
			Mesi = Bimestri * 2 - 1
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DataF = DateAdd(Microsoft.VisualBasic.DateInterval.Month, Mesi, DataF)
		'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Select Case Month(CDate(DataF))
			Case 11, 4, 6, 9
				'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DataF = "30" & Mid(DataF, 3)
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DataF = "28" & Mid(DataF, 3)
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object DataF. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DataF = "31" & Mid(DataF, 3)
		End Select

	End Sub

	Private Function ProgressivoEstrazione(ByRef AnnoRuolo As String, ByVal oDBManager As GetConnectionDB, ByVal sConnessioneDB As String) As Integer
		Dim drResult As OleDb.OleDbDataReader
		Dim sSQL As String

		Try
			sSQL = "SELECT * FROM _AE_ESTRAZIONI ORDER BY PROG_ESTRAZIONE DESC"
			drResult = oDBManager.GetDataReaderACCESS(sSQL, sConnessioneDB, "")
			If drResult.HasRows = True Then
				ProgressivoEstrazione = drResult("PROG_ESTRAZIONE").Value + 1
			Else
				ProgressivoEstrazione = 1
			End If
			sSQL = "INSERT INTO _AE_ESTRAZIONI(PROG_ESTRAZIONE,DATA_CREAZIONE,ANNO_RUOLO)"
			sSQL += "VALUES(" + ProgressivoEstrazione.ToString + ",#" + Now.ToShortDateString + "#," + AnnoRuolo + ")"
			oDBManager.RunExecuteQueryACCESS(sSQL, sConnessioneDB)
		Catch err As Exception
			Log.Debug("ProgressivoEstrazione::si è verificato il seguente errore::" + err.Message)
			Return -1
		Finally
			drResult.Close()
		End Try
	End Function

	Private Function OccupanteAttivo(ByRef cod_occupante As Short, ByRef cod_immobile As String, ByRef tDataDal As Date, ByRef tDataAl As Date, ByVal oDBManager As GetConnectionDB, ByVal sConnessioneDB As String) As String
		Dim sSQL As String
		Dim drResult As OleDb.OleDbDataReader

		Try
			'controllo se l'occupante è tra i proprietari attivi
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sSQL = "SELECT Proprietari_Catasto.QuotaIci, Proprietari_Catasto.CodiceProprietario, Catasto_p.CodiceImmobile,Proprietari_Catasto.tit_possesso"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sSQL += " FROM Catasto_p INNER JOIN Proprietari_Catasto ON Catasto_p.CodiceCatasto = Proprietari_Catasto.CodiceCatasto"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sSQL += " WHERE (Proprietari_Catasto.QuotaIci<>0)"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sSQL += " AND (Proprietari_Catasto.CodiceProprietario=" & cod_occupante & ")"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sSQL += " AND (Catasto_p.CodiceImmobile='" & cod_immobile & "')"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			sSQL += " AND ((Catasto_p.Dal<=#" & tDataDal & "# AND Catasto_p.Al Is Null) OR (Catasto_p.Dal<=#" & tDataDal & "# AND Catasto_p.Al>=#" & tDataAl & "#))"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			drResult = oDBManager.GetDataReaderACCESS(sSQL, sConnessioneDB, "")
			If drResult.Read Then
				Select Case drResult("tit_possesso")
					Case 0, 1, 4
						'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						OccupanteAttivo = "1"
					Case 2
						'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						OccupanteAttivo = "4"
					Case 3
						'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						OccupanteAttivo = "2"
					Case 6
						'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						OccupanteAttivo = "3"
					Case Else
						'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						OccupanteAttivo = "4"
				End Select
			Else
				drResult.Close()
				'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				sSQL = "SELECT Proprietari_Catasto.QuotaIci, Proprietari_Catasto.CodiceProprietario, Catasto_p.CodiceImmobile,Proprietari_Catasto.tit_possesso"
				'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				sSQL += " FROM Catasto_p INNER JOIN Proprietari_Catasto ON Catasto_p.CodiceCatasto = Proprietari_Catasto.CodiceCatasto"
				'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				sSQL += " WHERE (Proprietari_Catasto.QuotaIci<>0) "
				'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				sSQL += " AND (Catasto_p.CodiceImmobile='" & cod_immobile & "')"
				'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				sSQL += " AND ((Catasto_p.Dal<=#" & tDataDal & "# AND Catasto_p.Al Is Null) OR (Catasto_p.Dal<=#" & tDataDal & "# AND Catasto_p.Al>=#" & tDataAl & "#))"
				'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				drResult = oDBManager.GetDataReaderACCESS(sSQL, sConnessioneDB, "")
				If drResult.Read Then
					Select Case drResult("tit_possesso")
						Case 0, 1, 4
							'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							OccupanteAttivo = "1"
						Case 2
							'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							OccupanteAttivo = "4"
						Case 3
							'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							OccupanteAttivo = "2"
						Case 6
							'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							OccupanteAttivo = "3"
						Case Else
							'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							OccupanteAttivo = "4"
					End Select
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					OccupanteAttivo = "4"
				End If
			End If
			Return OccupanteAttivo
		Catch ex As Exception
			Log.Debug("OccupanteAttivo::si è verificato il seguente errore::" + ex.Message)
			Return ""
		Finally
			drResult.Close()
		End Try
	End Function

	Private Function DatiOccupazione(ByVal Tariffe() As Tariffe, ByVal Categoria As String, ByRef AnnoRuoloE As String, ByRef cod_immobile As String, ByRef cod_occupante As Double, ByRef occupato_dal As String, ByVal oDBManager As GetConnectionDB, ByVal sConnessioneDB As String) As String
		Dim sSQL, DataOccupazione As String
		Dim drResult As OleDb.OleDbDataReader
		Dim i As Integer

		Try
			If occupato_dal = "" Then
				DataOccupazione = "31/12/" & AnnoRuoloE

				For i = 0 To Tariffe.GetUpperBound(0)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If Tariffe(i).sCategoria = Categoria Then
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						DatiOccupazione = DatiOccupazione & "-" & Tariffe(i).sDatiOccupazione
						Exit For
					End If
				Next i

				Exit Function
			End If

			sSQL = "SELECT OCCUPANTI.AL,[COD STATO] From OCCUPANTI"
			sSQL += " WHERE (((OCCUPANTI.[COD IMMOBILE])='" & cod_immobile & "')"
			sSQL += " AND ((OCCUPANTI.[COD OCCUPANTE])=" & cod_occupante & ") "
			sSQL += " AND ((OCCUPANTI.[OCCUPATO DAL])=#" & occupato_dal & "#) "
			sSQL += ")"
			'UPGRADE_WARNING: Couldn't resolve default property of object SQL. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			drResult = oDBManager.GetDataReaderACCESS(sSQL, sConnessioneDB, "")
			If drResult.Read Then
				DatiOccupazione = ControlloCampo(drResult("al"))
				'UPGRADE_WARNING: Couldn't resolve default property of object drResult("COD STATO. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If drResult("COD STATO") = "NO" Then
					DatiOccupazione = DatiOccupazione & "-" & "2"
				Else
					For i = 0 To Tariffe.GetUpperBound(0)
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If Tariffe(i).sCategoria = Categoria Then
							'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							DatiOccupazione = DatiOccupazione & "-" & Tariffe(i).sDatiOccupazione
							Exit For
						End If
					Next i
				End If
			Else
				DataOccupazione = "31/12/" & AnnoRuoloE

				For i = 0 To Tariffe.GetUpperBound(0)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If Tariffe(i).sCategoria = Categoria Then
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						DatiOccupazione = DatiOccupazione & "-" & Tariffe(i).sDatiOccupazione
						Exit For
					End If
				Next i
			End If
			Return DatiOccupazione
			'UPGRADE_WARNING: Couldn't resolve default property of object drResult.Close. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Catch ex As Exception
			Log.Debug("DatiOccupazione::si è verificato il seguente errore::" + ex.Message)
			Return ""
		Finally
			drResult.Close()
		End Try
	End Function

	Private Function PrelevaTariffe(ByVal sAnno As String, ByVal oDBManager As GetConnectionDB, ByVal sConnessioneDB As String) As Tariffe()
		Dim sSQL As String
		Dim drResult As OleDb.OleDbDataReader
		Dim x As Integer = -1
		Dim oMyTariffa As Tariffe
		Dim oListTariffe() As Tariffe

		Try
			sSQL = "SELECT TARIFFE.CATEGORIA, TARIFFE.DESCRIZIONE, TARIFFE.ANNO"
			sSQL += " From TARIFFE Where (TARIFFE.ANNO= " & sAnno & ")"
			sSQL += " ORDER BY TARIFFE.CATEGORIA"
			drResult = oDBManager.GetDataReaderACCESS(sSQL, sConnessioneDB, "")
			Do While drResult.Read
				If Trim(drResult("descrizione")) <> "" Then
					oMyTariffa = New Tariffe
					oMyTariffa.sCategoria = drResult("categoria")
					If InStr(LCase(drResult("descrizione")), "abitazion") <> 0 Or InStr(LCase(drResult("descrizione")), "abitat") <> 0 Then
						oMyTariffa.sDatiOccupazione = 1
						oMyTariffa.sControlloOccupante = 2
					ElseIf InStr(LCase(drResult("descrizione")), "boi") <> 0 Or InStr(LCase(drResult("descrizione")), "autorimes") <> 0 Then
						oMyTariffa.sDatiOccupazione = 4
						oMyTariffa.sControlloOccupante = 4
					ElseIf InStr(LCase(drResult("descrizione")), "attivit") <> 0 Or InStr(LCase(drResult("descrizione")), "commerc") <> 0 Then
						oMyTariffa.sDatiOccupazione = 3
						oMyTariffa.sControlloOccupante = 3
					Else
						oMyTariffa.sDatiOccupazione = 5
						oMyTariffa.sControlloOccupante = 4
					End If
					x += 1
					ReDim Preserve oListTariffe(x)
					oListTariffe(x) = oMyTariffa
				End If
			Loop
			Return oListTariffe
		Catch ex As Exception
			Log.Debug("PrelevaTariffe::si è verificato il seguente errore::" + ex.Message)
			Return Nothing
		Finally
			drResult.Close()
		End Try
	End Function

	Private Function SetDatiFlusso(ByVal oMyDati As ObjDatiFlusso, ByVal oDBManager As GetConnectionDB, ByVal sConnessioneDB As String)
		Dim sSQLCol, sSQLVal As String

		Try
			sSQLCol += "INSERT INTO _AE_DATI_FLUSSO ("
			sSQLCol += "PROG_ESTRAZIONE"
			sSQLCol += ",CODISTAT,CODFISCALE,DENOMINAZIONECOM,COMUNELEGALE,PROVLEG"
			sSQLCol += ",PROGRESSIVO"
			sSQLCol += ",CONSOLIDATO"
			sSQLCol += ",COD_CONTRIB"
			sSQLCol += ",CFPIVA"
			sSQLCol += ",NSCHEDA"
			sSQLCol += ",COGNOME"
			sSQLCol += ",NOME"
			sSQLCol += ",DENOMINAZIONE"
			sSQLCol += ",COMUNE"
			sSQLCol += ",PROVINCIA"
			sSQLCol += ",TITOLO_OCCUPAZIONE"
			sSQLCol += ",OCCUPANTE"
			sSQLCol += ",DI_OCCUPAZIONE"
			sSQLCol += ",DF_OCCUPAZIONE"
			sSQLCol += ",DESTINAZIONE_USO"
			sSQLCol += ",ABBINATO"
			sSQLCol += ",TIPO_UNITA"
			sSQLCol += ",SEZIONE"
			sSQLCol += ",FOGLIO"
			sSQLCol += ",PARTICELLA"
			sSQLCol += ",EST_PARTICELLA"
			sSQLCol += ",TIPO_PARTICELLA"
			sSQLCol += ",SUBALTERNO"
			sSQLCol += ",UBICAZIONE"
			sSQLCol += ",CIVICO"
			sSQLCol += ",INTERNO"
			sSQLCol += ",SCALA"
			sSQLCol += ",CADC"
			sSQLCol += ")"

			sSQLVal = " VALUES("
			sSQLVal += oMyDati.PROGESTRAZIONE
			sSQLVal += ",'" + oMyDati.CODISTAT.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.CODFISCALE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.DENOMINAZIONECOM.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.COMUNELEGALE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.PROVLEGALE.Replace("'", "''") + "'"
			sSQLVal += "," + oMyDati.PROGUNIVOCO
			sSQLVal += "," + oMyDati.CONSOLIDATO
			sSQLVal += "," + oMyDati.COD_OCCUPANTE
			sSQLVal += ",'" + oMyDati.CFPIVA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.COD_IMMOBILE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.COGNOME.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.NOME.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.DENOMINAZIONE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.CITTA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.PROVINCIA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.TITOLOOCCUPAZIONE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.OCCUPANTE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.DATAINIZIO.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.DATAFINE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.DESTINAZIONEUSO.Replace("'", "''") + "'"
			sSQLVal += "," + oMyDati.ABBINATO
			sSQLVal += ",'" + oMyDati.TIPOUNITA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.SEZIONE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.FOGLIO.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.NUMERO.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.ESTPARTICELLA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.TIPOPARTICELLA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.SUBALTERNO.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.UBICAZIONE.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.CIVICO.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.INTERNO.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.SCALA.Replace("'", "''") + "'"
			sSQLVal += ",'" + oMyDati.CADC.Replace("'", "''") + "'"
			sSQLVal += ")"
			oDBManager.RunExecuteQueryACCESS(sSQLCol + sSQLVal, sConnessioneDB)
			Return True
		Catch ex As Exception
			Log.Debug("SetDatiFlusso::si è verificato il seguente errore::" + ex.Message)
			Return False
		End Try
	End Function
End Class

Public Class Tariffe
	Private _sCategoria As String = ""
	Private _sControlloOccupante As String = ""
	Private _sDatiOccupazione As String = ""

	Public Property sCategoria() As String
		Get
			Return _sCategoria
		End Get

		Set(ByVal Value As String)
			_sCategoria = Value
		End Set
	End Property
	Public Property sControlloOccupante() As String
		Get
			Return _sControlloOccupante
		End Get

		Set(ByVal Value As String)
			_sControlloOccupante = Value
		End Set
	End Property
	Public Property sDatiOccupazione() As String
		Get
			Return _sDatiOccupazione
		End Get

		Set(ByVal Value As String)
			_sDatiOccupazione = Value
		End Set
	End Property
End Class

Public Class ObjDatiFlusso
	Dim _PROGESTRAZIONE As String = ""
	Dim _CODISTAT As String = ""
	Dim _CODFISCALE As String = ""
	Dim _DENOMINAZIONECOM As String = ""
	Dim _COMUNELEGALE As String = ""
	Dim _PROVLEGALE As String = ""
	Dim _PROGUNIVOCO As String = ""
	Dim _CONSOLIDATO As String = ""
	Dim _COD_OCCUPANTE As String = ""
	Dim _CFPIVA As String = ""
	Dim _COD_IMMOBILE As String = ""
	Dim _COGNOME As String = ""
	Dim _NOME As String = ""
	Dim _DENOMINAZIONE As String = ""
	Dim _CITTA As String = ""
	Dim _PROVINCIA As String = ""
	Dim _TITOLOOCCUPAZIONE As String = ""
	Dim _OCCUPANTE As String = ""
	Dim _DATAINIZIO As String = ""
	Dim _DATAFINE As String = ""
	Dim _DESTINAZIONEUSO As String = ""
	Dim _ABBINATO As String = ""
	Dim _TIPOUNITA As String = ""
	Dim _SEZIONE As String = ""
	Dim _FOGLIO As String = ""
	Dim _NUMERO As String = ""
	Dim _ESTPARTICELLA As String = ""
	Dim _TIPOPARTICELLA As String = ""
	Dim _SUBALTERNO As String = ""
	Dim _UBICAZIONE As String = ""
	Dim _CIVICO As String = ""
	Dim _INTERNO As String = ""
	Dim _SCALA As String = ""
	Dim _CADC As String = ""

	Public Property PROGESTRAZIONE() As String
		Get
			Return _PROGESTRAZIONE

		End Get

		Set(ByVal Value As String)
			_PROGESTRAZIONE = Value
		End Set
	End Property
	Public Property CODISTAT() As String
		Get
			Return _CODISTAT

		End Get

		Set(ByVal Value As String)
			_CODISTAT = Value
		End Set
	End Property
	Public Property CODFISCALE() As String
		Get
			Return _CODFISCALE

		End Get

		Set(ByVal Value As String)
			_CODFISCALE = Value
		End Set
	End Property
	Public Property COMUNELEGALE() As String
		Get
			Return _COMUNELEGALE

		End Get

		Set(ByVal Value As String)
			_COMUNELEGALE = Value
		End Set
	End Property
	Public Property DENOMINAZIONECOM() As String
		Get
			Return _DENOMINAZIONECOM

		End Get

		Set(ByVal Value As String)
			_DENOMINAZIONECOM = Value
		End Set
	End Property
	Public Property PROVLEGALE() As String
		Get
			Return _PROVLEGALE

		End Get

		Set(ByVal Value As String)
			_PROVLEGALE = Value
		End Set
	End Property
	Public Property PROGUNIVOCO() As String
		Get
			Return _PROGUNIVOCO

		End Get

		Set(ByVal Value As String)
			_PROGUNIVOCO = Value
		End Set
	End Property
	Public Property CONSOLIDATO() As String
		Get
			Return _CONSOLIDATO
		End Get

		Set(ByVal Value As String)
			_CONSOLIDATO = Value
		End Set
	End Property
	Public Property COD_OCCUPANTE() As String
		Get
			Return _COD_OCCUPANTE
		End Get

		Set(ByVal Value As String)
			_COD_OCCUPANTE = Value
		End Set
	End Property
	Public Property CFPIVA() As String
		Get
			Return _CFPIVA
		End Get

		Set(ByVal Value As String)
			_CFPIVA = Value
		End Set
	End Property
	Public Property COD_IMMOBILE() As String
		Get
			Return _COD_IMMOBILE
		End Get

		Set(ByVal Value As String)
			_COD_IMMOBILE = Value
		End Set
	End Property
	Public Property COGNOME() As String
		Get
			Return _COGNOME
		End Get

		Set(ByVal Value As String)
			_COGNOME = Value
		End Set
	End Property
	Public Property NOME() As String
		Get
			Return _NOME

		End Get

		Set(ByVal Value As String)
			_NOME = Value
		End Set
	End Property
	Public Property DENOMINAZIONE() As String
		Get
			Return _DENOMINAZIONE
		End Get

		Set(ByVal Value As String)
			_DENOMINAZIONE = Value
		End Set
	End Property
	Public Property CITTA() As String
		Get
			Return _CITTA
		End Get

		Set(ByVal Value As String)
			_CITTA = Value
		End Set
	End Property
	Public Property PROVINCIA() As String
		Get
			Return _PROVINCIA
		End Get
		Set(ByVal Value As String)
			_PROVINCIA = Value
		End Set
	End Property
	Public Property TITOLOOCCUPAZIONE() As String
		Get
			Return _TITOLOOCCUPAZIONE
		End Get

		Set(ByVal Value As String)
			_TITOLOOCCUPAZIONE = Value
		End Set
	End Property
	Public Property OCCUPANTE() As String
		Get
			Return _OCCUPANTE
		End Get

		Set(ByVal Value As String)
			_OCCUPANTE = Value
		End Set
	End Property
	Public Property DATAINIZIO() As String
		Get
			Return _DATAINIZIO
		End Get

		Set(ByVal Value As String)
			_DATAINIZIO = Value
		End Set
	End Property
	Public Property DATAFINE() As String
		Get
			Return _DATAFINE
		End Get

		Set(ByVal Value As String)
			_DATAFINE = Value
		End Set
	End Property
	Public Property DESTINAZIONEUSO() As String
		Get
			Return _DESTINAZIONEUSO
		End Get

		Set(ByVal Value As String)
			_DESTINAZIONEUSO = Value
		End Set
	End Property
	Public Property ABBINATO() As String
		Get
			Return _ABBINATO
		End Get

		Set(ByVal Value As String)
			_ABBINATO = Value
		End Set
	End Property
	Public Property TIPOUNITA() As String
		Get
			Return _TIPOUNITA
		End Get

		Set(ByVal Value As String)
			_TIPOUNITA = Value
		End Set
	End Property
	Public Property SEZIONE() As String
		Get
			Return _SEZIONE
		End Get

		Set(ByVal Value As String)
			_SEZIONE = Value
		End Set
	End Property
	Public Property FOGLIO() As String
		Get
			Return _FOGLIO

		End Get

		Set(ByVal Value As String)
			_FOGLIO = Value
		End Set
	End Property
	Public Property NUMERO() As String
		Get
			Return _NUMERO
		End Get

		Set(ByVal Value As String)
			_NUMERO = Value
		End Set
	End Property
	Public Property ESTPARTICELLA() As String
		Get
			Return _ESTPARTICELLA
		End Get

		Set(ByVal Value As String)
			_ESTPARTICELLA = Value
		End Set
	End Property
	Public Property TIPOPARTICELLA() As String
		Get
			Return _TIPOPARTICELLA
		End Get

		Set(ByVal Value As String)
			_TIPOPARTICELLA = Value
		End Set
	End Property
	Public Property SUBALTERNO() As String
		Get
			Return _SUBALTERNO
		End Get

		Set(ByVal Value As String)
			_SUBALTERNO = Value
		End Set
	End Property
	Public Property UBICAZIONE() As String
		Get
			Return _UBICAZIONE
		End Get

		Set(ByVal Value As String)
			_UBICAZIONE = Value
		End Set
	End Property
	Public Property CIVICO() As String
		Get
			Return _CIVICO
		End Get

		Set(ByVal Value As String)
			_CIVICO = Value
		End Set
	End Property
	Public Property INTERNO() As String
		Get
			Return _INTERNO
		End Get

		Set(ByVal Value As String)
			_INTERNO = Value
		End Set
	End Property
	Public Property SCALA() As String
		Get
			Return _SCALA
		End Get

		Set(ByVal Value As String)
			_SCALA = Value
		End Set
	End Property
	Public Property CADC() As String
		Get
			Return _CADC
		End Get

		Set(ByVal Value As String)
			_CADC = Value
		End Set
	End Property
End Class