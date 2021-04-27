Imports log4net

Public Class ClsInterDB
    Private Shared Log As ILog = LogManager.GetLogger("ClsInterDB")
    Private mySQLCommand As SqlClient.SqlCommand
    Private myResult As Integer = -1
    Private FncReplace As General

    '#Region "Query di SELEZIONE"
    '    Public Function GetFlussiCodIstat(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String) As DataView
    '        Dim dvResult As DataView
    '        Try
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "SELECT AE_FLUSSI_ESTRATTI.ID, AE_FLUSSI_ESTRATTI.CODICE_ISTAT, AE_FLUSSI_ESTRATTI.ANNO, AE_FLUSSI_ESTRATTI.NOME_FILE, AE_FLUSSI_ESTRATTI.DATA_ESTRAZIONE, AE_FLUSSI_ESTRATTI.NUTENTI, AE_FLUSSI_ESTRATTI.NRECORD, AE_FLUSSI_ESTRATTI.NARTICOLI"
    '            mySQLCommand.CommandText += " FROM AE_FLUSSI_ESTRATTI"
    '            mySQLCommand.CommandText += " WHERE (NOT AE_FLUSSI_ESTRATTI.NUTENTI IS NULL) AND (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CodiceISTAT)"
    '            mySQLCommand.CommandText += " ORDER BY AE_FLUSSI_ESTRATTI.ANNO"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CodiceISTAT", SqlDbType.NVarChar)).Value = sCodIstat
    '            dvResult = oDBManager.GetDataView(mySQLCommand, "RESULT")

    '            Return dvResult
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::GetFlussiCodIstat::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::GetFlussiCodIstat::" & Err.Message)
    '            Return Nothing
    '        End Try
    '    End Function

    '    Public Function GetIdFlusso(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
    '        Dim drDati As SqlClient.SqlDataReader

    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "INSERT INTO AE_FLUSSI_ESTRATTI(CODICE_ISTAT, ANNO)"
    '            mySQLCommand.CommandText += " VALUES ('" + sCodIstat + "', '" + sAnno + "')"
    '            mySQLCommand.CommandText += " SELECT @@IDENTITY"
    '            'eseguo la query
    '            drDati = oDBManager.GetDataReader(mySQLCommand.CommandText)
    '            Do While drDati.Read
    '                myResult = drDati(0)
    '            Loop
    '            drDati.Close()
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::GetIdFlusso::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::GetIdFlusso::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function GetNUtenti(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
    '        Dim drDati As SqlClient.SqlDataReader

    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "SELECT COUNT(TMPCOUNT.COD_CONTRIBUENTE) AS NUTENTI"
    '            mySQLCommand.CommandText += " FROM ("
    '            mySQLCommand.CommandText += " SELECT DISTINCT AE_DATI_FILE.COD_CONTRIBUENTE "
    '            mySQLCommand.CommandText += " FROM AE_DATI_FILE "
    '            mySQLCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CODICEISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@ANNORIF)) AS TMPCOUNT"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
    '            drDati = oDBManager.GetDataReader(mySQLCommand)
    '            Do While drDati.Read
    '                If Not IsDBNull(drDati("nutenti")) Then
    '                    myResult += CInt(drDati("nutenti"))
    '                End If
    '            Loop
    '            drDati.Close()
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::GetNUtenti::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::GetNUtenti::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function GetNArticoli(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
    '        Dim drDati As SqlClient.SqlDataReader

    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "SELECT COUNT(TMPCOUNT.ID_RUOLO) AS NARTICOLI"
    '            mySQLCommand.CommandText += " FROM ("
    '            mySQLCommand.CommandText += " SELECT DISTINCT AE_DATI_FILE.ID_RUOLO "
    '            mySQLCommand.CommandText += " FROM AE_DATI_FILE "
    '            mySQLCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CODICEISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@ANNORIF)) AS TMPCOUNT"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
    '            drDati = oDBManager.GetDataReader(mySQLCommand)
    '            Do While drDati.Read
    '                If Not IsDBNull(drDati("narticoli")) Then
    '                    myResult += CInt(drDati("narticoli"))
    '                End If
    '            Loop
    '            drDati.Close()
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::GetNArticoli::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::GetNArticoli::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function GetDisposizione(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sTributo As String, ByVal sAnno As String) As DataView
    '        Dim dvResult As DataView
    '        Try
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "SELECT *"
    '            mySQLCommand.CommandText += " FROM AE_DATI_FILE"
    '            mySQLCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CODICEISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_DATI_FILE.COD_TRIBUTO=@CODTRIBUTO)"
    '            mySQLCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@ANNORIF)"
    '            mySQLCommand.CommandText += " ORDER BY AE_DATI_FILE.COD_CONTRIBUENTE, AE_DATI_FILE.ID_RUOLO"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
    '            dvResult = oDBManager.GetDataView(mySQLCommand, "RESULT")

    '            Return dvResult
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::GetDisposizione::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::GetDisposizione::" & Err.Message)
    '            Return Nothing
    '        End Try
    '    End Function
    '#End Region

    '#Region "Query di INSERIMENTO/UPDATE"
    '    Public Function SetFlusso(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String, ByVal nUtenti As Integer, ByVal nArticoli As Integer) As Integer
    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "UPDATE AE_FLUSSI_ESTRATTI"
    '            mySQLCommand.CommandText += " SET NUTENTI=@NUTENTI, NARTICOLI=@NARTICOLI"
    '            mySQLCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CODICEISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.ANNO=@ANNORIF)"
    '            mySQLCommand.CommandText = mySQLCommand.CommandText
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@NUTENTI", SqlDbType.Int)).Value = nUtenti
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@NARTICOLI", SqlDbType.Int)).Value = nArticoli
    '            oDBManager.ExecuteNonQuery(mySQLCommand)

    '            myResult = 1
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::SetFlusso::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::SetFlusso::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function SetFlusso(ByVal oDBManager As Utility.DBManager, ByVal sNomeFileTracciati As String, ByVal nRcFile As Integer) As Integer
    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "UPDATE AE_FLUSSI_ESTRATTI"
    '            mySQLCommand.CommandText += " SET NOME_FILE=@NOMEFILE, DATA_ESTRAZIONE=@DATAESTRAZIONE, NRECORD=@NRC"
    '            mySQLCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CodiceISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.ANNO=@AnnoRif)"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMEFILE", SqlDbType.NVarChar)).Value = sNomeFileTracciati
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAESTRAZIONE", SqlDbType.NVarChar)).Value = FncReplace.ReplaceDataForDB(Now.ToString)
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@NRC", SqlDbType.Int)).Value = nRcFile
    '            oDBManager.ExecuteNonQuery(mySQLCommand)

    '            myResult = 1
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::SetFlusso::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::SetFlusso::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function SetDisposizione(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String, ByVal nFlusso As Integer) As Integer
    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "INSERT INTO AE_DATI_FILE("
    '            mySQLCommand.CommandText += " ID_FLUSSO, CODICE_ISTAT, COD_TRIBUTO, ANNO,"
    '            mySQLCommand.CommandText += " COD_FISCALE_ENTE, COGNOME_ENTE, NOME_ENTE, SESSO_ENTE, DATA_NASCITA_ENTE, COMUNE_NASCITA_SEDE_ENTE, PV_NASCITA_SEDE_ENTE,"
    '            mySQLCommand.CommandText += " COD_CONTRIBUENTE, ID_RUOLO, COD_FISCALE,"
    '            mySQLCommand.CommandText += " COGNOME, NOME, COMUNE_SEDE, PV_SEDE,"
    '            mySQLCommand.CommandText += " COMUNE_UBICAZIONE, PV_UBICAZIONE, COD_COMUNE_UBICAZIONE_CATAST,"
    '            mySQLCommand.CommandText += " DATA_INIZIO, DATA_FINE,"
    '            mySQLCommand.CommandText += " ID_TITOLO_OCCUPAZIONE, ID_TIPO_OCCUPANTE, ID_DESTINAZIONE_USO, ID_TIPO_UNITA,"
    '            mySQLCommand.CommandText += " SEZIONE, FOGLIO, PARTICELLA, ESTENSIONE_PARTICELLA, ID_TIPO_PARTICELLA, SUBALTERNO,"
    '            mySQLCommand.CommandText += " INDIRIZZO, CIVICO, INTERNO, SCALA,"
    '            mySQLCommand.CommandText += " ID_ASSENZA_DATI_CATASTALI)"

    '            mySQLCommand.CommandText += " SELECT " & nFlusso & ","
    '            mySQLCommand.CommandText += " RUOLO_TARSU.CODICE_ISTAT, RUOLO_TARSU.COD_TRIBUTO, RUOLO_TARSU.ANNO,"
    '            mySQLCommand.CommandText += " ENTI_IN_LAVORAZIONE.COD_FISCALE, ENTI_IN_LAVORAZIONE.COGNOME, ENTI_IN_LAVORAZIONE.NOME, ENTI_IN_LAVORAZIONE.SESSO, ENTI_IN_LAVORAZIONE.COMUNE_NASCITA_SEDE, ENTI_IN_LAVORAZIONE.PV_NASCITA_SEDE,"
    '            mySQLCommand.CommandText += " RUOLO_TARSU.COD_CONTRIBUENTE, RUOLO_TARSU.ID_RUOLO, CASE WHEN ANAGRAFICA.PARTITA_IVA IS NULL OR ANAGRAFICA.PARTITA_IVA='' THEN ANAGRAFICA.COD_FISCALE ELSE ANAGRAFICA.PARTITA_IVA END,"
    '            mySQLCommand.CommandText += " ANAGRAFICA.COGNOME_DENOMINAZIONE, ANAGRAFICA.NOME, ANAGRAFICA.COMUNE_NASCITA, ANAGRAFICA.PROV_NASCITA,"
    '            mySQLCommand.CommandText += " ENTI_IN_LAVORAZIONE.DESCRIZIONE, ENTI_IN_LAVORAZIONE.PROVINCIA, ENTI_IN_LAVORAZIONE.COD_BELFIORE,"
    '            mySQLCommand.CommandText += " RUOLO_TARSU.DATA_INIZIO, RUOLO_TARSU.DATA_FINE,"
    '            mySQLCommand.CommandText += " RUOLO_TARSU.ID_TITOLO_OCCUPAZIONE, RUOLO_TARSU.ID_NATURA_OCCUPANTE, RUOLO_TARSU.ID_DESTINAZIONE_USO, RUOLO_TARSU.ID_TIPO_UNITA,"
    '            mySQLCommand.CommandText += " RUOLO_TARSU.SEZIONE, RUOLO_TARSU.FOGLIO, RUOLO_TARSU.PARTICELLA, RUOLO_TARSU.ESTENSIONE_PARTICELLA, RUOLO_TARSU.ID_TIPO_PARTICELLA, RUOLO_TARSU.SUBALTERNO,"
    '            mySQLCommand.CommandText += " RUOLO_TARSU.UBICAZIONE, RUOLO_TARSU.CIVICO, RUOLO_TARSU.INTERNO, RUOLO_TARSU.SCALA, 3"
    '            mySQLCommand.CommandText += " FROM ANAGRAFICA"
    '            mySQLCommand.CommandText += " INNER JOIN RUOLO_TARSU ON ANAGRAFICA.COD_CONTRIBUENTE = RUOLO_TARSU.COD_CONTRIBUENTE"
    '            mySQLCommand.CommandText += " INNER JOIN ENTI_IN_LAVORAZIONE ON RUOLO_TARSU.CODICE_ISTAT=ENTI_IN_LAVORAZIONE.CODICE_ISTAT"
    '            mySQLCommand.CommandText += " WHERE (ANAGRAFICA.DATA_FINE_VALIDITA IS NULL)"
    '            mySQLCommand.CommandText += " AND (RUOLO_TARSU.CODICE_ISTAT=@CODICEISTAT)"
    '            mySQLCommand.CommandText += " AND (RUOLO_TARSU.ANNO=@ANNORIF)"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
    '            myResult = oDBManager.ExecuteNonQuery(mySQLCommand)
    '            myResult = 1
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::SetDisposizione::" & Err.Message & vbCrLf & "SQL::" & mySQLCommand.CommandText)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::SetDisposizione::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function SetDisposizione(ByVal oDBManager As Utility.DBManager, ByVal oMyDati As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Integer
    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "INSERT INTO AE_DATI_FILE("
    '            mySQLCommand.CommandText += " ID_FLUSSO, CODICE_ISTAT, COD_TRIBUTO, ANNO,"
    '            mySQLCommand.CommandText += " COD_FISCALE_ENTE, COGNOME_ENTE, NOME_ENTE, SESSO_ENTE, DATA_NASCITA_ENTE, COMUNE_NASCITA_SEDE_ENTE, PV_NASCITA_SEDE_ENTE,"
    '            mySQLCommand.CommandText += " COD_CONTRIBUENTE, ID_RUOLO, COD_FISCALE,"
    '            mySQLCommand.CommandText += " COGNOME, NOME, SESSO, DATA_NASCITA, COMUNE_SEDE, PV_SEDE,"
    '            mySQLCommand.CommandText += " COMUNE_DOMICILIOFISC, PV_DOMICILIOFISC,"
    '            mySQLCommand.CommandText += " COMUNE_UBICAZIONE, PV_UBICAZIONE, COMUNE_CATAST_UBICAZIONE, COD_COMUNE_UBICAZIONE_CATAST,"
    '            mySQLCommand.CommandText += " ESTREMI_CONTRATTO, DATA_INIZIO, DATA_FINE,"
    '            mySQLCommand.CommandText += " ID_TITOLO_OCCUPAZIONE, ID_TIPO_OCCUPANTE, ID_TIPO_UTENZA, ID_DESTINAZIONE_USO, ID_TIPO_UNITA,"
    '            mySQLCommand.CommandText += " SEZIONE, FOGLIO, PARTICELLA, ESTENSIONE_PARTICELLA, ID_TIPO_PARTICELLA, SUBALTERNO,"
    '            mySQLCommand.CommandText += " INDIRIZZO, CIVICO, INTERNO, SCALA,"
    '            mySQLCommand.CommandText += " ID_ASSENZA_DATI_CATASTALI, MESI_FATTURAZIONE, SEGNO_SPESA, SPESA_CONSUMO)"

    '            mySQLCommand.CommandText += " VALUES (" & oMyDati.nIDFlusso & "," & oMyDati.sCodISTAT & "," & oMyDati.sTributo & "," & oMyDati.sAnno & ","
    '            mySQLCommand.CommandText += oMyDati.sCodFiscaleEnte & "," & oMyDati.sCognomeEnte & "," & oMyDati.sNomeEnte & "," & oMyDati.sSessoEnte & "," & oMyDati.sDataNascitaEnte & "," & oMyDati.sComuneNascitaSedeEnte & "," & oMyDati.sPVNascitaSedeEnte & ","
    '            mySQLCommand.CommandText += oMyDati.nIDContribuente & "," & oMyDati.nIDCollegamento & "," & oMyDati.sCodFiscale & ","
    '            mySQLCommand.CommandText += oMyDati.sCognome & "," & oMyDati.sNome & "," & oMyDati.sSesso & "," & oMyDati.sDataNascita & "," & oMyDati.sComuneNascitaSede & "," & oMyDati.sPVNascitaSede & ","
    '            mySQLCommand.CommandText += oMyDati.sComuneDomFisc & "," & oMyDati.sPVDomFisc & ","
    '            mySQLCommand.CommandText += oMyDati.sComuneAmmUbicazione & "," & oMyDati.sPVAmmUbicazione & "," & oMyDati.sComuneCatastUbicazione & "," & oMyDati.sCodComuneUbicazioneCatast & ","
    '            mySQLCommand.CommandText += oMyDati.sEstremiContratto & "," & oMyDati.sDataInizio & "," & oMyDati.sDataFine & ","
    '            mySQLCommand.CommandText += oMyDati.nIDTitoloOccupazione & "," & oMyDati.nIDTipoOccupante & "," & oMyDati.nIDTipoUtenza & "," & oMyDati.nIDDestinazioneUso & "," & oMyDati.sIDTipoUnita & ","
    '            mySQLCommand.CommandText += oMyDati.sSezione & "," & oMyDati.sFoglio & "," & oMyDati.sParticella & "," & oMyDati.sEstensioneParticella & "," & oMyDati.sIDTipoParticella & "," & oMyDati.sSubalterno & ","
    '            mySQLCommand.CommandText += oMyDati.sIndirizzo & "," & oMyDati.sCivico & "," & oMyDati.sInterno & "," & oMyDati.sScala & ","
    '            mySQLCommand.CommandText += oMyDati.nIDAssenzaDatiCatastali & "," & oMyDati.nMesiFatturazione & "," & oMyDati.sSegno & "," & oMyDati.nSpesaConsumo & ")"

    '            myResult = oDBManager.ExecuteNonQuery(mySQLCommand)
    '            myResult = 1
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::SetDisposizione::" & Err.Message & vbCrLf & "SQL::" & mySQLCommand.CommandText)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::SetDisposizione::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function
    '#End Region

    '#Region "Query di CANCELLAZIONE"
    '    Public Function DeleteDisposizione(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "DELETE"
    '            mySQLCommand.CommandText += " FROM AE_DATI_FILE"
    '            mySQLCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CodiceISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@AnnoRif)"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CodiceISTAT", sCodIstat))
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@AnnoRif", sAnno))
    '            myResult = oDBManager.ExecuteNonQuery(mySQLCommand)
    '            myResult = 1
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::DeleteDisposizione::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::DeleteDisposizione::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function

    '    Public Function DeleteFlusso(ByVal oDBManager As Utility.DBManager, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
    '        Try
    '            myResult = -1
    '            mySQLCommand = New SqlClient.SqlCommand
    '            mySQLCommand.CommandText = "DELETE"
    '            mySQLCommand.CommandText += " FROM AE_FLUSSI_ESTRATTI"
    '            mySQLCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CodiceISTAT)"
    '            mySQLCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.ANNO=@AnnoRif)"
    '            'valorizzo i parameters
    '            mySQLCommand.Parameters.Clear()
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@CodiceISTAT", sCodIstat))
    '            mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@AnnoRif", sAnno))
    '            myResult = oDBManager.ExecuteNonQuery(mySQLCommand)
    '            myResult = 1
    '        Catch Err As Exception
    '            Log.Debug("Si è verificato un errore in ServiceOPENae::DeleteFlusso::" & Err.Message)
    '            Log.Warn("Si è verificato un errore in ServiceOPENae::DeleteFlusso::" & Err.Message)
    '        End Try
    '        Return myResult
    '    End Function
    '#End Region
End Class
