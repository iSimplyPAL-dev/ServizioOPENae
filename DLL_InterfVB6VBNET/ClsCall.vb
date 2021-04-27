Imports System
Imports System.Runtime.InteropServices
Imports System.Data.OleDb

<Serializable()> _
Public Class ClsCall
#Region "RDB"
    Public Function CallServiceAE(ByVal StrConn As String, ByVal PathLog As String, ByVal FileLog As String, ByVal PathFile As String) As String
        'Dim FileLOG As String

        Try
            Dim AppReader As New System.Configuration.AppSettingsReader
            Dim CnEnte As New GetConnectionDB

            FileLog = GetLog(PathLog, FileLog)

            Dim DrElementiAE As OleDbDataReader
            Dim NonInserisci As Boolean
            Dim TotElementi As Integer

            Dim list() As ServiceOPENae.DisposizioneAE
            Dim osingle As ServiceOPENae.DisposizioneAE

            Dim SQLString As String

            Dim x As Integer

            Dim AnnoRif, CodISTAT As String

            'LEGGO IL DB DI ACCESS E ASSEGNO I VALORI

            Dim ControlloErrori As String = ""
            SQLString = "SELECT [_AE_ESTRAZIONI].*, [_AE_DATI_FLUSSO].*,[_AE_ESTRAZIONI].PROG_ESTRAZIONE as pe"
            SQLString = SQLString & " FROM _AE_DATI_FLUSSO INNER JOIN _AE_ESTRAZIONI ON [_AE_DATI_FLUSSO].PROG_ESTRAZIONE = [_AE_ESTRAZIONI].PROG_ESTRAZIONE"
            SQLString = SQLString & " WHERE([_AE_ESTRAZIONI].DATA_ESTRAZIONE Is Null)"
            DrElementiAE = CnEnte.GetDataReaderACCESS(SQLString, StrConn, ControlloErrori)
            If ControlloErrori <> "" Then
                WriteLOG(FileLog, ControlloErrori)
                Return "Si sono verificati errori nell'estrazione. Controllare il file di log al percorso " & FileLog
            End If

            'If DrElementiAE.Read = True Then
            'ho degli elementi da estrarre
            x = 0
            'For x = 0 To TotElementi
            Do While DrElementiAE.Read
                AnnoRif = DrElementiAE.Item("ANNO_RUOLO")
                CodISTAT = DrElementiAE.Item("CODISTAT")

                'inizializzo il nuovo oggetto
                osingle = New ServiceOPENae.DisposizioneAE
                'assegno i valori della testata
                ControlloErrori = RecordTesta(DrElementiAE, osingle)
                If ControlloErrori <> "" Then
                    WriteLOG(FileLog, ControlloErrori)
                    Return "Si sono verificati errori nell'estrazione. Controllare il file di log al percorso " & FileLog
                End If

                'assegno i valori di dettaglio
                osingle.sCodISTAT = CType(CodISTAT, String).Trim
                osingle.sTributo = "0434"
                ControlloErrori = RecordDettaglio(DrElementiAE, osingle)
                If ControlloErrori <> "" Then
                    WriteLOG(FileLog, ControlloErrori)
                    Return "Si sono verificati errori nell'estrazione. Controllare il file di log al percorso " & FileLog
                End If

                'assegno i valori della testata
                ControlloErrori = RecordCoda(DrElementiAE, osingle)
                If ControlloErrori <> "" Then
                    WriteLOG(FileLog, ControlloErrori)
                    Return "Si sono verificati errori nell'estrazione. Controllare il file di log al percorso " & FileLog
                End If

                'DrElementiAE.NextResult()

                'ridimensiono l'array di oggetti
                ReDim Preserve list(x)
                'assegno l'oggetto all'array
                list(x) = osingle
                x = x + 1

            Loop

            If x = 0 Then
                Return "Non sono stati trovati elementi per l'estrazione"
            End If
            DrElementiAE.Close()


            'RICHIAMO IL METODO PER CREARE IL TRACCIATO
            Dim ObjTracciato As New ServiceOPENae.ServiceOPENae
            'setto l'url del servizio
            ObjTracciato.Url = System.Configuration.ConfigurationSettings.AppSettings("OPENae_InterfVB6VBNET.ServiceOPENae.ServiceOPENae")
            ObjTracciato.Timeout = 3600000

            'POPOLO LA TABELLA DI APPOGGIO
            If ObjTracciato.PopolaTabAppoggioAE(list) = False Then
                WriteLOG(FileLog, "Errore nel popolamento tabella appoggio WebService")
                Return "Si sono verificati errori nell'estrazione. Controllare il file di log " & FileLog
            End If

            Dim FileAgenziaEntrate As String

            Dim PathNameFileAE As String
            PathNameFileAE = ObjTracciato.EstraiTracciatoAE("0434", AnnoRif, CodISTAT, FileAgenziaEntrate)
            If PathNameFileAE = "" Then
                WriteLOG(FileLog, "Errore nella creazione del tracciato da parte del WebService")
                Return "Si sono verificati errori nell'estrazione. Controllare il file di log " & FileLog
            Else
                FileAgenziaEntrate = PathFile + FileAgenziaEntrate
                'RICHIAMO IL METODO PER SCARICARE IL TRACCIATO
                Dim objW As New Net.WebClient
                objW.DownloadFile(PathNameFileAE, FileAgenziaEntrate)
            End If

            Return ""

        Catch ex As Exception
            Dim err As String

            err = ex.Message
            WriteLOG(FileLog, ex.Message)
            Return "Si sono verificati errori nell'estrazione. Controllare il file di log " & FileLog
        End Try

    End Function

    Public Sub WriteLOG(ByVal glbFileLog As String, ByVal TextLOG As String)
        Dim DataOra As String

        FileOpen(1, glbFileLog, OpenMode.Append)
        PrintLine(1, "****************************************")
        DataOra = "Data e Ora: " & Now.Now
        PrintLine(1, DataOra & " Operazione: " & TextLOG)
        FileClose(1)

        'Dim StreamWriter As IO.StreamWriter = IO.File.AppendText(glbFileLog)

        'Try
        '    If TextLOG <> "" Then
        '        StreamWriter.WriteLine("****************************************")
        '        StreamWriter.WriteLine("Data e Ora: " & " Operazione: " & TextLOG)
        '    End If
        '    StreamWriter.Flush()
        '    StreamWriter.Close()
        'Catch err As Exception

        'End Try

    End Sub

    Public Function GetLog(ByVal PercorsoLOG As String, ByVal NomeFileLOG As String) As String
        'Dim AppReader As New System.Configuration.AppSettingsReader

        Dim PathLog As String

        PathLog = PercorsoLOG
        PathLog = PathLog & CType(Now.Today, String).Replace("/", "") & "_"
        PathLog = PathLog & NomeFileLOG

        'PathLog = CType(AppReader.GetValue("PercorsoLOG", GetType(String)), String)
        'PathLog = PathLog & CType(Now.Date, String) & "_"
        'PathLog = PathLog & CType(AppReader.GetValue("NomeFileLOG", GetType(String)), String)

        Return PathLog
    End Function

    Public Function RecordTesta(ByVal DrElemento As OleDbDataReader, ByRef ObjAE As ServiceOPENae.DisposizioneAE) As String

        Try
            If IsDBNull(DrElemento.Item("CODFISCALE")) = False Then
                ObjAE.sCodFiscaleEnte = CType(DrElemento.Item("CODFISCALE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("DENOMINAZIONECOM")) = False Then
                ObjAE.sCognomeEnte = CType(DrElemento.Item("DENOMINAZIONECOM"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("COMUNELEGALE")) = False Then
                ObjAE.sComuneNascitaSedeEnte = CType(DrElemento.Item("COMUNELEGALE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("PROVLEG")) = False Then
                ObjAE.sPVNascitaSedeEnte = CType(DrElemento.Item("PROVLEG"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("ANNO_RUOLO")) = False Then
                ObjAE.sAnno = CType(DrElemento.Item("ANNO_RUOLO"), String).Trim
            End If

            Return ""
        Catch ex As Exception
            Return "Creazione Record di Testa: " & ex.Message
        End Try

    End Function

    Public Function RecordDettaglio(ByVal DrElemento As OleDbDataReader, ByRef ObjAE As ServiceOPENae.DisposizioneAE) As String

        Try
            Dim collegamento As String
            collegamento = CType(DrElemento.Item("PROGRESSIVO"), String)
            collegamento = collegamento & CType(DrElemento.Item("pe"), String).PadLeft(5, "0")
            ObjAE.nIDCollegamento = CType(collegamento, String)

            If IsDBNull(DrElemento.Item("COD_CONTRIB")) = False Then
                ObjAE.nIDContribuente = CType(DrElemento.Item("COD_CONTRIB"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("CFPIVA")) = False Then
                ObjAE.sCodFiscale = CType(DrElemento.Item("CFPIVA"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("COGNOME")) = False Then
                ObjAE.sCognome = CType(DrElemento.Item("COGNOME"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("NOME")) = False Then
                ObjAE.sNome = CType(DrElemento.Item("NOME"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("DENOMINAZIONE")) = False Then
                ObjAE.sCognome = CType(DrElemento.Item("DENOMINAZIONE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("COMUNE")) = False Then
                ObjAE.sComuneNascitaSede = CType(DrElemento.Item("COMUNE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("PROVINCIA")) = False Then
                ObjAE.sPVNascitaSede = CType(DrElemento.Item("PROVINCIA"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("TITOLO_OCCUPAZIONE")) = False Then
                ObjAE.nIDTitoloOccupazione = CType(DrElemento.Item("TITOLO_OCCUPAZIONE"), Int16)
            End If

            If IsDBNull(DrElemento.Item("OCCUPANTE")) = False Then
                ObjAE.nIDTipoOccupante = CType(DrElemento.Item("OCCUPANTE"), Int16)
            End If

            If IsDBNull(DrElemento.Item("DI_OCCUPAZIONE")) = False Then
                Dim DataIO As String
                DataIO = CType(DrElemento.Item("DI_OCCUPAZIONE"), String).Trim
                DataIO = DataIO.Replace("/", "")
                DataIO = DataIO.Substring(4, 4) & DataIO.Substring(2, 2) & DataIO.Substring(0, 2)
                ObjAE.sDataInizio = DataIO
            End If

            If IsDBNull(DrElemento.Item("DF_OCCUPAZIONE")) = False Then
                Dim DataFO As String
                DataFO = CType(DrElemento.Item("DF_OCCUPAZIONE"), String).Trim
                If DataFO <> "" Then
                    DataFO = DataFO.Replace("/", "")
                    DataFO = DataFO.Substring(4, 4) & DataFO.Substring(2, 2) & DataFO.Substring(0, 2)
                End If
                ObjAE.sDataFine = DataFO
            End If

            If IsDBNull(DrElemento.Item("DESTINAZIONE_USO")) = False Then
                ObjAE.nIDDestinazioneUso = CType(DrElemento.Item("DESTINAZIONE_USO"), Int16)
            End If

            If IsDBNull(DrElemento.Item("COMUNEAMM")) = False Then
                ObjAE.sComuneAmmUbicazione = CType(DrElemento.Item("COMUNEAMM"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("PROV")) = False Then
                ObjAE.sPVAmmUbicazione = CType(DrElemento.Item("PROV"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("COMUNECAT")) = False Then
                ObjAE.sComuneCatastUbicazione = CType(DrElemento.Item("COMUNECAT"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("CODICE")) = False Then
                ObjAE.sCodComuneUbicazioneCatast = CType(DrElemento.Item("CODICE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("TIPO_UNITA")) = False Then
                ObjAE.sIDTipoUnita = CType(DrElemento.Item("TIPO_UNITA"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("SEZIONE")) = False Then
                ObjAE.sSezione = CType(DrElemento.Item("SEZIONE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("FOGLIO")) = False Then
                ObjAE.sFoglio = CType(DrElemento.Item("FOGLIO"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("PARTICELLA")) = False Then
                ObjAE.sParticella = CType(DrElemento.Item("PARTICELLA"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("SUBALTERNO")) = False Then
                ObjAE.sSubalterno = CType(DrElemento.Item("SUBALTERNO"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("UBICAZIONE")) = False Then
                ObjAE.sIndirizzo = CType(DrElemento.Item("UBICAZIONE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("CIVICO")) = False Then
                ObjAE.sCivico = CType(DrElemento.Item("CIVICO"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("CADC")) = False Then
                ObjAE.nIDAssenzaDatiCatastali = CType(DrElemento.Item("CADC"), Int16)
            End If

            Return ""

        Catch ex As Exception
            Return "Creazione Record di Dettaglio: " & ex.Message
        End Try

    End Function

    Public Function RecordCoda(ByVal DrElemento As OleDbDataReader, ByRef ObjAE As ServiceOPENae.DisposizioneAE) As String

        Try
            If IsDBNull(DrElemento.Item("CODFISCALE")) = False Then
                ObjAE.sCodFiscaleEnte = CType(DrElemento.Item("CODFISCALE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("DENOMINAZIONECOM")) = False Then
                ObjAE.sCognomeEnte = CType(DrElemento.Item("DENOMINAZIONECOM"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("COMUNELEGALE")) = False Then
                ObjAE.sComuneNascitaSedeEnte = CType(DrElemento.Item("COMUNELEGALE"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("PROVLEG")) = False Then
                ObjAE.sPVNascitaSedeEnte = CType(DrElemento.Item("PROVLEG"), String).Trim
            End If

            If IsDBNull(DrElemento.Item("ANNO_RUOLO")) = False Then
                ObjAE.sAnno = CType(DrElemento.Item("ANNO_RUOLO"), String).Trim
            End If

            Return ""
        Catch ex As Exception
            Return "Creazione Record di Coda: " & ex.Message
        End Try

    End Function
#End Region

#Region "OPENrendicontazione"
    Public Function CreaFlussoMEF(ByVal sCodiceISTAT As String, ByVal sCodBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByVal sFileImport As String, ByVal sFileToImport As String, ByVal MyFileLog As String, ByVal sProvenienza As String, ByRef sPathFileDownload As String) As String
        Try
            WriteLOG(MyFileLog, "CreaFlussoMEF::1")
            Dim AppReader As New System.Configuration.AppSettingsReader
            WriteLOG(MyFileLog, "CreaFlussoMEF::2")
            Dim FncAE As New ServiceOPENae.ServiceOPENae
            WriteLOG(MyFileLog, "CreaFlussoMEF::3")
            Dim sFileExport, sNameExport As String
            WriteLOG(MyFileLog, "CreaFlussoMEF::4")
            Dim objW As New Net.WebClient
            WriteLOG(MyFileLog, "CreaFlussoMEF::5")

            WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::inizio procedura")
            WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::richiamo il popolamento della tabella d'appoggio")
            FncAE.Timeout = 360000000
            'setto l'url del servizio
            FncAE.Url = System.Configuration.ConfigurationSettings.AppSettings("OPENae_InterfVB6VBNET.ServiceOPENae.ServiceOPENae")
            WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::devo uplodare il file::" & sFileToImport & "::al percorso:" & sFileImport)
            objW.UploadFile(sFileToImport, "PUT", sFileImport)
            WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::upload fatto")
            'richiamo l'acquisizione del file
			If FncAE.PopolaTabAppoggioAE(sTributo, sAnnoRif, sCodiceISTAT, sFileImport, sProvenienza) = False Then
				Return ""
			End If
			WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::richiamo l'estrazione del file")
			'richiamo l'estrazione del file
			sFileExport = FncAE.EstraiTracciatoAE(sTributo, sAnnoRif, sCodiceISTAT, sCodBelfiore, sDescrEnte, sCAPEnte, sDataScadenza, nProgInvio, sNameExport)
			WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::fine procedura")
			WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::se provengo da RDB faccio il download del file")
			If sFileExport = "" Then
				WriteLOG(MyFileLog, "Errore nella creazione del tracciato da parte del WebService")
				Return ""
			Else
				sPathFileDownload = sPathFileDownload + sNameExport + ".txt"
				'RICHIAMO IL METODO PER SCARICARE IL TRACCIATO
				WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::devo fare il download da:" & sFileExport & " a:" & sPathFileDownload)
				objW.DownloadFile(sFileExport, sPathFileDownload)
				WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::download fatto")
			End If
			Return sPathFileDownload
		Catch Err As Exception
            WriteLOG(MyFileLog, "ServiceAE::ClsCall::CreaFlussoMEF::si è verificato il seguente errore -> " & Err.Message)
            Return ""
        End Try
    End Function
#End Region
End Class
