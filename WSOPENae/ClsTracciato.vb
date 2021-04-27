Imports log4net
'Imports Utility

Public Class ClsTracciato
    Private Shared Log As ILog = LogManager.GetLogger("ClsTracciato")
    Private AppReader As New System.Configuration.AppSettingsReader
    Private oConst As New General
    Private oQueryManager As ClsInterDB
    Private sMsgErr, sDati As String
    Private Structure StructureTestaCoda
        Dim sCFPIVA As String
        Dim sCognome As String
        Dim sNome As String
        Dim sSesso As String
        Dim sDataNascita As String
        Dim sComuneNascita As String
        Dim sPvNascita As String
        Dim sRagSociale As String
        Dim sComuneSede As String
        Dim sPvSede As String
        Dim sAnno As String
        Dim sProgFornitura As String
        Dim sDataInvio As String
        Dim sFiller As String
    End Structure
    Private Structure StructureDettaglio
        Dim sCFPIVA As String
        Dim sCognome As String
        Dim sNome As String
        Dim sSesso As String
        Dim sDataNascita As String
        Dim sComuneNascita As String
        Dim sPVNascita As String
        Dim sRagSociale As String
        Dim sComuneSede As String
        Dim sPvSede As String
        Dim sComuneDomFiscale As String
        Dim sPVDomFiscale As String
        Dim sIdTitoloOccupazione As String
        Dim sIdNaturaOccupante As String
        Dim sDataInizio As String
        Dim sDataFine As String
        Dim sEstremiContratto As String
        Dim sIdTipologiaUtenza As String
        Dim sIdDestinazioneUso As String
        Dim sComuneAmmUbicazione As String
        Dim sPVUbicazione As String
        Dim sComuneCatUbicazione As String
        Dim sCodComuneUbicazCatast As String
        Dim sIdTipoUnita As String
        Dim sSezione As String
        Dim sFoglio As String
        Dim sParticella As String
        Dim sEstensioneParticella As String
        Dim sIdTipoParticella As String
        Dim sSubalterno As String
        Dim sIndirizzo As String
        Dim sCivico As String
        Dim sInterno As String
        Dim sScala As String
        Dim sIdAssenzaDatiCatastali As String
        Dim sMesiFatturazione As String
        Dim sSegnoSpesa As String
        Dim sSpesaConsumo As String
        Dim sFiller As String
    End Structure

    Public Function EstraiTracciato(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByRef sNomeFileTracciati As String) As String
        'Dim cnn As New DBConnection
        'Dim oDBManager As DBManager
        'Dim dvMyDati As DataView
        'Dim sPathCartellaTracciatiCompleto As String
        'Dim nRcFile As Integer = 0

        'Try
        '    Log.Debug("ServiceOPENae::EstraiTracciato::inizio procedura")
        '    'il nome del tracciato è formattato nel modo seguente:
        '    '-Codice ISTAT
        '    '-Anno
        '    '-RUOLO
        '    '-AGENZIA_ENTRATE
        '    '-AAAAMMGGHHMMSS
        '    sNomeFileTracciati = sCodiceISTAT + "_" + sAnnoRif + "_RUOLO_AGENZIA_ENTRATE" + DateTime.Now.ToString("yyyyMMdd_hhmmss")
        '    sPathCartellaTracciatiCompleto = AppReader.GetValue("PathTracciati", GetType(String)) + AppReader.GetValue("NomeCartellaTracciati", GetType(String)) + sNomeFileTracciati + ".txt"

        '    oDBManager = cnn.DBConnection()

        '    dvMyDati = oQueryManager.GetDisposizione(oDBManager, sCodiceISTAT, sTributo, sAnnoRif)
        '    If dvMyDati.Count > 0 Then
        '        Select Case sTributo
        '            Case "0465", "0434"
        '                If TARSU_EstraiTracciato(dvMyDati, sPathCartellaTracciatiCompleto, nRcFile) = False Then
        '                    Return ""
        '                End If
        '            Case "9000"
        '                If H2O_EstraiTracciato(dvMyDati, sPathCartellaTracciatiCompleto, nRcFile) = False Then
        '                    Return ""
        '                End If
        '        End Select
        '    End If
        '    dvMyDati.Dispose()
        '    '*********************************************************************
        '    'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
        '    '*********************************************************************
        '    If oQueryManager.SetFlusso(oDBManager, sNomeFileTracciati, nRcFile) = -1 Then
        '        Return ""
        '    End If
        '    Log.Debug("ServiceOPENae::EstraiTracciato::fine procedura")
        '    Return sPathCartellaTracciatiCompleto
        'Catch Err As Exception
        '    Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::" & Err.Message)
        '    Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::" & Err.Message)
        '    Return ""
        'End Try
    End Function

    Private Function WriteFile(ByVal sFile As String, ByVal DatiToPrint As String, ByVal ErrWriteFile As String) As Integer
        Dim MyFileToWrite As IO.StreamWriter = IO.File.AppendText(sFile)
        Dim sDatiFile As String = ""

        Try
            sDatiFile = DatiToPrint

            MyFileToWrite.WriteLine(sDatiFile)
            MyFileToWrite.Flush()

            Return 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::WriteFile::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::WriteFile::" & Err.Message)
            ErrWriteFile = Err.Message
            Return 0
        Finally
            MyFileToWrite.Close()
        End Try
    End Function

#Region "TARSU"
    Private Function TARSU_EstraiTracciato(ByVal dvMyDati As DataView, ByVal sPathCartellaTracciatiCompleto As String, ByRef nRcFile As Integer) As Boolean
        Dim x As Integer

        Try
            '***************************************************
            'scrivo il record di testa
            '***************************************************
            If TARSU_WriteRcTesta(dvMyDati, sPathCartellaTracciatiCompleto, nRcFile) = False Then
                Return False
            End If
            For x = 0 To dvMyDati.Count
                '***************************************************
                'record di dettaglio
                '***************************************************
                If TARSU_WriteRcDettaglio(dvMyDati, x, sPathCartellaTracciatiCompleto, nRcFile) = False Then
                    Return False
                End If
            Next
            '***************************************************
            'scrivo il record di coda
            '***************************************************
            If TARSU_WriteRcTesta(Nothing, sPathCartellaTracciatiCompleto, -1) = False Then
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_EstraiTracciato::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_EstraiTracciato::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function TARSU_WriteRcTesta(ByVal dvTesta As DataView, ByVal sPathCartellaTracciatiCompleto As String, ByRef nRcFile As Integer) As Boolean
        Dim oTestaCoda As New StructureTestaCoda

        Try
            'Codice Fiscale - Partita IVA
            oTestaCoda.sCFPIVA = CStr(dvTesta.Item(0)("cod_fiscale_ente"))
            If IsNumeric(dvTesta.Item(0)("cod_fiscale_ente")) Then
                '***persona giuridica***
                'Cognome
                oTestaCoda.sCognome = ""
                'Nome
                oTestaCoda.sNome = ""
                'sesso
                oTestaCoda.sSesso = ""
                'data nascita
                oTestaCoda.sDataNascita = ""
                'comune nascita
                oTestaCoda.sComuneNascita = ""
                'provincia nascita
                oTestaCoda.sPvNascita = ""
                'Denominazione o Ragion Sociale
                If Not IsDBNull(dvTesta.Item(0)("cognome_ente")) Then
                    oTestaCoda.sRagSociale = CStr(dvTesta.Item(0)("cognome_ente"))
                    If Not IsDBNull(dvTesta.Item(0)("nome_ente")) Then
                        oTestaCoda.sRagSociale += " " & CStr(dvTesta.Item(0)("nome_ente"))
                    End If
                Else
                    oTestaCoda.sRagSociale = ""
                End If
                'Comune Sede
                If Not IsDBNull(dvTesta.Item(0)("comune_nascita_sede_ente")) Then
                    oTestaCoda.sComuneSede = CStr(dvTesta.Item(0)("comune_nascita_sede_ente"))
                Else
                    oTestaCoda.sComuneSede = ""
                End If
                'Provincia Sede
                If Not IsDBNull(dvTesta.Item(0)("pv_nascita_sede_ente")) Then
                    oTestaCoda.sPvSede = CStr(dvTesta.Item(0)("pv_nascita_sede_ente"))
                Else
                    oTestaCoda.sPvSede = ""
                End If
            Else
                '***persona fisica***
                'Cognome
                If Not IsDBNull(dvTesta.Item(0)("cognome_ente")) Then
                    oTestaCoda.sCognome = CStr(dvTesta.Item(0)("cognome_ente"))
                Else
                    oTestaCoda.sCognome = ""
                End If
                'Nome
                If Not IsDBNull(dvTesta.Item(0)("nome_ente")) Then
                    oTestaCoda.sNome = CStr(dvTesta.Item(0)("nome_ente"))
                Else
                    oTestaCoda.sNome = ""
                End If
                'sesso
                If Not IsDBNull(dvTesta.Item(0)("sesso_ente")) Then
                    oTestaCoda.sSesso = CStr(dvTesta.Item(0)("sesso_ente"))
                Else
                    oTestaCoda.sSesso = ""
                End If
                'data nascita
                If Not IsDBNull(dvTesta.Item(0)("data_nascita_ente")) Then
                    oTestaCoda.sDataNascita = oConst.ReplaceDataForTXT(CStr(dvTesta.Item(0)("data_nascita_ente")))
                Else
                    oTestaCoda.sDataNascita = ""
                End If
                'comune nascita
                If Not IsDBNull(dvTesta.Item(0)("comune_nascita_sede_ente")) Then
                    oTestaCoda.sComuneNascita = CStr(dvTesta.Item(0)("comune_nascita_sede_ente"))
                Else
                    oTestaCoda.sComuneNascita = ""
                End If
                'provincia nascita
                If Not IsDBNull(dvTesta.Item(0)("pv_nascita_sede_ente")) Then
                    oTestaCoda.sPvNascita = CStr(dvTesta.Item(0)("pv_nascita_sede_ente"))
                Else
                    oTestaCoda.sPvNascita = ""
                End If
                'Denominazione o Ragion Sociale
                oTestaCoda.sRagSociale = ""
                'Comune Sede
                oTestaCoda.sComuneSede = ""
                'Provincia Sede
                oTestaCoda.sPvSede = ""
            End If
            'anno
            oTestaCoda.sAnno = CStr(dvTesta.Item(0)("anno"))
            'filler
            oTestaCoda.sFiller = ""
            sDati = TARSU_FormattaRcTesta(oTestaCoda)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
                nRcFile += 1
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcTesta::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcTesta::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function TARSU_WriteRcCoda(ByVal oCoda As StructureTestaCoda, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Try
            sDati = TARSU_FormattaRcCoda(oCoda)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcCoda::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcCoda::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function TARSU_WriteRcDettaglio(ByVal dvDettaglio As DataView, ByVal nDett As Integer, ByVal sPathCartellaTracciatiCompleto As String, ByRef nRcFile As Integer) As Boolean
        Dim oDettaglio As New StructureDettaglio

        Try
            'Codice Fiscale - Partita IVA
            oDettaglio.sCFPIVA = CStr(dvDettaglio.Item(nDett)("cod_fiscale"))
            If IsNumeric(dvDettaglio.Item(nDett)("cod_fiscale")) Then
                'Cognome
                oDettaglio.sCognome = ""
                'Nome
                oDettaglio.sNome = ""
                'Denominazione o Ragion Sociale
                If Not IsDBNull(dvDettaglio.Item(nDett)("cognome")) Then
                    oDettaglio.sRagSociale = CStr(dvDettaglio.Item(nDett)("cognome"))
                    If Not IsDBNull(dvDettaglio.Item(nDett)("nome")) Then
                        oDettaglio.sRagSociale += " " & CStr(dvDettaglio.Item(nDett)("nome"))
                    End If
                Else
                    oDettaglio.sRagSociale = ""
                End If
                'Comune Sede
                If Not IsDBNull(dvDettaglio.Item(nDett)("comune_nascita_sede")) Then
                    oDettaglio.sComuneSede = CStr(dvDettaglio.Item(nDett)("comune_nascita_sede"))
                Else
                    oDettaglio.sComuneSede = ""
                End If
                'Provincia Sede
                If Not IsDBNull(dvDettaglio.Item(nDett)("pv_nascita_sede")) Then
                    oDettaglio.sPvSede = CStr(dvDettaglio.Item(nDett)("pv_nascita_sede"))
                Else
                    oDettaglio.sPvSede = ""
                End If
            Else
                'Cognome
                If Not IsDBNull(dvDettaglio.Item(nDett)("cognome")) Then
                    oDettaglio.sCognome = CStr(dvDettaglio.Item(nDett)("cognome"))
                Else
                    oDettaglio.sCognome = ""
                End If
                'Nome
                If Not IsDBNull(dvDettaglio.Item(nDett)("nome")) Then
                    oDettaglio.sNome = CStr(dvDettaglio.Item(nDett)("nome"))
                Else
                    oDettaglio.sNome = ""
                End If
                'Denominazione o Ragion Sociale
                oDettaglio.sRagSociale = ""
                'Comune Sede
                oDettaglio.sComuneSede = ""
                'Provincia Sede
                oDettaglio.sPvSede = ""
            End If
            'titolo occupazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_titolo_occupazione")) Then
                oDettaglio.sIdTitoloOccupazione = CStr(dvDettaglio.Item(nDett)("id_titolo_occupazione"))
            Else
                oDettaglio.sIdTitoloOccupazione = ""
            End If
            'occupazione singolo o nucleo familiare
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipo_occupante")) Then
                oDettaglio.sIdNaturaOccupante = CStr(dvDettaglio.Item(nDett)("id_tipo_occupante"))
            Else
                oDettaglio.sIdNaturaOccupante = ""
            End If
            'data inizio occupazione - formato GGMMAAAA
            oDettaglio.sDataInizio = oConst.ReplaceDataForTXT(CStr(dvDettaglio.Item(nDett)("data_inizio")))
            'data fine occupazione - formato GGMMAAAA
            If Not IsDBNull(dvDettaglio.Item(nDett)("data_fine")) Then
                oDettaglio.sDataFine = oConst.ReplaceDataForTXT(CStr(dvDettaglio.Item(nDett)("data_fine")))
            Else
                oDettaglio.sDataFine = ""
            End If
            'destinazione d'uso
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_destinazione_uso")) Then
                oDettaglio.sIdDestinazioneUso = CStr(dvDettaglio.Item(nDett)("id_destinazione_uso"))
            Else
                oDettaglio.sIdDestinazioneUso = ""
            End If
            'comune amministrativo di ubicazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("comune_ubicazione")) Then
                oDettaglio.sComuneAmmUbicazione = CStr(dvDettaglio.Item(nDett)("comune_ubicazione"))
            Else
                oDettaglio.sComuneAmmUbicazione = ""
            End If
            'provincia di ubicazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("pv_ubicazione")) Then
                oDettaglio.sPVUbicazione = CStr(dvDettaglio.Item(nDett)("pv_ubicazione"))
            Else
                oDettaglio.sPVUbicazione = ""
            End If
            'comune catastale di ubicazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("comune_ubicazione_catast")) Then
                oDettaglio.sComuneCatUbicazione = CStr(dvDettaglio.Item(nDett)("comune_ubicazione_catast"))
            Else
                oDettaglio.sComuneCatUbicazione = ""
            End If
            'codice comune catastale
            If Not IsDBNull(dvDettaglio.Item(nDett)("cod_comune_ubicazione_catast")) Then
                oDettaglio.sCodComuneUbicazCatast = CStr(dvDettaglio.Item(nDett)("cod_comune_ubicazione_catast"))
            Else
                oDettaglio.sCodComuneUbicazCatast = ""
            End If
            'tipo unità
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipo_unita")) Then
                oDettaglio.sIdTipoUnita = CStr(dvDettaglio.Item(nDett)("id_tipo_unita"))
            Else
                oDettaglio.sIdTipoUnita = ""
            End If
            'sezione
            If Not IsDBNull(dvDettaglio.Item(nDett)("sezione")) Then
                oDettaglio.sSezione = CStr(dvDettaglio.Item(nDett)("sezione"))
            Else
                oDettaglio.sSezione = ""
            End If
            'foglio
            If Not IsDBNull(dvDettaglio.Item(nDett)("foglio")) Then
                oDettaglio.sFoglio = CStr(dvDettaglio.Item(nDett)("foglio"))
            Else
                oDettaglio.sFoglio = ""
            End If
            'numero
            If Not IsDBNull(dvDettaglio.Item(nDett)("particella")) Then
                oDettaglio.sParticella = CStr(dvDettaglio.Item(nDett)("particella"))
            Else
                oDettaglio.sParticella = ""
            End If
            'estensione particella
            If Not IsDBNull(dvDettaglio.Item(nDett)("estensione_particella")) Then
                oDettaglio.sEstensioneParticella = CStr(dvDettaglio.Item(nDett)("estensione_particella"))
            Else
                oDettaglio.sEstensioneParticella = ""
            End If
            'tipo particella
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipo_particella")) Then
                oDettaglio.sIdTipoParticella = CStr(dvDettaglio.Item(nDett)("id_tipo_particella"))
            Else
                oDettaglio.sIdTipoParticella = ""
            End If
            'subalterno
            If Not IsDBNull(dvDettaglio.Item(nDett)("subalterno")) Then
                oDettaglio.sSubalterno = CStr(dvDettaglio.Item(nDett)("subalterno"))
            Else
                oDettaglio.sSubalterno = ""
            End If
            'Via/Piazza/C.so
            If Not IsDBNull(dvDettaglio.Item(nDett)("indirizzo")) Then
                oDettaglio.sIndirizzo = CStr(dvDettaglio.Item(nDett)("indirizzo"))
            Else
                oDettaglio.sIndirizzo = ""
            End If
            'Civico
            If Not IsDBNull(dvDettaglio.Item(nDett)("civico")) Then
                oDettaglio.sCivico = CStr(dvDettaglio.Item(nDett)("civico"))
            Else
                oDettaglio.sCivico = ""
            End If
            'Interno
            If Not IsDBNull(dvDettaglio.Item(nDett)("interno")) Then
                oDettaglio.sInterno = CStr(dvDettaglio.Item(nDett)("interno"))
            Else
                oDettaglio.sInterno = ""
            End If
            'Scala
            If Not IsDBNull(dvDettaglio.Item(nDett)("scala")) Then
                oDettaglio.sScala = CStr(dvDettaglio.Item(nDett)("scala"))
            Else
                oDettaglio.sScala = ""
            End If
            'codice assenza dati castatali
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_assenza_dati_catastali")) Then
                oDettaglio.sIdAssenzaDatiCatastali = CStr(dvDettaglio.Item(nDett)("id_assenza_dati_catastali"))
            Else
                oDettaglio.sIdAssenzaDatiCatastali = ""
            End If
            'filler
            oDettaglio.sFiller = ""
            '***************************************************
            'Scrivo la riga nel file
            '***************************************************
            sDati = TARSU_FormattaRcDettaglio(oDettaglio)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
                nRcFile += 1
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcDettaglio::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcDettaglio::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function TARSU_FormattaRcTesta(ByVal MyRcTesta As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            Log.Debug("ServiceOPENae::TARSU_FormattaRcTesta::inizio procedura")
            'tipo record = Vale sempre 0
            sDatiTxt = oConst.TIPORCTESTA
            'identificativo fornitura = "SMRIF"
            sDatiTxt += oConst.TARSU_IDFORNITURA
            'Codice numerico fornitura
            sDatiTxt += oConst.TARSU_CODNUMFORNITURA
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcTesta.sCFPIVA.PadLeft(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sCognome).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sNome).PadRight(25, " ")
            'sesso
            sDatiTxt += MyRcTesta.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcTesta.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sComuneNascita).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcTesta.sPvNascita.PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sRagSociale).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sComuneSede).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcTesta.sPvSede.PadRight(2, " ")
            'anno di riferimento
            sDatiTxt += MyRcTesta.sAnno.PadRight(4, " ")
            'filler
            sDatiTxt += MyRcTesta.sFiller.PadRight(135, " ")
            'carattere di controllo
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            Log.Debug("ServiceOPENae::TARSU_FormattaRcTesta::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcTesta::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcTesta::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function TARSU_FormattaRcDettaglio(ByVal MyRcDettaglio As StructureDettaglio) As String
        Dim sDatiTxt As String = ""
        Try
            Log.Debug("ServiceOPENae::TARSU_FormattaRcDettaglio::inizio procedura")

            'tipo record = Vale sempre 1
            sDatiTxt = oConst.TIPORCDETTAGLIO
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcDettaglio.sCFPIVA.PadLeft(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sCognome).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sNome).PadRight(25, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sRagSociale).PadRight(50, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sComuneSede).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcDettaglio.sPvSede.PadRight(2, " ")
            'titolo occupazione
            sDatiTxt += MyRcDettaglio.sIdTitoloOccupazione
            'occupazione singolo o nucleo familiare
            sDatiTxt += MyRcDettaglio.sIdNaturaOccupante
            'data inizio occupazione - formato GGMMAAAA
            sDatiTxt += MyRcDettaglio.sDataInizio.PadRight(8, " ")
            'data fine occupazione - formato GGMMAAAA
            sDatiTxt += MyRcDettaglio.sDataFine.PadRight(8, " ")
            'destinazione d'uso
            sDatiTxt += MyRcDettaglio.sIdDestinazioneUso
            'comune amministrativo di ubicazione dell'immobile
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sComuneAmmUbicazione).PadRight(20, " ")
            'provincia di ubicazione dell'immobile
            sDatiTxt += MyRcDettaglio.sPVUbicazione.PadRight(2, " ")
            'comune catastale di ubicazione dell'immobile
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sComuneCatUbicazione).PadRight(20, " ")
            'codice comune catastale
            sDatiTxt += MyRcDettaglio.sCodComuneUbicazCatast.PadRight(5, " ")
            'tipo unità
            sDatiTxt += MyRcDettaglio.sIdTipoUnita
            'sezione
            sDatiTxt += MyRcDettaglio.sSezione.PadRight(3, " ")
            'foglio
            sDatiTxt += MyRcDettaglio.sFoglio.PadRight(5, " ")
            'numero
            sDatiTxt += MyRcDettaglio.sParticella.PadRight(5, " ")
            'estensione particella
            sDatiTxt += MyRcDettaglio.sEstensioneParticella.PadRight(4, " ")
            'tipo particella
            sDatiTxt += MyRcDettaglio.sIdTipoParticella
            'subalterno
            sDatiTxt += MyRcDettaglio.sSubalterno.PadRight(4, " ")
            'Via/Piazza/C.so
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sIndirizzo).PadRight(30, " ")
            'Civico
            sDatiTxt += MyRcDettaglio.sCivico.PadRight(6, " ")
            'Interno
            sDatiTxt += MyRcDettaglio.sInterno.PadRight(2, " ")
            'Scala
            sDatiTxt += MyRcDettaglio.sScala.PadRight(1, " ")
            'Codice assenza dati catastali
            sDatiTxt += MyRcDettaglio.sIdAssenzaDatiCatastali
            'Filler
            sDatiTxt += MyRcDettaglio.sFiller.PadRight(78, " ")
            'Carattere di controllo="A"
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            Log.Debug("ServiceOPENae::TARSU_FormattaRcDettaglio::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcDettaglio::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcDettaglio::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function TARSU_FormattaRcCoda(ByVal MyRcCoda As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            Log.Debug("ServiceOPENae::TARSU_FormattaRcCoda::inizio procedura")

            'tipo record = Vale sempre 9
            sDatiTxt = oConst.TIPORCCODA
            'identificativo fornitura = "SMRIF"
            sDatiTxt += oConst.TARSU_IDFORNITURA
            'Codice numerico fornitura
            sDatiTxt += oConst.TARSU_CODNUMFORNITURA
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcCoda.sCFPIVA.PadLeft(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sCognome).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sNome).PadRight(25, " ")
            'sesso
            sDatiTxt += MyRcCoda.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcCoda.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sComuneNascita).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcCoda.sPvNascita.PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sRagSociale).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sComuneSede).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcCoda.sPvSede.PadRight(2, " ")
            'anno di riferimento
            sDatiTxt += MyRcCoda.sAnno.PadRight(4, " ")
            'filler
            sDatiTxt += MyRcCoda.sFiller.PadRight(135, " ")
            'carattere di controllo
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            Log.Debug("ServiceOPENae::TARSU_FormattaRcCoda::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcCoda::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcCoda::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function
#End Region

#Region "H2O"
    Private Function H2O_EstraiTracciato(ByVal dvMyDati As DataView, ByVal sPathCartellaTracciatiCompleto As String, ByRef nRcFile As Integer) As Boolean
        Dim x As Integer

        Try
            '***************************************************
            'scrivo il record di testa
            '***************************************************
            If H2O_WriteRcTesta(dvMyDati, sPathCartellaTracciatiCompleto, nRcFile) = False Then
                Return False
            End If
            For x = 0 To dvMyDati.Count
                '***************************************************
                'record di dettaglio
                '***************************************************
                If H2O_WriteRcDettaglio(dvMyDati, x, sPathCartellaTracciatiCompleto, nRcFile) = False Then
                    Return False
                End If
            Next
            '***************************************************
            'scrivo il record di coda
            '***************************************************
            If H2O_WriteRcTesta(Nothing, sPathCartellaTracciatiCompleto, -1) = False Then
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::H2O_EstraiTracciato::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::H2O_EstraiTracciato::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function H2O_WriteRcTesta(ByVal dvTesta As DataView, ByVal sPathCartellaTracciatiCompleto As String, ByRef nRcFile As Integer) As Boolean
        Dim oTestaCoda As New StructureTestaCoda

        Try
            'Codice Fiscale - Partita IVA
            oTestaCoda.sCFPIVA = CStr(dvTesta.Item(0)("cod_fiscale_ente"))
            If IsNumeric(dvTesta.Item(0)("cod_fiscale_ente")) Then
                '***persona giuridica***
                'Cognome
                oTestaCoda.sCognome = ""
                'Nome
                oTestaCoda.sNome = ""
                'sesso
                oTestaCoda.sSesso = ""
                'data nascita
                oTestaCoda.sDataNascita = ""
                'comune nascita
                oTestaCoda.sComuneNascita = ""
                'provincia nascita
                oTestaCoda.sPvNascita = ""
                'Denominazione o Ragion Sociale
                If Not IsDBNull(dvTesta.Item(0)("cognome_ente")) Then
                    oTestaCoda.sRagSociale = CStr(dvTesta.Item(0)("cognome_ente"))
                    If Not IsDBNull(dvTesta.Item(0)("nome_ente")) Then
                        oTestaCoda.sRagSociale += " " & CStr(dvTesta.Item(0)("nome_ente"))
                    End If
                Else
                    oTestaCoda.sRagSociale = ""
                End If
                'Comune Sede
                If Not IsDBNull(dvTesta.Item(0)("comune_nascita_sede_ente")) Then
                    oTestaCoda.sComuneSede = CStr(dvTesta.Item(0)("comune_nascita_sede_ente"))
                Else
                    oTestaCoda.sComuneSede = ""
                End If
                'Provincia Sede
                If Not IsDBNull(dvTesta.Item(0)("pv_nascita_sede_ente")) Then
                    oTestaCoda.sPvSede = CStr(dvTesta.Item(0)("pv_nascita_sede_ente"))
                Else
                    oTestaCoda.sPvSede = ""
                End If
            Else
                '***persona fisica***
                'Cognome
                If Not IsDBNull(dvTesta.Item(0)("cognome_ente")) Then
                    oTestaCoda.sCognome = CStr(dvTesta.Item(0)("cognome_ente"))
                Else
                    oTestaCoda.sCognome = ""
                End If
                'Nome
                If Not IsDBNull(dvTesta.Item(0)("nome_ente")) Then
                    oTestaCoda.sNome = CStr(dvTesta.Item(0)("nome_ente"))
                Else
                    oTestaCoda.sNome = ""
                End If
                'sesso
                If Not IsDBNull(dvTesta.Item(0)("sesso_ente")) Then
                    oTestaCoda.sSesso = CStr(dvTesta.Item(0)("sesso_ente"))
                Else
                    oTestaCoda.sSesso = ""
                End If
                'data nascita
                If Not IsDBNull(dvTesta.Item(0)("data_nascita_ente")) Then
                    oTestaCoda.sDataNascita = oConst.ReplaceDataForTXT(CStr(dvTesta.Item(0)("data_nascita_ente")))
                Else
                    oTestaCoda.sDataNascita = ""
                End If
                'comune nascita
                If Not IsDBNull(dvTesta.Item(0)("comune_nascita_sede_ente")) Then
                    oTestaCoda.sComuneNascita = CStr(dvTesta.Item(0)("comune_nascita_sede_ente"))
                Else
                    oTestaCoda.sComuneNascita = ""
                End If
                'provincia nascita
                If Not IsDBNull(dvTesta.Item(0)("pv_nascita_sede_ente")) Then
                    oTestaCoda.sPvNascita = CStr(dvTesta.Item(0)("pv_nascita_sede_ente"))
                Else
                    oTestaCoda.sPvNascita = ""
                End If
                'Denominazione o Ragion Sociale
                oTestaCoda.sRagSociale = ""
                'Comune Sede
                oTestaCoda.sComuneSede = ""
                'Provincia Sede
                oTestaCoda.sPvSede = ""
            End If
            'anno
            oTestaCoda.sAnno = CStr(dvTesta.Item(0)("anno"))
            'numero progressivo della fornitura nel formato AAAANNN
            oTestaCoda.sprogfornitura = CStr(dvTesta.Item(0)("anno")) & CStr(dvTesta.Item(0)("id_flusso")).PadLeft(3, "0")
            'data di invio nel formato GGMMAAAA
            oTestaCoda.sdatainvio = oConst.ReplaceDataForTXT(oConst.ReplaceDataForDB(Now.ToString))
            'filler
            oTestaCoda.sFiller = ""
            sDati = H2O_FormattaRcTesta(oTestaCoda)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
                nRcFile += 1
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcTesta::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcTesta::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function H2O_WriteRcCoda(ByVal oCoda As StructureTestaCoda, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Try
            sDati = H2O_FormattaRcCoda(oCoda)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcCoda::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcCoda::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function H2O_WriteRcDettaglio(ByVal dvDettaglio As DataView, ByVal nDett As Integer, ByVal sPathCartellaTracciatiCompleto As String, ByRef nRcFile As Integer) As Boolean
        Dim oDettaglio As New StructureDettaglio

        Try
            'Codice Fiscale - Partita IVA
            oDettaglio.sCFPIVA = CStr(dvDettaglio.Item(nDett)("cod_fiscale"))
            If IsNumeric(dvDettaglio.Item(nDett)("cod_fiscale")) Then
                'Cognome
                oDettaglio.sCognome = ""
                'Nome
                oDettaglio.sNome = ""
                'sesso
                oDettaglio.ssesso = ""
                'data di nascita
                oDettaglio.sdatanascita = ""
                'comune di nascita
                oDettaglio.scomunenascita = ""
                'provincia di nascita
                oDettaglio.spvnascita = ""
                'Denominazione o Ragion Sociale
                If Not IsDBNull(dvDettaglio.Item(nDett)("cognome")) Then
                    oDettaglio.sRagSociale = CStr(dvDettaglio.Item(nDett)("cognome"))
                    If Not IsDBNull(dvDettaglio.Item(nDett)("nome")) Then
                        oDettaglio.sRagSociale += " " & CStr(dvDettaglio.Item(nDett)("nome"))
                    End If
                Else
                    oDettaglio.sRagSociale = ""
                End If
                'Comune Sede
                If Not IsDBNull(dvDettaglio.Item(nDett)("comune_nascita_sede")) Then
                    oDettaglio.sComuneSede = CStr(dvDettaglio.Item(nDett)("comune_nascita_sede"))
                Else
                    oDettaglio.sComuneSede = ""
                End If
                'Provincia Sede
                If Not IsDBNull(dvDettaglio.Item(nDett)("pv_nascita_sede")) Then
                    oDettaglio.sPvSede = CStr(dvDettaglio.Item(nDett)("pv_nascita_sede"))
                Else
                    oDettaglio.sPvSede = ""
                End If
            Else
                'Cognome
                If Not IsDBNull(dvDettaglio.Item(nDett)("cognome")) Then
                    oDettaglio.sCognome = CStr(dvDettaglio.Item(nDett)("cognome"))
                Else
                    oDettaglio.sCognome = ""
                End If
                'Nome
                If Not IsDBNull(dvDettaglio.Item(nDett)("nome")) Then
                    oDettaglio.sNome = CStr(dvDettaglio.Item(nDett)("nome"))
                Else
                    oDettaglio.sNome = ""
                End If
                'sesso
                If Not IsDBNull(dvDettaglio.Item(0)("sesso")) Then
                    oDettaglio.sSesso = CStr(dvDettaglio.Item(0)("sesso"))
                Else
                    oDettaglio.sSesso = ""
                End If
                'data nascita
                If Not IsDBNull(dvDettaglio.Item(0)("data_nascita")) Then
                    oDettaglio.sDataNascita = oConst.ReplaceDataForTXT(CStr(dvDettaglio.Item(0)("data_nascita")))
                Else
                    oDettaglio.sDataNascita = ""
                End If
                'comune nascita
                If Not IsDBNull(dvDettaglio.Item(0)("comune_nascita_sede")) Then
                    oDettaglio.sComuneNascita = CStr(dvDettaglio.Item(0)("comune_nascita_sede"))
                Else
                    oDettaglio.sComuneNascita = ""
                End If
                'provincia nascita
                If Not IsDBNull(dvDettaglio.Item(0)("pv_nascita_sede")) Then
                    oDettaglio.sPvNascita = CStr(dvDettaglio.Item(0)("pv_nascita_sede"))
                Else
                    oDettaglio.sPvNascita = ""
                End If
                'Denominazione o Ragion Sociale
                oDettaglio.sRagSociale = ""
                'Comune Sede
                oDettaglio.sComuneSede = ""
                'Provincia Sede
                oDettaglio.sPvSede = ""
            End If
            'comune del domicilio fiscale
            If Not IsDBNull(dvDettaglio.Item(0)("comune_dom_fiscale")) Then
                oDettaglio.sComunedomfiscale = CStr(dvDettaglio.Item(0)("comune_dom_fiscale"))
            Else
                oDettaglio.sComunedomfiscale = ""
            End If
            'provincia del domicilio fiscale
            If Not IsDBNull(dvDettaglio.Item(0)("pv_dom_fiscale")) Then
                oDettaglio.sPvdomfiscale = CStr(dvDettaglio.Item(0)("pv_dom_fiscale"))
            Else
                oDettaglio.sPvdomfiscale = ""
            End If
            'titolo occupazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_titolo_occupazione")) Then
                oDettaglio.sIdTitoloOccupazione = CStr(dvDettaglio.Item(nDett)("id_titolo_occupazione"))
            Else
                oDettaglio.sIdTitoloOccupazione = ""
            End If
            'estremi del contratto
            If Not IsDBNull(dvDettaglio.Item(nDett)("estremi_contratto")) Then
                oDettaglio.sestremicontratto = CStr(dvDettaglio.Item(nDett)("estremi_contratto"))
            Else
                oDettaglio.sestremicontratto = ""
            End If
            'data inizio occupazione - formato GGMMAAAA
            oDettaglio.sDataInizio = oConst.ReplaceDataForTXT(CStr(dvDettaglio.Item(nDett)("data_inizio")))
            'tipologia utenza
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipologia_utenza")) Then
                oDettaglio.sidtipologiautenza = CStr(dvDettaglio.Item(nDett)("id_tipologia_utenza"))
            Else
                oDettaglio.sidtipologiautenza = ""
            End If
            'comune amministrativo di ubicazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("comune_ubicazione")) Then
                oDettaglio.sComuneAmmUbicazione = CStr(dvDettaglio.Item(nDett)("comune_ubicazione"))
            Else
                oDettaglio.sComuneAmmUbicazione = ""
            End If
            'provincia di ubicazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("pv_ubicazione")) Then
                oDettaglio.sPVUbicazione = CStr(dvDettaglio.Item(nDett)("pv_ubicazione"))
            Else
                oDettaglio.sPVUbicazione = ""
            End If
            'comune catastale di ubicazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("comune_ubicazione_catast")) Then
                oDettaglio.sComuneCatUbicazione = CStr(dvDettaglio.Item(nDett)("comune_ubicazione_catast"))
            Else
                oDettaglio.sComuneCatUbicazione = ""
            End If
            'codice comune catastale
            If Not IsDBNull(dvDettaglio.Item(nDett)("cod_comune_ubicazione_catast")) Then
                oDettaglio.sCodComuneUbicazCatast = CStr(dvDettaglio.Item(nDett)("cod_comune_ubicazione_catast"))
            Else
                oDettaglio.sCodComuneUbicazCatast = ""
            End If
            'Via/Piazza/C.so
            If Not IsDBNull(dvDettaglio.Item(nDett)("indirizzo")) Then
                oDettaglio.sIndirizzo = CStr(dvDettaglio.Item(nDett)("indirizzo"))
            Else
                oDettaglio.sIndirizzo = ""
            End If
            'tipo unità
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipo_unita")) Then
                oDettaglio.sIdTipoUnita = CStr(dvDettaglio.Item(nDett)("id_tipo_unita"))
            Else
                oDettaglio.sIdTipoUnita = ""
            End If
            'sezione
            If Not IsDBNull(dvDettaglio.Item(nDett)("sezione")) Then
                oDettaglio.sSezione = CStr(dvDettaglio.Item(nDett)("sezione"))
            Else
                oDettaglio.sSezione = ""
            End If
            'foglio
            If Not IsDBNull(dvDettaglio.Item(nDett)("foglio")) Then
                oDettaglio.sFoglio = CStr(dvDettaglio.Item(nDett)("foglio"))
            Else
                oDettaglio.sFoglio = ""
            End If
            'numero
            If Not IsDBNull(dvDettaglio.Item(nDett)("particella")) Then
                oDettaglio.sParticella = CStr(dvDettaglio.Item(nDett)("particella"))
            Else
                oDettaglio.sParticella = ""
            End If
            'estensione particella
            If Not IsDBNull(dvDettaglio.Item(nDett)("estensione_particella")) Then
                oDettaglio.sEstensioneParticella = CStr(dvDettaglio.Item(nDett)("estensione_particella"))
            Else
                oDettaglio.sEstensioneParticella = ""
            End If
            'tipo particella
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipo_particella")) Then
                oDettaglio.sIdTipoParticella = CStr(dvDettaglio.Item(nDett)("id_tipo_particella"))
            Else
                oDettaglio.sIdTipoParticella = ""
            End If
            'subalterno
            If Not IsDBNull(dvDettaglio.Item(nDett)("subalterno")) Then
                oDettaglio.sSubalterno = CStr(dvDettaglio.Item(nDett)("subalterno"))
            Else
                oDettaglio.sSubalterno = ""
            End If
            'codice assenza dati castatali
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_assenza_dati_catastali")) Then
                oDettaglio.sIdAssenzaDatiCatastali = CStr(dvDettaglio.Item(nDett)("id_assenza_dati_catastali"))
            Else
                oDettaglio.sIdAssenzaDatiCatastali = ""
            End If
            'numero mesi di fatturazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("mesi_fatturazione")) Then
                oDettaglio.smesifatturazione = CStr(dvDettaglio.Item(nDett)("mesi_fatturazione"))
            Else
                oDettaglio.smesifatturazione = ""
            End If
            'segno importo
            If Not IsDBNull(dvDettaglio.Item(nDett)("segno_spesa")) Then
                oDettaglio.sSegnoSpesa = CStr(dvDettaglio.Item(nDett)("segno_spesa"))
            Else
                oDettaglio.sSegnoSpesa = ""
            End If
            'spesa consumo
            If Not IsDBNull(dvDettaglio.Item(nDett)("spesa_consumo")) Then
                oDettaglio.sspesaconsumo = CStr(dvDettaglio.Item(nDett)("spesa_consumo"))
            Else
                oDettaglio.sspesaconsumo = ""
            End If
            'filler
            oDettaglio.sFiller = ""
            '***************************************************
            'Scrivo la riga nel file
            '***************************************************
            sDati = H2O_FormattaRcDettaglio(oDettaglio)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
                nRcFile += 1
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcDettaglio::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcDettaglio::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function H2O_FormattaRcTesta(ByVal MyRcTesta As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            Log.Debug("ServiceOPENae::H2O_FormattaRcTesta::inizio procedura")
            'tipo record = Vale sempre 0
            sDatiTxt = oConst.TIPORCTESTA
            'identificativo fornitura = "NWIDR"
            sDatiTxt += oConst.H2O_IDFORNITURA
            'Codice numerico fornitura=24
            sDatiTxt += oConst.H2O_CODNUMFORNITURA
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcTesta.sCFPIVA.PadLeft(16, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sRagSociale).PadRight(50, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sComuneSede).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcTesta.sPvSede.PadRight(2, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sCognome).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sNome).PadRight(25, " ")
            'sesso
            sDatiTxt += MyRcTesta.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcTesta.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerXML(MyRcTesta.sComuneNascita).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcTesta.sPvNascita.PadRight(2, " ")
            'anno di riferimento
            sDatiTxt += MyRcTesta.sAnno.PadRight(4, " ")
            'progressivo invio
            sDatiTxt += MyRcTesta.sprogfornitura
            'data di invio
            sDatiTxt += MyRcTesta.sDataInvio.PadRight(8, " ")
            'filler
            sDatiTxt += MyRcTesta.sFiller.PadRight(130, " ")
            'carattere di controllo
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            Log.Debug("ServiceOPENae::H2O_FormattaRcTesta::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcTesta::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcTesta::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function H2O_FormattaRcDettaglio(ByVal MyRcDettaglio As StructureDettaglio) As String
        Dim sDatiTxt As String = ""
        Try
            Log.Debug("ServiceOPENae::H2O_FormattaRcDettaglio::inizio procedura")

            'tipo record = Vale sempre 1
            sDatiTxt = oConst.TIPORCDETTAGLIO
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcDettaglio.sCFPIVA.PadLeft(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sCognome).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sNome).PadRight(25, " ")
            'Sesso
            sDatiTxt += MyRcDettaglio.ssesso.PadRight(1, " ")
            'Data di nascita
            sDatiTxt += MyRcDettaglio.sdatanascita.PadRight(8, " ")
            'comune di nascita
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.scomunenascita).PadRight(20, " ")
            'Provincia di nascita
            sDatiTxt += MyRcDettaglio.spvnascita.PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sRagSociale).PadRight(50, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sComuneSede).PadRight(20, " ")
            'Provincia Sede
            sDatiTxt += MyRcDettaglio.sPvSede.PadRight(2, " ")
            'titolo occupazione
            sDatiTxt += MyRcDettaglio.sIdTitoloOccupazione
            'estremi contratto
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sestremicontratto).padright(30, " ")
            'data prima attivazioen - formato GGMMAAAA
            sDatiTxt += MyRcDettaglio.sDataInizio.PadRight(8, " ")
            'tipologia utenza
            sDatiTxt += MyRcDettaglio.sIdTipologiaUtenza
            'comune amministrativo di ubicazione dell'immobile
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sComuneAmmUbicazione).PadRight(20, " ")
            'provincia di ubicazione dell'immobile
            sDatiTxt += MyRcDettaglio.sPVUbicazione.PadRight(2, " ")
            'comune catastale di ubicazione dell'immobile
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sComuneCatUbicazione).PadRight(20, " ")
            'codice comune catastale
            sDatiTxt += MyRcDettaglio.sCodComuneUbicazCatast.PadRight(5, " ")
            'Via/Piazza/C.so
            sDatiTxt += oConst.FormattaPerXML(MyRcDettaglio.sIndirizzo).PadRight(35, " ")
            'tipo unità
            sDatiTxt += MyRcDettaglio.sIdTipoUnita
            'sezione
            sDatiTxt += MyRcDettaglio.sSezione.PadRight(3, " ")
            'foglio
            sDatiTxt += MyRcDettaglio.sFoglio.PadRight(5, " ")
            'numero
            sDatiTxt += MyRcDettaglio.sParticella.PadRight(5, " ")
            'estensione particella
            sDatiTxt += MyRcDettaglio.sEstensioneParticella.PadRight(4, " ")
            'tipo particella
            sDatiTxt += MyRcDettaglio.sIdTipoParticella
            'subalterno
            sDatiTxt += MyRcDettaglio.sSubalterno.PadRight(4, " ")
            'Codice assenza dati catastali
            sDatiTxt += MyRcDettaglio.sIdAssenzaDatiCatastali
            'numero di mesi di fatturazione
            sDatiTxt += MyRcDettaglio.smesifatturazione.PadRight(2, " ")
            'segno spesa
            sDatiTxt += MyRcDettaglio.ssegnospesa
            'spesa consumo
            sDatiTxt += MyRcDettaglio.sspesaconsumo.Padleft(9, "0")
            'Filler
            sDatiTxt += MyRcDettaglio.sFiller.PadRight(16, " ")
            'Carattere di controllo="A"
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            Log.Debug("ServiceOPENae::H2O_FormattaRcDettaglio::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcDettaglio::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcDettaglio::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function H2O_FormattaRcCoda(ByVal MyRcCoda As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            Log.Debug("ServiceOPENae::H2O_FormattaRcCoda::inizio procedura")

            'tipo record = Vale sempre 9
            sDatiTxt = oConst.TIPORCCODA
            'identificativo fornitura = "NWIDR"
            sDatiTxt += oConst.H2O_IDFORNITURA
            'Codice numerico fornitura
            sDatiTxt += oConst.H2O_CODNUMFORNITURA
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcCoda.sCFPIVA.PadLeft(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sCognome).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sNome).PadRight(25, " ")
            'sesso
            sDatiTxt += MyRcCoda.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcCoda.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sComuneNascita).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcCoda.sPvNascita.PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sRagSociale).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerXML(MyRcCoda.sComuneSede).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcCoda.sPvSede.PadRight(2, " ")
            'anno di riferimento
            sDatiTxt += MyRcCoda.sAnno.PadRight(4, " ")
            'filler
            sDatiTxt += MyRcCoda.sFiller.PadRight(135, " ")
            'Carattere di controllo="A"
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            Log.Debug("ServiceOPENae::H2O_FormattaRcCoda::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcCoda::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcCoda::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function
#End Region
End Class
