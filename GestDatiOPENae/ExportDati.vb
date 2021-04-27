Imports log4net
Imports Utility
Imports System.Configuration

Public Class ExportDati
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ExportDati))
    Private oConst As New General
    Private sMsgErr, sDati As String
    Private AppReader As New System.Configuration.AppSettingsReader
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
        Dim sTipoContratto As String
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
        Dim sConsumo As String
        Dim sImportoFatturato As String
        Dim sFiller As String
    End Structure

    Public Sub New()
        log.Debug("Istanziata la classe ExportDati")
    End Sub

    Public Function EstraiTracciato(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByRef sNomeFileTracciati As String) As String
        Dim cnn As New DBConnection
        Dim oDBManager As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim oQueryManager As New ClsInterDB
        Dim dvMyDati As DataView
        Dim sPathCartellaTracciatiCompleto As String
        Dim nRcFile As Integer = 0

        Try
            log.Debug("ServiceOPENae::EstraiTracciato::inizio procedura")
            'il nome del tracciato è formattato nel modo seguente:
            '-EXPORT_AGENZIA_ENTRATE
            '-Tributo
            '-Codice ISTAT
            '-Anno
            '-AAAAMMGGHHMMSS
            sNomeFileTracciati = "EXPORT_AGENZIA_ENTRATE_" + sTributo + "_" + sCodiceISTAT + "_" + sAnnoRif + "_" + DateTime.Now.ToString("yyyyMMdd_hhmmss")
            sPathCartellaTracciatiCompleto = ConfigurationSettings.AppSettings("PathTracciati") + ConfigurationSettings.AppSettings("NomeCartellaTracciati") + sNomeFileTracciati + ".txt"

            'oDBManager = cnn.DBConnection()

            dvMyDati = oQueryManager.GetDisposizione(oDBManager, sCodiceISTAT, sTributo, sAnnoRif)
            If dvMyDati.Count > 0 Then
                Select Case sTributo
                    Case oConst.TRIBUTO_TARSU, oConst.TRIBUTO_TIA
                        If TARSU_EstraiTracciato(dvMyDati, sPathCartellaTracciatiCompleto, nRcFile) = False Then
                            Return ""
                        End If
                    Case oConst.TRIBUTO_H2O
                        If H2O_EstraiTracciato(dvMyDati, sPathCartellaTracciatiCompleto, nRcFile) = False Then
                            Return ""
                        End If
                    Case Else
                        Return ""
                End Select
            End If
            dvMyDati.Dispose()
            '*********************************************************************
            'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
            '*********************************************************************
            If oQueryManager.SetFlussoEstratto(oDBManager, sCodiceISTAT, sTributo, sAnnoRif, sNomeFileTracciati, nRcFile) = -1 Then
                Return ""
            End If
            sPathCartellaTracciatiCompleto = sPathCartellaTracciatiCompleto.Replace(AppReader.GetValue("PathTracciati", GetType(String)), AppReader.GetValue("PathTracciatiForDownload", GetType(String)))
            log.Debug("ServiceOPENae::EstraiTracciato::file da scaricare::" & sPathCartellaTracciatiCompleto)
            log.Debug("ServiceOPENae::EstraiTracciato::fine procedura")
            Return sPathCartellaTracciatiCompleto
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::" & Err.Message)
            Return ""
        End Try
    End Function

    Public Function EstraiTracciato(ByVal sCodiceISTAT As String, ByVal sCodBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByRef sNomeFileTracciati As String) As String
        Dim cnn As New DBConnection
        Dim oDBManager As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim oQueryManager As New ClsInterDB
        Dim dvMyDati As DataView
        Dim sPathCartellaTracciatiCompleto As String
        Dim nMyIdFlusso As Integer = -1

        Try
            log.Debug("ServiceOPENae::EstraiTracciato::inizio procedura")
            'il nome del tracciato è formattato nel modo seguente:
            '-EXPORT_AGENZIA_ENTRATE
            '-Tributo
            '-Codice ISTAT
            '-Anno
            '-AAAAMMGGHHMMSS
            sNomeFileTracciati = "EXPORT_MEF_" + sTributo + "_" + sCodiceISTAT + "_" + sAnnoRif + "_" + DateTime.Now.ToString("yyyyMMdd_hhmmss")
            sPathCartellaTracciatiCompleto = ConfigurationSettings.AppSettings("PathTracciati") + ConfigurationSettings.AppSettings("NomeCartellaTracciati") + sNomeFileTracciati + ".txt"

            'oDBManager = cnn.DBConnection()

            dvMyDati = oQueryManager.GetDisposizioneICI(oDBManager, sCodiceISTAT, sAnnoRif)
            If Not IsNothing(dvMyDati) Then
                If dvMyDati.Count > 0 Then
                    Select Case sTributo
                        Case oConst.TRIBUTO_ICI
                            If ICI_EstraiTracciato(oDBManager, dvMyDati, sCodiceISTAT, sCodBelfiore, sDescrEnte, sCAPEnte, sTributo, sAnnoRif, sDataScadenza, nProgInvio, sPathCartellaTracciatiCompleto) = False Then
                                Return ""
                            End If
                        Case Else
                            Return ""
                    End Select
                End If
                dvMyDati.Dispose()
                '*********************************************************************
                'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
                '*********************************************************************
                nMyIdFlusso = oQueryManager.SetIdFlussoEstratto(oDBManager, sCodiceISTAT, sAnnoRif, sTributo)
                If nMyIdFlusso = -1 Then
                    Return ""
                End If
                If oQueryManager.SetFlussoEstratto(oDBManager, sCodiceISTAT, sTributo, sAnnoRif, sNomeFileTracciati, 1) = -1 Then
                    Return ""
                End If
            End If
            sPathCartellaTracciatiCompleto = sPathCartellaTracciatiCompleto.Replace(AppReader.GetValue("PathTracciati", GetType(String)), AppReader.GetValue("PathTracciatiForDownload", GetType(String)))
            log.Debug("ServiceOPENae::EstraiTracciato::file da scaricare::" & sPathCartellaTracciatiCompleto)
            log.Debug("ServiceOPENae::EstraiTracciato::fine procedura")
            Return sPathCartellaTracciatiCompleto
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::" & Err.Message)
            Return ""
        End Try
    End Function

    Private Function WriteFile(ByVal sFile As String, ByVal DatiToPrint As String, ByVal ErrWriteFile As String) As Integer
        Dim MyFileToWrite As IO.StreamWriter = IO.File.AppendText(sFile)
        Dim sDatiFile As String = ""

        Try
            sDatiFile = DatiToPrint

            MyFileToWrite.WriteLine(sDatiFile.ToUpper)
            MyFileToWrite.Flush()

            Return 1
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::WriteFile::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::WriteFile::" & Err.Message)
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
            For x = 0 To dvMyDati.Count - 1
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
            If TARSU_WriteRcTesta(dvMyDati, sPathCartellaTracciatiCompleto, -1) = False Then
                Return False
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_EstraiTracciato::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_EstraiTracciato::" & Err.Message)
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
            If nRcFile = -1 Then
                sDati = TARSU_FormattaRcCoda(oTestaCoda)
            Else
                sDati = TARSU_FormattaRcTesta(oTestaCoda)
            End If
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
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcTesta::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcTesta::" & Err.Message)
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
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcCoda::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcCoda::" & Err.Message)
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
                If Not IsDBNull(dvDettaglio.Item(nDett)("comune_sede")) Then
                    oDettaglio.sComuneSede = CStr(dvDettaglio.Item(nDett)("comune_sede"))
                Else
                    oDettaglio.sComuneSede = ""
                End If
                'Provincia Sede
                If Not IsDBNull(dvDettaglio.Item(nDett)("pv_sede")) Then
                    oDettaglio.sPvSede = CStr(dvDettaglio.Item(nDett)("pv_sede"))
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
                If CStr(dvDettaglio.Item(nDett)("estensione_particella")) <> "0" Then
                    oDettaglio.sEstensioneParticella = CStr(dvDettaglio.Item(nDett)("estensione_particella"))
                Else
                    oDettaglio.sEstensioneParticella = ""
                End If
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
            'controllo la presenza di tutti i dati castali per valorizzare correttamente il flag assenza
            If oDettaglio.sFoglio = "" Or oDettaglio.sParticella = "" Then
                oDettaglio.sIdTipoUnita = "" : oDettaglio.sFoglio = "" : oDettaglio.sParticella = "" : oDettaglio.sSubalterno = ""
                If oDettaglio.sIdAssenzaDatiCatastali = "" Then
                    oDettaglio.sIdAssenzaDatiCatastali = "3"
                End If
            Else
                If oDettaglio.sIdAssenzaDatiCatastali = "3" Then
                    oDettaglio.sIdAssenzaDatiCatastali = ""
                End If
            End If
            '***************************************************
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
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcDettaglio::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::TARSU_WriteRcDettaglio::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function TARSU_FormattaRcTesta(ByVal MyRcTesta As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::TARSU_FormattaRcTesta::inizio procedura")
            'tipo record = Vale sempre 0
            sDatiTxt = oConst.TIPORCTESTA
            'identificativo fornitura = "SMRIF"
            sDatiTxt += oConst.TARSU_IDFORNITURA
            'Codice numerico fornitura
            sDatiTxt += oConst.TARSU_CODNUMFORNITURA
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcTesta.sCFPIVA.PadRight(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sCognome, 26).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sNome, 25).PadRight(25, " ")
            'sesso
            sDatiTxt += MyRcTesta.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcTesta.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sComuneNascita, 40).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcTesta.sPvNascita.PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sRagSociale, 60).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sComuneSede, 40).PadRight(40, " ")
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

            log.Debug("ServiceOPENae::TARSU_FormattaRcTesta::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcTesta::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcTesta::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function TARSU_FormattaRcDettaglio(ByVal MyRcDettaglio As StructureDettaglio) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::TARSU_FormattaRcDettaglio::inizio procedura")

            'tipo record = Vale sempre 1
            sDatiTxt = oConst.TIPORCDETTAGLIO
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcDettaglio.sCFPIVA.PadRight(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sCognome, 26).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sNome, 25).PadRight(25, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sRagSociale, 50).PadRight(50, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneSede, 40).PadRight(40, " ")
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
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneAmmUbicazione, 20).PadRight(20, " ")
            'provincia di ubicazione dell'immobile
            sDatiTxt += MyRcDettaglio.sPVUbicazione.PadRight(2, " ")
            'comune catastale di ubicazione dell'immobile
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneCatUbicazione, 20).PadRight(20, " ")
            'codice comune catastale
            sDatiTxt += MyRcDettaglio.sCodComuneUbicazCatast.PadRight(5, " ")
            'tipo unità
            sDatiTxt += MyRcDettaglio.sIdTipoUnita.PadRight(1, " ")
            'sezione
            sDatiTxt += MyRcDettaglio.sSezione.PadRight(3, " ")
            'foglio
            sDatiTxt += MyRcDettaglio.sFoglio.PadRight(5, " ")
            'numero
            sDatiTxt += MyRcDettaglio.sParticella.PadRight(5, " ")
            'estensione particella
            sDatiTxt += MyRcDettaglio.sEstensioneParticella.PadRight(4, " ")
            'tipo particella
            sDatiTxt += MyRcDettaglio.sIdTipoParticella.PadRight(1, " ")
            'subalterno
            sDatiTxt += MyRcDettaglio.sSubalterno.PadRight(4, " ")
            'Via/Piazza/C.so
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sIndirizzo, 30).PadRight(30, " ")
            'Civico
            sDatiTxt += MyRcDettaglio.sCivico.PadRight(6, " ")
            'Interno
            sDatiTxt += MyRcDettaglio.sInterno.PadRight(2, " ")
            'Scala
            sDatiTxt += MyRcDettaglio.sScala.PadRight(1, " ")
            'Codice assenza dati catastali
            sDatiTxt += MyRcDettaglio.sIdAssenzaDatiCatastali.PadRight(1, " ")
            'Filler
            sDatiTxt += MyRcDettaglio.sFiller.PadRight(78, " ")
            'Carattere di controllo="A"
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            log.Debug("ServiceOPENae::TARSU_FormattaRcDettaglio::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcDettaglio::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcDettaglio::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function TARSU_FormattaRcCoda(ByVal MyRcCoda As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::TARSU_FormattaRcCoda::inizio procedura")

            'tipo record = Vale sempre 9
            sDatiTxt = oConst.TIPORCCODA
            'identificativo fornitura = "SMRIF"
            sDatiTxt += oConst.TARSU_IDFORNITURA
            'Codice numerico fornitura
            sDatiTxt += oConst.TARSU_CODNUMFORNITURA
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcCoda.sCFPIVA.PadRight(16, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sCognome, 26).PadRight(26, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sNome, 25).PadRight(25, " ")
            'sesso
            sDatiTxt += MyRcCoda.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcCoda.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sComuneNascita, 40).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcCoda.sPvNascita.PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sRagSociale, 60).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sComuneSede, 40).PadRight(40, " ")
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

            log.Debug("ServiceOPENae::TARSU_FormattaRcCoda::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcCoda::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::TARSU_FormattaRcCoda::" & Err.Message)
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
            For x = 0 To dvMyDati.Count - 1
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
            If H2O_WriteRcTesta(dvMyDati, sPathCartellaTracciatiCompleto, -1) = False Then
                Return False
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::H2O_EstraiTracciato::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::H2O_EstraiTracciato::" & Err.Message)
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
            oTestaCoda.sProgFornitura = (CStr(dvTesta.Item(0)("anno")) & CStr(dvTesta.Item(0)("id_flusso")).PadLeft(3, "0")).Substring(0, 7)
            'data di invio nel formato GGMMAAAA
            oTestaCoda.sDataInvio = Now.ToString("ddMMyyyy") 'oConst.ReplaceDataForTXT(oConst.ReplaceDataForDB(Now.ToString))
            'filler
            oTestaCoda.sFiller = ""
            If nRcFile = -1 Then
                sDati = H2O_FormattaRcCoda(oTestaCoda)
            Else
                sDati = H2O_FormattaRcTesta(oTestaCoda)
            End If
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
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcTesta::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcTesta::" & Err.Message)
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
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcCoda::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcCoda::" & Err.Message)
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
                oDettaglio.sSesso = ""
                'data di nascita
                oDettaglio.sDataNascita = ""
                'comune di nascita
                oDettaglio.sComuneNascita = ""
                'provincia di nascita
                oDettaglio.sPVNascita = ""
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
                If Not IsDBNull(dvDettaglio.Item(nDett)("comune_sede")) Then
                    oDettaglio.sComuneSede = CStr(dvDettaglio.Item(nDett)("comune_sede"))
                Else
                    oDettaglio.sComuneSede = ""
                End If
                'Provincia Sede
                If Not IsDBNull(dvDettaglio.Item(nDett)("pv_sede")) Then
                    oDettaglio.sPvSede = CStr(dvDettaglio.Item(nDett)("pv_sede"))
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
                If Not IsDBNull(dvDettaglio.Item(nDett)("sesso")) Then
                    oDettaglio.sSesso = CStr(dvDettaglio.Item(nDett)("sesso"))
                Else
                    oDettaglio.sSesso = ""
                End If
                'data nascita
                If Not IsDBNull(dvDettaglio.Item(nDett)("data_nascita")) Then
                    oDettaglio.sDataNascita = oConst.ReplaceDataForTXT(CStr(dvDettaglio.Item(nDett)("data_nascita")))
                Else
                    oDettaglio.sDataNascita = ""
                End If
                'comune nascita
                If Not IsDBNull(dvDettaglio.Item(nDett)("comune_sede")) Then
                    oDettaglio.sComuneNascita = CStr(dvDettaglio.Item(nDett)("comune_sede"))
                Else
                    oDettaglio.sComuneNascita = ""
                End If
                'provincia nascita
                If Not IsDBNull(dvDettaglio.Item(nDett)("pv_sede")) Then
                    oDettaglio.sPVNascita = CStr(dvDettaglio.Item(nDett)("pv_sede"))
                Else
                    oDettaglio.sPVNascita = ""
                End If
                'Denominazione o Ragion Sociale
                oDettaglio.sRagSociale = ""
                'Comune Sede
                oDettaglio.sComuneSede = ""
                'Provincia Sede
                oDettaglio.sPvSede = ""
            End If
            'comune del domicilio fiscale
            If Not IsDBNull(dvDettaglio.Item(nDett)("comune_domiciliofisc")) Then
                oDettaglio.sComuneDomFiscale = CStr(dvDettaglio.Item(nDett)("comune_domiciliofisc"))
            Else
                oDettaglio.sComuneDomFiscale = ""
            End If
            'provincia del domicilio fiscale
            If Not IsDBNull(dvDettaglio.Item(nDett)("pv_domiciliofisc")) Then
                oDettaglio.sPVDomFiscale = CStr(dvDettaglio.Item(nDett)("pv_domiciliofisc"))
            Else
                oDettaglio.sPVDomFiscale = ""
            End If
            'titolo occupazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_titolo_occupazione")) Then
                oDettaglio.sIdTitoloOccupazione = CStr(dvDettaglio.Item(nDett)("id_titolo_occupazione"))
            Else
                oDettaglio.sIdTitoloOccupazione = ""
            End If
            'estremi del contratto
            If Not IsDBNull(dvDettaglio.Item(nDett)("estremi_contratto")) Then
                oDettaglio.sEstremiContratto = CStr(dvDettaglio.Item(nDett)("estremi_contratto"))
            Else
                oDettaglio.sEstremiContratto = ""
            End If
            oDettaglio.stipocontratto = CStr(dvDettaglio.Item(nDett)("tipocontratto"))
            'data inizio occupazione - formato GGMMAAAA
            oDettaglio.sDataInizio = oConst.ReplaceDataForTXT(CStr(dvDettaglio.Item(nDett)("data_inizio")))
            'tipologia utenza
            If Not IsDBNull(dvDettaglio.Item(nDett)("id_tipo_utenza")) Then
                oDettaglio.sIdTipologiaUtenza = CStr(dvDettaglio.Item(nDett)("id_tipo_utenza"))
            Else
                oDettaglio.sIdTipologiaUtenza = ""
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
                'If (dvDettaglio.Item(nDett)("id_assenza_dati_catastali")) > 3 Then
                '    oDettaglio.sIdAssenzaDatiCatastali = CStr(oDettaglio.sIdAssenzaDatiCatastali - 1)
                'End If
                oDettaglio.sIdAssenzaDatiCatastali = CStr(dvDettaglio.Item(nDett)("id_assenza_dati_catastali"))
            Else
                oDettaglio.sIdAssenzaDatiCatastali = ""
            End If
            'numero mesi di fatturazione
            If Not IsDBNull(dvDettaglio.Item(nDett)("mesi_fatturazione")) Then
                oDettaglio.sMesiFatturazione = CStr(dvDettaglio.Item(nDett)("mesi_fatturazione"))
            Else
                oDettaglio.sMesiFatturazione = ""
            End If
            'segno importo
            If Not IsDBNull(dvDettaglio.Item(nDett)("segno_spesa")) Then
                oDettaglio.sSegnoSpesa = CStr(dvDettaglio.Item(nDett)("segno_spesa"))
            Else
                oDettaglio.sSegnoSpesa = ""
            End If
            'spesa consumo
            If Not IsDBNull(dvDettaglio.Item(nDett)("consumo")) Then
                oDettaglio.sConsumo = CStr(dvDettaglio.Item(nDett)("consumo"))
            Else
                oDettaglio.sConsumo = ""
            End If
            'ammontare fatturato
            If Not IsDBNull(dvDettaglio.Item(nDett)("importofatturato")) Then
                oDettaglio.sImportoFatturato = CStr(dvDettaglio.Item(nDett)("importofatturato"))
            Else
                oDettaglio.sImportoFatturato = ""
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
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcDettaglio::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::H2O_WriteRcDettaglio::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function H2O_FormattaRcTesta(ByVal MyRcTesta As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::H2O_FormattaRcTesta::inizio procedura")
            'tipo record = Vale sempre 0
            sDatiTxt = oConst.TIPORCTESTA
            'identificativo fornitura = "NWIDR"
            sDatiTxt += oConst.H2O_IDFORNITURA
            'Codice numerico fornitura=24
            sDatiTxt += oConst.H2O_CODNUMFORNITURA
            'tipo comunicazione 
            sDatiTxt += oConst.H2O_TIPOCOMUNICAZIONE
            'PROT COMUNICAZIONE
            sDatiTxt += oConst.H2O_PROTCOMUNICAZIONE.PadRight(17, " ")
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcTesta.sCFPIVA.PadRight(16, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sRagSociale, 60).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sComuneSede, 40).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcTesta.sPvSede.PadRight(2, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sCognome, 24).PadRight(24, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sNome, 20).PadRight(20, " ")
            'sesso
            sDatiTxt += MyRcTesta.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcTesta.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerTXT(MyRcTesta.sComuneNascita, 40).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcTesta.sPvNascita.PadRight(2, " ")
            'anno di riferimento
            sDatiTxt += MyRcTesta.sAnno.PadRight(4, " ")

            ' CAMPO ELIMINATO PER LA MODIFICA DEL TRACCIATO 
            ''progressivo invio
            'sDatiTxt += MyRcTesta.sProgFornitura
            ' codice intermediario 
            sDatiTxt += oConst.H2O_CFINTERMEDIARIO.PadRight(16, " ")
            ' numero iscrizione CAF
            sDatiTxt += oConst.H2O_NUMCAF.PadRight(5, " ")
            ' codice impegno trasmissione 
            sDatiTxt += oConst.H2O_IMPEGNOTRASMISSIONE
            'data impegno trasmissione 
            sDatiTxt += oConst.H2O_DATAIMPEGNO.PadRight(8, " ")

            ' CAMPO ELIMINATO PER LA MODIFICA DEL TRACCIATO 
            ''data di invio
            ' sDatiTxt += MyRcTesta.sDataInvio.PadRight(8, " ")
            'filler
            sDatiTxt += MyRcTesta.sFiller.PadRight(1524, " ")
            'carattere di controllo
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            log.Debug("ServiceOPENae::H2O_FormattaRcTesta::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcTesta::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcTesta::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function H2O_FormattaRcDettaglio(ByVal MyRcDettaglio As StructureDettaglio) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::H2O_FormattaRcDettaglio::inizio procedura")

            'tipo record = Vale sempre 1
            sDatiTxt = oConst.TIPORCDETTAGLIO
            'Codice Fiscale - Partita IVA
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sCFPIVA, 16).PadRight(16, " ").ToUpper
            'Cognome
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sCognome, 24).PadRight(24, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sNome, 20).PadRight(20, " ")
            'Sesso
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sSesso, 1).PadRight(1, " ")
            'Data di nascita
            sDatiTxt += MyRcDettaglio.sDataNascita.PadRight(8, " ")
            'comune di nascita
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneNascita, 40).PadRight(40, " ")
            'Provincia di nascita
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sPVNascita, 2).PadRight(2, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sRagSociale, 60).PadRight(60, " ").ToUpper
            'eliminato per la modificato del tracciato 
            'Comune Sede
            'sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneSede, 40).PadRight(40, " ")
            ''Provincia Sede
            'sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sPvSede, 2).PadRight(2, " ")
            'comune del domicilio fiscale
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneDomFiscale, 40).PadRight(40, " ")
            'provincia del domicilio fiscale
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sPVDomFiscale, 2).PadRight(2, " ")
            'titolo occupazione
            sDatiTxt += MyRcDettaglio.sIdTitoloOccupazione
            'estremi contratto
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sEstremiContratto, 30).PadRight(30, " ")
            'tipologia utenza
            sDatiTxt += MyRcDettaglio.sIdTipologiaUtenza
            ' tipo di contratto 
            sDatiTxt += MyRcDettaglio.sTIPOCONTRATTO
            'data prima attivazione - formato GGMMAAAA
            sDatiTxt += MyRcDettaglio.sDataInizio.PadRight(8, " ")
            ' eliminato per la modifica del tracciato record
            'comune amministrativo di ubicazione dell'immobile
            'sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneAmmUbicazione, 20).PadRight(20, " ")
            'provincia di ubicazione dell'immobile
            'sDatiTxt += MyRcDettaglio.sPVUbicazione.PadRight(2, " ")
            'comune catastale di ubicazione dell'immobile
            'sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sComuneCatUbicazione, 20).PadRight(20, " ")
            'Via/Piazza/C.so
            sDatiTxt += oConst.FormattaPerTXT(MyRcDettaglio.sIndirizzo, 40).PadRight(40, " ").ToUpper
            'codice comune catastale
            sDatiTxt += MyRcDettaglio.sCodComuneUbicazCatast.PadRight(4, " ")
            'tipo unità
            sDatiTxt += MyRcDettaglio.sIdTipoUnita
            'Codice assenza dati catastali
            sDatiTxt += MyRcDettaglio.sIdAssenzaDatiCatastali.PadRight(1, " ")
            'sezione
            sDatiTxt += MyRcDettaglio.sSezione.PadRight(3, " ")
            'foglio
            sDatiTxt += MyRcDettaglio.sFoglio.PadRight(5, " ")
            'numero
            sDatiTxt += MyRcDettaglio.sParticella.PadRight(5, " ")
            'subalterno
            sDatiTxt += MyRcDettaglio.sSubalterno.PadRight(4, " ")
            'estensione particella
            sDatiTxt += MyRcDettaglio.sEstensioneParticella.PadRight(4, " ")
            'tipo particella
            sDatiTxt += MyRcDettaglio.sIdTipoParticella.PadRight(1, " ")
            'numero di mesi di fatturazione
            sDatiTxt += MyRcDettaglio.sMesiFatturazione.PadLeft(2, " ")
            ''segno spesa
            'sDatiTxt += MyRcDettaglio.sSegnoSpesa.PadRight(1, " ")
            'consumo fatturato
            sDatiTxt += MyRcDettaglio.sConsumo.PadLeft(9, "0")
            'ammontare fatturato
            sDatiTxt += MyRcDettaglio.sImportoFatturato.PadLeft(9, "0")
            'Filler
            sDatiTxt += MyRcDettaglio.sFiller.PadRight(1454, " ")
            'Carattere di controllo="A"
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************
            log.Debug("ServiceOPENae::H2O_FormattaRcDettaglio::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcDettaglio::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcDettaglio::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function H2O_FormattaRcCoda(ByVal MyRcCoda As StructureTestaCoda) As String
        Dim sDatiTxt As String = ""

        Try
            log.Debug("ServiceOPENae::H2O_FormattaRcCoda::inizio procedura")

            'tipo record = Vale sempre 9
            sDatiTxt = oConst.TIPORCCODA
            'identificativo fornitura = "NWIDR"
            sDatiTxt += oConst.H2O_IDFORNITURA
            'Codice numerico fornitura=24
            sDatiTxt += oConst.H2O_CODNUMFORNITURA
            'tipo comunicazione 
            sDatiTxt += oConst.H2O_TIPOCOMUNICAZIONE
            'PROT COMUNICAZIONE
            sDatiTxt += oConst.H2O_PROTCOMUNICAZIONE.PadRight(17, " ")
            'Codice Fiscale - Partita IVA
            sDatiTxt += MyRcCoda.sCFPIVA.PadRight(16, " ")
            'Denominazione o Ragion Sociale
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sRagSociale, 60).PadRight(60, " ")
            'Comune Sede
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sComuneSede, 40).PadRight(40, " ")
            'Provincia Sede
            sDatiTxt += MyRcCoda.sPvSede.PadRight(2, " ")
            'Cognome
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sCognome, 24).PadRight(24, " ")
            'Nome
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sNome, 20).PadRight(20, " ")
            'sesso
            sDatiTxt += MyRcCoda.sSesso.PadRight(1, " ")
            'data nascita
            sDatiTxt += MyRcCoda.sDataNascita.PadRight(8, " ")
            'comune o stato estero di nascita
            sDatiTxt += oConst.FormattaPerTXT(MyRcCoda.sComuneNascita, 40).PadRight(40, " ")
            'provincia di nascita
            sDatiTxt += MyRcCoda.sPvNascita.PadRight(2, " ")
            'anno di riferimento
            sDatiTxt += MyRcCoda.sAnno.PadRight(4, " ")

            ' CAMPO ELIMINATO PER LA MODIFICA DEL TRACCIATO 
            ''progressivo invio
            'sDatiTxt += MyRcCoda.sProgFornitura
            ' codice intermediario 
            sDatiTxt += oConst.H2O_CFINTERMEDIARIO.PadRight(16, " ")
            ' numero iscrizione CAF
            sDatiTxt += oConst.H2O_NUMCAF.PadRight(5, " ")
            ' codice impegno trasmissione 
            sDatiTxt += oConst.H2O_IMPEGNOTRASMISSIONE
            'data impegno trasmissione 
            sDatiTxt += oConst.H2O_DATAIMPEGNO.PadRight(8, " ")

            ' CAMPO ELIMINATO PER LA MODIFICA DEL TRACCIATO 
            ''data di invio
            ' sDatiTxt += MyRcCoda.sDataInvio.PadRight(8, " ")
            'filler
            sDatiTxt += MyRcCoda.sFiller.PadRight(1524, " ")
            'carattere di controllo
            sDatiTxt += oConst.CHRCONTROLLO
            '***per come scriviamo il file il carattere di a capo viene già messo in automatico***
            'Carattere ASCII "CR" - "LF"
            'sDatiTxt += oConst.CHRASCIIFINERIGA
            '**********************************************************************

            log.Debug("ServiceOPENae::H2O_FormattaRcCoda::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcCoda::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::H2O_FormattaRcCoda::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function
#End Region

#Region "ICI"
    Private Function ICI_EstraiTracciato(ByVal oDBManager As Utility.DBModel, ByVal dvMyDati As DataView, ByVal sCodIstat As String, ByVal sBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal sTributo As String, ByVal sAnno As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Try
            Dim oQueryManager As New ClsInterDB
            Dim x As Integer
            Dim nTotRiversato As Double
            Dim nDisposizioni, nTipoRiversamento, nProgRc, nFlagReperibilita As Integer
            Dim nNumQuietanzaPrec As Integer = -1
            Dim nTipoRiversamentoprec As Integer
            Dim sTipoRiscossioni As String
            Dim sAnnoPrec, sDataRiversamentoPrec As String
            Dim nTotRc1 As Double = 0
            Dim nTotRc3 As Double = 0
            Dim nTotRc45 As Double = 0
            Dim nTotRc6 As Double = 0
            Dim bIsViolazione As Boolean = False

            For x = 0 To dvMyDati.Count - 1
                If sAnnoPrec <> dvMyDati.Item(x)("anno") Then
                    If sAnnoPrec <> "" Then
                        '***************************************************
                        'scrivo il record di coda
                        '***************************************************
                        If ICI_WriteRcCoda(oDBManager, sAnnoPrec, sDataScadenza, nProgInvio, nTotRc1, nTotRc3, nTotRc45, nTotRc6, sPathCartellaTracciatiCompleto) = False Then
                            Return False
                        End If
                        nTotRc1 = 0 : nTotRc3 = 0 : nTotRc45 = 0 : nTotRc6 = 0
                    End If
                    '***************************************************
                    'scrivo il record di testa
                    '***************************************************
                    If ICI_WriteRcTesta(dvMyDati.Item(x)("anno"), sDataScadenza, nProgInvio, sPathCartellaTracciatiCompleto) = False Then
                        Return False
                    End If
                End If
                bIsViolazione = False
                sAnnoPrec = dvMyDati.Item(x)("anno")
                Select Case dvMyDati.Item(x)("provenienza")
                    Case oConst.ICI_RIVERSAMENTOCONTOCORRENTE
                        nTipoRiversamento = 2
                    Case Else
                        nTipoRiversamento = 0
                End Select
                'controllo se devo inserire un record di riversamento
                If CInt(dvMyDati.Item(x)("numero_quietanza")) <> nNumQuietanzaPrec Or CStr(dvMyDati.Item(x)("data_accredito")) <> sDataRiversamentoPrec Or nTipoRiversamento <> nTipoRiversamentoprec Then
                    If oQueryManager.GetTotaliInvioICI(oDBManager, sCodIstat, dvMyDati.Item(x)("anno"), CInt(dvMyDati.Item(x)("numero_quietanza")), CStr(dvMyDati.Item(x)("data_accredito")), nTipoRiversamento, sTipoRiscossioni, nTotRiversato, nDisposizioni) = False Then
                        Return False
                    End If
                    '***************************************************
                    'scrivo il record di riversamento
                    '***************************************************
                    nProgRc += 1
                    If ICI_WriteRcRiversamento(sBelfiore, CStr(dvMyDati.Item(x)("data_accredito")), CInt(dvMyDati.Item(x)("numero_quietanza")), nTotRiversato, nDisposizioni, nTipoRiversamento, sTipoRiscossioni, nProgRc, sPathCartellaTracciatiCompleto) = False Then
                        Return False
                    End If
                    nTotRc1 += 1
                    nNumQuietanzaPrec = CInt(dvMyDati.Item(x)("numero_quietanza"))
                    sDataRiversamentoPrec = CStr(dvMyDati.Item(x)("data_accredito"))
                    nTipoRiversamentoprec = nTipoRiversamento
                End If
                nFlagReperibilita = CheckCinCFPIva(dvMyDati.Item(x)("cf_piva"))
                If nFlagReperibilita = -1 Then
                    Return False
                End If
                '***************************************************
                'scrivo il record contabile e anagrafico
                '***************************************************
                If Not IsDBNull(dvMyDati.Item(x)("n_sanzione")) Then
                    If CStr(dvMyDati.Item(x)("n_sanzione")) <> "" Then
                        bIsViolazione = True
                    End If
                End If
                If Not IsDBNull(dvMyDati.Item(x)("data_sanzione")) Then
                    If CStr(dvMyDati.Item(x)("data_sanzione")) <> "" Then
                        bIsViolazione = True
                    End If
                End If
                If Not IsDBNull(dvMyDati.Item(x)("tipo_bollettino_violazioni")) Then
                    If CStr(dvMyDati.Item(x)("tipo_bollettino_violazioni")) <> "" Then
                        bIsViolazione = True
                    End If
                End If
                'valorizzo il progressivo record
                nProgRc += 1
                If ICI_WriteRcContabile(dvMyDati, sBelfiore, sDescrEnte, sCAPEnte, nProgRc, x, nTipoRiversamento, nFlagReperibilita, nTotRc3, nTotRc6, sPathCartellaTracciatiCompleto) = False Then
                    Return False
                End If
                If nFlagReperibilita = 1 Then
                    'devo scrivere prima il contabile e poi l'anagrafico
                    If ICI_WriteRcAnagrafico(dvMyDati, sBelfiore, nProgRc, x, sPathCartellaTracciatiCompleto) = False Then
                        Return False
                    End If
                    nTotRc45 += 1
                End If
            Next

            '***************************************************
            'scrivo il record di coda
            '***************************************************
            If ICI_WriteRcCoda(oDBManager, sAnnoPrec, sDataScadenza, nProgInvio, nTotRc1, nTotRc3, nTotRc45, nTotRc6, sPathCartellaTracciatiCompleto) = False Then
                Return False
            End If

            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ICI_EstraiTracciato::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ICI_EstraiTracciato::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ICI_WriteRcTesta(ByVal sAnno As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Dim oTesta As New ExportTestaICI
        Dim oQueryManager As New ClsInterDB

        Try
            'codice concessione
            oTesta.sCodConcessione = oConst.ICI_CODCONCESSIONE
            'periodo di riferimento
            oTesta.sPeriodoRifRiscossioni = sAnno
            'data di scadenza
            oTesta.sDataScadenza = sDataScadenza
            'progressivo invio
            oTesta.nProgressivoInvio = nProgInvio
            'numero di supporti
            oTesta.nNumSupporti = ConfigurationSettings.AppSettings("TipoSupportoTracciato") '"1"
            'numero d'ordine del supporto
            oTesta.nNumOrdineSupporto = ConfigurationSettings.AppSettings("TipoSupportoTracciato") '"1"
            'filler
            oTesta.sFiller = "0"
            sDati = ICI_FormattaRcTesta(oTesta)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcTesta::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcTesta::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ICI_WriteRcCoda(ByVal oDBManager As Utility.DBModel, ByVal sAnno As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByVal nRc1 As Double, ByVal nRc3 As Double, ByVal nRc45 As Double, ByVal nRc6 As Double, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Dim oCoda As New ExportCodaICI
        Dim oQueryManager As New ClsInterDB

        Try
            'codice concessione
            oCoda.sCodConcessione = oConst.ICI_CODCONCESSIONE
            'periodo di riferimento
            oCoda.sPeriodoRifRiscossioni = sAnno
            'data di scadenza
            oCoda.sDataScadenza = sDataScadenza
            'progressivo invio
            oCoda.nProgressivoInvio = nProgInvio
            'numero di record di tipo 1
            oCoda.nTotRc1 = nRc1
            'numero di record di tipo 3
            oCoda.nTotRc3 = nRc3
            'numero di record di tipo 4/5
            oCoda.nTotRc45 = nRc45
            'numero di record di tipo 6
            oCoda.nTotRc6 = nRc6
            'filler
            oCoda.sFiller = "0"
            sDati = ICI_FormattaRcCoda(oCoda)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcCoda::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcCoda::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ICI_WriteRcRiversamento(ByVal sBelfiore As String, ByVal sDataRiversamento As String, ByVal nNumQuietanza As Integer, ByVal nTotRiversato As Double, ByVal nDisposizioni As Integer, ByVal nTipoRiversamento As Integer, ByVal sTipoRiscossioni As String, ByVal nProgRc As Integer, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Dim oRiversamento As New ExportRiversamentoICI

        Try
            'codice concessione
            oRiversamento.sCodConcessione = oConst.ICI_CODCONCESSIONE
            'codice ente
            oRiversamento.sCodBelfiore = sBelfiore
            'numero quietanza
            oRiversamento.nNumQuietanza = nNumQuietanza
            'progressivo record
            oRiversamento.nProgressivoRc = nProgRc
            'data riversamento
            oRiversamento.sDataRiversamento = sDataRiversamento
            'codice tesoreria
            oRiversamento.nCodTesoreria = oConst.ICI_CODTESORERIA
            'importo riversato
            oRiversamento.nTotImpRiversato = nTotRiversato * 100
            'commissione
            oRiversamento.nCommissione = 0
            'numero di riscossioni
            oRiversamento.nTotNumRiscossioni = nDisposizioni
            'flag tipo versamento
            oRiversamento.nTipoRiversamento = nTipoRiversamento
            'tipologia riscossioni
            oRiversamento.sTipoRiscossioni = sTipoRiscossioni
            'filler
            oRiversamento.sFiller = "0"
            sDati = ICI_FormattaRcRiversamento(oRiversamento)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcRiversamento::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcRiversamento::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ICI_WriteRcContabile(ByVal dvContabile As DataView, ByVal sBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal nProgRc As Integer, ByVal nRc As Integer, ByVal nTipoRiversamento As Integer, ByVal nReperibilita As Integer, ByRef nRc3 As Double, ByRef nRc6 As Double, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Dim oContabile As New ExportContabileICI
        Dim oQueryManager As New ClsInterDB
        Dim nTotRiversato As Double
        Dim nDisposizioni, nTipoContabile As Integer

        Try
            'tipo versamento: violazione o ordinario
            If Not IsDBNull(dvContabile.Item(nRc)("n_sanzione")) Then
                If CStr(dvContabile.Item(nRc)("n_sanzione")) <> "" Then
                    oContabile.bIsViolazione = True
                End If
            End If
            If Not IsDBNull(dvContabile.Item(nRc)("data_sanzione")) Then
                If CStr(dvContabile.Item(nRc)("data_sanzione")) <> "" Then
                    oContabile.bIsViolazione = True
                End If
            End If
            If Not IsDBNull(dvContabile.Item(nRc)("tipo_bollettino_violazioni")) Then
                If CStr(dvContabile.Item(nRc)("tipo_bollettino_violazioni")) <> "" Then
                    oContabile.bIsViolazione = True
                End If
            End If
            'codice concessione
            oContabile.sCodConcessione = oConst.ICI_CODCONCESSIONE
            'codice ente
            oContabile.sCodBelfiore = sBelfiore
            'numero quietanza
            oContabile.nNumQuietanza = dvContabile.Item(nRc)("numero_quietanza")
            'progressivo record
            oContabile.nProgressivoRc = nProgRc
            'data versamento
            oContabile.sDataVersamento = dvContabile.Item(nRc)("data_pagamento")
            'codice fiscale
            oContabile.sCFPIVA = dvContabile.Item(nRc)("cf_piva")
            'periodo di riferimento/anno di imposta
            oContabile.nAnnoImposta = dvContabile.Item(nRc)("anno")
            'numero di riferimento quietanza
            If Not IsDBNull(dvContabile.Item(nRc)("n_movimento")) Then
                oContabile.sNumRifQuietanza = dvContabile.Item(nRc)("n_movimento")
            End If
            'importo riversato
            oContabile.nImpVersato = dvContabile.Item(nRc)("importo") * 100
            log.Info("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcContabile::Importo Versato da DB: " & dvContabile.Item(nRc)("importo"))
            log.Info("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcContabile::Importo Versato per file: " & oContabile.nImpVersato)
            'importo terreni agricoli
            oContabile.nImpTerAgr = dvContabile.Item(nRc)("imp_ter_agr") * 100
            'importo aree fabbicabili
            oContabile.nImpAreeFab = dvContabile.Item(nRc)("imp_aree_fab") * 100
            'importo abitazione principale
            oContabile.nImpAbiPrinc = dvContabile.Item(nRc)("imp_abi_prin") * 100
            'importo altri fabbricati
            oContabile.nImpAltriFab = dvContabile.Item(nRc)("imp_altri_fab") * 100
            'importo detrazione
            oContabile.nImpDetrazione = dvContabile.Item(nRc)("detrazione") * 100
            'flag quadratura
            oContabile.nQuadratura = CheckImporti(oContabile)
            If oContabile.nQuadratura = -1 Then
                Return False
            End If
            'flag reperibilità
            oContabile.nReperibilita = nReperibilita
            'tipo versamento
            oContabile.nTipoVersamento = nTipoRiversamento
            'data registrazione
            oContabile.sDataAccredito = dvContabile.Item(nRc)("data_accredito")
            'flag di competenza
            oContabile.nCompentenza = 0
            'comune
            oContabile.sComune = sDescrEnte
            'cap
            oContabile.sCAP = sCAPEnte
            'numero fabbricati
            oContabile.nNumFab = dvContabile.Item(nRc)("n_fab")
            'flag acconto/saldo
            oContabile.nFlagAS = dvContabile.Item(nRc)("flag_acconto_saldo")
            'flag identificazione
            oContabile.nIdentificazione = 0
            'ravvedimento
            If CStr(dvContabile.Item(nRc)("flag_ravvedimento_operoso")) <> "" Then
                oContabile.nRavvedimento = dvContabile.Item(nRc)("flag_ravvedimento_operoso")
            End If
            'numero provvedimento liquidazione
            oContabile.sNumProvLiq = dvContabile.Item(nRc)("n_sanzione")
            'data provvedimento liquidazione
            oContabile.sDataProvLiq = dvContabile.Item(nRc)("data_sanzione")
            'filler
            oContabile.sFiller = "0"
            sDati = ICI_FormattaRcContabile(oContabile)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            If oContabile.bIsViolazione = True Then
                nRc6 += 1
            Else
                nRc3 += 1
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcContabile::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcContabile::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ICI_WriteRcAnagrafico(ByVal dvAnagrafico As DataView, ByVal sBelfiore As String, ByVal nProgRc As Integer, ByVal nRc As Integer, ByVal sPathCartellaTracciatiCompleto As String) As Boolean
        Dim oAnagrafico As New ExportAnagraficoICI
        Dim oQueryManager As New ClsInterDB
        Dim nTotRiversato As Double
        Dim nDisposizioni, nTipoAnagrafico As Integer

        Try
            'codice concessione
            oAnagrafico.sCodConcessione = oConst.ICI_CODCONCESSIONE
            'codice ente
            oAnagrafico.sCodBelfiore = sBelfiore
            'numero quietanza
            oAnagrafico.nNumQuietanza = dvAnagrafico.Item(nRc)("numero_quietanza")
            'progressivo record
            oAnagrafico.nProgressivoRc = nProgRc
            'cognome/ragione sociale
            oAnagrafico.sCognomeRagSoc = dvAnagrafico.Item(nRc)("cognome")
            'nome
            If CStr(dvAnagrafico.Item(nRc)("cf_piva")).Length = 11 Or CStr(dvAnagrafico.Item(nRc)("sesso")) = "G" Then
                oAnagrafico.sSesso = "G"
            Else
                oAnagrafico.sNome = dvAnagrafico.Item(nRc)("nome")
                oAnagrafico.sSesso = "F"
            End If
            'comune
            oAnagrafico.sComune = dvAnagrafico.Item(nRc)("citta_res")
            'filler
            oAnagrafico.sFiller = "0"
            sDati = ICI_FormattaRcAnagrafico(oAnagrafico)
            If sDati <> "" Then
                If WriteFile(sPathCartellaTracciatiCompleto, sDati, sMsgErr) < 1 Then
                    Return False
                End If
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcAnagrafico::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::EstraiTracciato::ICI_WriteRcAnagrafico::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ICI_FormattaRcTesta(ByVal MyRcTesta As ExportTestaICI) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::ICI_FormattaRcTesta::inizio procedura")
            'tipo record = Vale sempre ICI0
            sDatiTxt = oConst.ICI_TIPORCTESTA
            'codice concessione
            sDatiTxt += MyRcTesta.sCodConcessione.PadRight(3, " ")
            'periodo di riferimento
            sDatiTxt += MyRcTesta.sPeriodoRifRiscossioni.PadRight(4, " ")
            'data di scadenza
            sDatiTxt += MyRcTesta.sDataScadenza.PadRight(8, "0")
            'progressivo invio
            sDatiTxt += MyRcTesta.nProgressivoInvio.ToString.PadLeft(2, "0")
            'numero di supporti
            sDatiTxt += MyRcTesta.nNumSupporti.ToString.PadLeft(2, "0")
            'numero d'ordine del supporto
            sDatiTxt += MyRcTesta.nNumOrdineSupporto.ToString.PadLeft(2, "0")
            'filler
            sDatiTxt += MyRcTesta.sFiller.PadLeft(175, "0")

            log.Debug("ServiceOPENae::ICI_FormattaRcTesta::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcTesta::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcTesta::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function ICI_FormattaRcCoda(ByVal MyRcCoda As ExportCodaICI) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::ICI_FormattaRcCoda::inizio procedura")
            'tipo record = Vale sempre ICI0
            sDatiTxt = oConst.ICI_TIPORCCODA
            'codice concessione
            sDatiTxt += MyRcCoda.sCodConcessione.PadRight(3, " ")
            'periodo di riferimento
            sDatiTxt += MyRcCoda.sPeriodoRifRiscossioni.PadRight(4, " ")
            'data di scadenza
            sDatiTxt += MyRcCoda.sDataScadenza.PadRight(8, "0")
            'progressivo invio
            sDatiTxt += MyRcCoda.nProgressivoInvio.ToString.PadLeft(2, "0")
            'numero di record di tipo 1
            sDatiTxt += MyRcCoda.nTotRc1.ToString.PadLeft(10, "0")
            'numero di record di tipo 3
            sDatiTxt += MyRcCoda.nTotRc3.ToString.PadLeft(10, "0")
            'numero di record di tipo 4/5
            sDatiTxt += MyRcCoda.nTotRc45.ToString.PadLeft(10, "0")
            'numero di record di tipo 6
            sDatiTxt += MyRcCoda.nTotRc6.ToString.PadLeft(10, "0")
            'filler
            sDatiTxt += MyRcCoda.sFiller.PadLeft(139, "0")

            log.Debug("ServiceOPENae::ICI_FormattaRcCoda::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcCoda::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcCoda::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function ICI_FormattaRcRiversamento(ByVal MyRcRiversamento As ExportRiversamentoICI) As String
        Dim sDatiTxt As String = ""
        Try
            log.Debug("ServiceOPENae::ICI_FormattaRcRiversamento::inizio procedura")
            'codice concessione
            sDatiTxt += MyRcRiversamento.sCodConcessione.PadRight(3, " ")
            'codice ente
            sDatiTxt += MyRcRiversamento.sCodBelfiore.PadRight(4, " ")
            'numero quietanza
            sDatiTxt += MyRcRiversamento.nNumQuietanza.ToString.PadLeft(10, "0")
            'progressivo record
            sDatiTxt += MyRcRiversamento.nProgressivoRc.ToString.PadLeft(8, "0")
            'tipo record = Vale sempre 1
            sDatiTxt += oConst.TIPORCDETTAGLIO
            'data riversamento
            sDatiTxt += MyRcRiversamento.sDataRiversamento.PadRight(8, "0")
            'codice tesoreria
            sDatiTxt += MyRcRiversamento.nCodTesoreria.ToString.PadLeft(3, "0")
            'importo riversato
            sDatiTxt += MyRcRiversamento.nTotImpRiversato.ToString.PadLeft(13, "0")
            'commissione
            sDatiTxt += MyRcRiversamento.nCommissione.ToString.PadLeft(10, "0")
            'numero di riscossioni
            sDatiTxt += MyRcRiversamento.nTotNumRiscossioni.ToString.PadLeft(6, "0")
            'flag tipo versamento
            sDatiTxt += MyRcRiversamento.nTipoRiversamento.ToString.PadLeft(1, "0")
            'tipologia riscossioni
            sDatiTxt += MyRcRiversamento.sTipoRiscossioni.PadRight(1, " ")
            'filler
            sDatiTxt += MyRcRiversamento.sFiller.PadLeft(132, "0")

            log.Debug("ServiceOPENae::ICI_FormattaRcRiversamento::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcRiversamento::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcRiversamento::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function ICI_FormattaRcContabile(ByVal MyRcContabile As ExportContabileICI) As String
        Dim sDatiTxt As String = ""

        Try
            log.Debug("ServiceOPENae::ICI_FormattaRcContabile::inizio procedura")
            'codice concessione
            sDatiTxt += MyRcContabile.sCodConcessione.PadRight(3, " ")
            'codice ente
            sDatiTxt += MyRcContabile.sCodBelfiore.PadRight(4, " ")
            'numero quietanza
            sDatiTxt += MyRcContabile.nNumQuietanza.ToString.PadLeft(10, "0")
            'progressivo record
            sDatiTxt += MyRcContabile.nProgressivoRc.ToString.PadLeft(8, "0")
            'tipo record = Vale sempre 3-ordinario, 6-violazioni
            If MyRcContabile.bIsViolazione = False Then
                sDatiTxt += oConst.ICI_TIPORCCONTABILEORD
            Else
                sDatiTxt += oConst.ICI_TIPORCCONTABILEVIOL
            End If
            'data versamento
            sDatiTxt += MyRcContabile.sDataVersamento.PadRight(8, " ")
            'codice fiscale
            sDatiTxt += MyRcContabile.sCFPIVA.PadRight(16, " ")
            If MyRcContabile.bIsViolazione = False Then
                'periodo di riferimento/anno di imposta
                sDatiTxt += MyRcContabile.nAnnoImposta.ToString.Substring(2, 2).PadLeft(2, "0")
            Else
                sDatiTxt += MyRcContabile.sFiller.PadLeft(2, "0")
            End If
            'numero di riferimento quietanza
            sDatiTxt += MyRcContabile.sNumRifQuietanza.PadLeft(11, "0")
            'importo riversato
            sDatiTxt += MyRcContabile.nImpVersato.ToString.PadLeft(11, "0")
            If MyRcContabile.bIsViolazione = False Then
                'importo terreni agricoli
                sDatiTxt += MyRcContabile.nImpTerAgr.ToString.PadLeft(10, "0")
                'importo aree fabbicabili
                sDatiTxt += MyRcContabile.nImpAreeFab.ToString.PadLeft(10, "0")
                'importo abitazione principale
                sDatiTxt += MyRcContabile.nImpAbiPrinc.ToString.PadLeft(10, "0")
                'importo altri fabbricati
                sDatiTxt += MyRcContabile.nImpAltriFab.ToString.PadLeft(10, "0")
                'importo detrazione
                sDatiTxt += MyRcContabile.nImpDetrazione.ToString.PadLeft(8, "0")
            Else
                sDatiTxt += MyRcContabile.sFiller.PadLeft(48, "0")
            End If
            'quadratura
            sDatiTxt += MyRcContabile.nQuadratura.ToString.PadLeft(1, "0")
            'flag reperibilità
            sDatiTxt += MyRcContabile.nReperibilita.ToString.PadLeft(1, "0")
            'tipo versamento
            sDatiTxt += MyRcContabile.nTipoVersamento.ToString.PadLeft(1, "0")
            'data registrazione
            sDatiTxt += MyRcContabile.sDataAccredito.PadRight(8, "0")
            'flag di competenza
            sDatiTxt += MyRcContabile.nCompentenza.ToString.PadLeft(1, "0")
            'comune
            sDatiTxt += MyRcContabile.sComune.PadRight(25, " ")
            'cap
            sDatiTxt += MyRcContabile.sCAP.PadRight(5, " ")
            If MyRcContabile.bIsViolazione = False Then
                'numero fabbricati
                sDatiTxt += MyRcContabile.nNumFab.ToString.PadLeft(4, "0")
                'flag acconto/saldo
                sDatiTxt += MyRcContabile.nFlagAS.ToString.PadLeft(1, "0")
            Else
                sDatiTxt += MyRcContabile.sFiller.PadLeft(5, "0")
            End If
            'flag identificazione
            sDatiTxt += MyRcContabile.nIdentificazione.ToString.PadLeft(1, "0")
            If MyRcContabile.bIsViolazione = False Then
                'periodo di riferimento/anno di imposta
                sDatiTxt += MyRcContabile.nAnnoImposta.ToString.PadLeft(4, "0")
                'ravvedimento
                sDatiTxt += MyRcContabile.nRavvedimento.ToString.PadLeft(1, "0")
                'filler
                sDatiTxt += MyRcContabile.sFiller.PadLeft(25, "0")
            Else
                'flag tipo imposta
                sDatiTxt += "1"
                'numero provvedimento liquidazione
                sDatiTxt += MyRcContabile.sNumProvLiq.PadLeft(9, "0")
                'data provvedimento liquidazione
                sDatiTxt += MyRcContabile.sDataProvLiq.PadLeft(8, "0")
                'filler
                sDatiTxt += MyRcContabile.sFiller.PadLeft(12, "0")
            End If

            log.Debug("ServiceOPENae::ICI_FormattaRcContabile::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcContabile::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcContabile::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function ICI_FormattaRcAnagrafico(ByVal MyRcAnagrafico As ExportAnagraficoICI) As String
        Dim sDatiTxt As String = ""

        Try
            log.Debug("ServiceOPENae::ICI_FormattaRcAnagrafico::inizio procedura")
            'codice concessione
            sDatiTxt += MyRcAnagrafico.sCodConcessione.PadRight(3, " ")
            'codice ente
            sDatiTxt += MyRcAnagrafico.sCodBelfiore.PadRight(4, " ")
            'numero quietanza
            sDatiTxt += MyRcAnagrafico.nNumQuietanza.ToString.PadLeft(10, "0")
            'progressivo record
            sDatiTxt += MyRcAnagrafico.nProgressivoRc.ToString.PadLeft(8, "0")
            'tipo record = Vale sempre 4-persona fisica, 5-persona giuridica
            If MyRcAnagrafico.sSesso = "F" Then
                sDatiTxt += oConst.ICI_TIPORCANAGRAFICOFIS
                'cognome
                sDatiTxt += MyRcAnagrafico.sCognomeRagSoc.PadRight(24, " ")
                'nome
                sDatiTxt += MyRcAnagrafico.sNome.PadRight(20, " ")
            Else
                sDatiTxt += oConst.ICI_TIPORCANAGRAFICOGIUR
                'ragione sociale
                sDatiTxt += MyRcAnagrafico.sCognomeRagSoc.PadRight(60, " ")
            End If
            'comune
            sDatiTxt += MyRcAnagrafico.sComune.PadRight(25, " ")
            'filler
            If MyRcAnagrafico.sSesso = "F" Then
                sDatiTxt += MyRcAnagrafico.sFiller.PadLeft(105, "0")
            Else
                sDatiTxt += MyRcAnagrafico.sFiller.PadLeft(89, "0")
            End If

            log.Debug("ServiceOPENae::ICI_FormattaRcAnagrafico::fine procedura")
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcAnagrafico::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ICI_FormattaRcAnagrafico::" & Err.Message)
            sDatiTxt = ""
        End Try
        Return sDatiTxt
    End Function

    Private Function CheckCinCFPIva(ByVal sCFPIva As String) As Integer
        Try
            Select Case sCFPIva.Length
                Case 16
                    If oConst.ControlloCinCF(sCFPIva) = False Then
                        Return 1
                    End If
                Case 11
                    If oConst.ControlloCinPIVA(sCFPIva) = False Then
                        Return 1
                    End If
                Case Else
                    Return 1
            End Select

            Return 0
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::CheckCinCFPIva::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::CheckCinCFPIva::" & Err.Message)
            Return -1
        End Try
    End Function

    Private Function CheckImporti(ByVal oMyContabile As ExportContabileICI) As Integer
        Try
            If oMyContabile.nImpVersato <> (oMyContabile.nImpTerAgr + oMyContabile.nImpAreeFab + oMyContabile.nImpAbiPrinc + oMyContabile.nImpAltriFab) Then
                Return 1
            Else
                Return 0
            End If
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::CheckImporti::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::CheckImporti::" & Err.Message)
            Return -1
        End Try
    End Function

    Private Function SetFineEstrazione() As Boolean
        Try
            '          SQL_inserimento = "INSERT INTO TABELLA_FLUSSI_TRACCIATI_ICI(COD_ENTE,DATA_ESTRAZIONE,DATA_SCADENZA,ANNO_IMPOSTA,NOME_FILE,OPERATORE,PROGRESSIVO_INVIO,IMPORTO_TOTALE_VERSATO,IMPORTO_TOTALE_AREE_FAB,IMPORTO_TOTALE_ABIT_PRINC,IMPORTO_TOTALE_TER_AGR,IMPORTO_TOTALE_ALTRI_FAB,IMPORTO_TOTALE_DETRAZ,N_VERSAMENTI,N_RECORD_TIPO_1,N_RECORD_TIPO_3,N_RECORD_TIPO_4,N_RECORD_TIPO_5,ID_AMBIENTE)VALUES('" & codice_comune & "','" & dData & "','" & Data_di_scadenza & "','" & Periodo_di_riferimento_riscossioni & "','" & nome_file & "','" & nome_utente12 & "'," & Progressivo_inv & "," & conta_importo & "," & terreni_fab & "," & importo_abit & "," & terreni_agric & "," & importo_altri & "," & importo_detra & "," & Progressivo_record & "," & contatore_record1 & "," & contatore_record3 & "," & contatore_record4 & "," & contatore_record5 & "," & codice_ambiente & ")"
            '          record_insert.Open(SQL_inserimento, connessione_cs)
            '          tipo_estrazione = Request.Item("valore")
            '          'Eseguo la query
            '          dvDati.Close()

            '          If tipo_estrazione <> 0 Then
            'SQL_CNC = "UPDATE VERSAMENTI_ICI INNER JOIN ANAGRAFE ON VERSAMENTI_ICI.COD_CONTRIBUENTE = ANAGRAFE.COD_CONTRIBUENTE SET VERSAMENTI_ICI.FLAG_ESTRATTO = 1 , VERSAMENTI_ICI.DATA_ESTRAZIONE = #"&ddData&"# "
            'SQL_CNC = SQL_CNC & " WHERE ((VERSAMENTI_ICI.FLAG_ESTRATTO) Is Null) AND ((VERSAMENTI_ICI.ANNO)="&AnnoRif&") OR ((VERSAMENTI_ICI.FLAG_ESTRATTO)=0) AND ((VERSAMENTI_ICI.ANNO)="&AnnoRif&") "
            '          Else
            '              SQL_CNC = "UPDATE VERSAMENTI_ICI INNER JOIN ANAGRAFE ON VERSAMENTI_ICI.COD_CONTRIBUENTE = ANAGRAFE.COD_CONTRIBUENTE SET VERSAMENTI_ICI.FLAG_ESTRATTO = 1 ,"
            'SQL_CNC = SQL_CNC & " VERSAMENTI_ICI.DATA_ESTRAZIONE = #"&ddData&"# WHERE  ((VERSAMENTI_ICI.ANNO)="&AnnoRif&")"
            '          End If

            '          'Response.end
            '          dvDati.Open(SQL_CNC, connessione_ente)

        Catch ex As Exception

        End Try
    End Function
#End Region
End Class
