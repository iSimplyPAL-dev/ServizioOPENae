Imports log4net
Imports Utility

Public Class ClsTabAppoggio
    Private Shared Log As ILog = LogManager.GetLogger("ClsTabAppoggio")
    Private oQueryManager As New ClsInterDB

    Public Function GetFlussiTracciati(ByVal sTributo As String, ByVal sCodiceISTAT As String) As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE()
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim dvMyDati As DataView
        Dim x As Integer
        Dim oListFlussi() As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE

        Try
            Log.Debug("GestDatiOPENae::GetFlussiTracciati::inizio procedura")
            'oDBModel = cnn.DBConnection()
            'carico i dati da db
            dvMyDati = oQueryManager.GetFlussiCodIstat(oDBModel, sCodiceISTAT, sTributo)
            If dvMyDati.Count > 0 Then
                For x = 0 To dvMyDati.Count - 1
                    'popolo l'oggetto elenco 
                    If ListFlussi(dvMyDati, x, oListFlussi) = False Then
                        Return Nothing
                    End If
                Next
            End If
            Return oListFlussi
            Log.Debug("GestDatiOPENae::GetFlussiTracciati::fine procedura")
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetFlussiTracciati::" & Err.Message)
            Return Nothing
        End Try
    End Function

    Public Function SetTabAppoggio(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String) As Boolean
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim nMyIdFlusso As Integer
        Dim nResult As Integer
        Dim nUtenti As Integer
        Dim nArticoli As Integer

        Try
            Log.Debug("GestDatiOPENae::SetTabAppoggio::inizio procedura")
            'oDBModel = cnn.DBConnection()
            'apro la transazione
            'oDBModel.BeginTransaction()
            '*********************************************************************
            'devo compilare la tabella dei flussi elaborati
            '*********************************************************************
            Log.Debug("GestDatiOPENae::SetTabAppoggio::devo compilare la tabella dei flussi elaborati")
            nMyIdFlusso = oQueryManager.SetIdFlussoEstratto(oDBModel, sCodiceISTAT, sAnnoRif, sTributo)
            If nMyIdFlusso <> -1 Then
                '*********************************************************************
                'inserisco i record che saranno da trasmettere
                '*********************************************************************
                Log.Debug("GestDatiOPENae::SetTabAppoggio::inserisco i record che saranno da trasmettere")
                nResult = oQueryManager.SetDisposizione(oDBModel, sCodiceISTAT, sAnnoRif, nMyIdFlusso)
                If nResult = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
                '*********************************************************************
                'devo compilare i totalizzatori nella tabella dei flussi elaborati
                '*********************************************************************
                Log.Debug("GestDatiOPENae::SetTabAppoggio::devo compilare i totalizzatori nella tabella dei flussi elaborati")
                'numero utenti
                '*********************************************************************
                nUtenti = oQueryManager.GetNUtenti(oDBModel, sCodiceISTAT, sAnnoRif, sTributo)
                If nUtenti = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
                '*********************************************************************
                'numero articoli
                '*********************************************************************
                nArticoli = oQueryManager.GetNArticoli(oDBModel, sCodiceISTAT, sAnnoRif, sTributo)
                If nArticoli = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
                '*********************************************************************
                'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
                '*********************************************************************
                nUtenti = oQueryManager.SetFlussoEstratto(oDBModel, sCodiceISTAT, sTributo, sAnnoRif, nUtenti, nArticoli)
                If nUtenti = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
            Else
                'resetto la transazione
                'oDBModel.RollBackTransaction()
                Return False
            End If
            'confermo la transazione
            'oDBModel.CommitTransaction()
            Log.Debug("GestDatiOPENae::SetTabAppoggio::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetTabAppoggio::" & Err.Message)
            Throw Err
        End Try
    End Function

    Public Function SetTabAppoggio(ByVal oDisposizioni() As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Boolean
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim nMyIdFlusso, x As Integer
        Dim nUtenti As Integer
        Dim nArticoli As Integer

        Try
            Log.Debug("GestDatiOPENae::SetTabAppoggio::inizio procedura")
            'oDBModel = cnn.DBConnection()
            'apro la transazione
            'oDBModel.BeginTransaction()
            '*********************************************************************
            'devo compilare la tabella dei flussi elaborati
            '*********************************************************************
            Log.Debug("GestDatiOPENae::SetTabAppoggio::devo compilare la tabella dei flussi elaborati")
            nMyIdFlusso = oQueryManager.SetIdFlussoEstratto(oDBModel, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno, oDisposizioni(0).sTributo)
            If nMyIdFlusso <> -1 Then
                '*********************************************************************
                'inserisco i record che saranno da trasmettere
                '*********************************************************************
                Log.Debug("GestDatiOPENae::SetTabAppoggio::inserisco i record che saranno da trasmettere")
                For x = 0 To oDisposizioni.GetUpperBound(0)
                    oDisposizioni(x).nIDFlusso = nMyIdFlusso
                    If oQueryManager.SetDisposizione(oDBModel, oDisposizioni(x)) = -1 Then
                        'resetto la transazione
                        'oDBModel.RollBackTransaction()
                        Return False
                    End If
                Next
                '*********************************************************************
                'devo compilare i totalizzatori nella tabella dei flussi elaborati
                '*********************************************************************
                Log.Debug("GestDatiOPENae::SetTabAppoggio::devo compilare i totalizzatori nella tabella dei flussi elaborati")
                'numero utenti
                '*********************************************************************
                nUtenti = oQueryManager.GetNUtenti(oDBModel, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno, oDisposizioni(0).sTributo)
                If nUtenti = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
                '*********************************************************************
                'numero articoli
                '*********************************************************************
                nArticoli = oQueryManager.GetNArticoli(oDBModel, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno, oDisposizioni(0).sTributo)
                If nArticoli = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
                '*********************************************************************
                'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
                '*********************************************************************
                nUtenti = oQueryManager.SetFlussoEstratto(oDBModel, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sTributo, oDisposizioni(0).sAnno, nUtenti, nArticoli)
                If nUtenti = -1 Then
                    'resetto la transazione
                    'oDBModel.RollBackTransaction()
                    Return False
                End If
            Else
                'resetto la transazione
                'oDBModel.RollBackTransaction()
                Return False
            End If
            'confermo la transazione
            'oDBModel.CommitTransaction()
            Log.Debug("GestDatiOPENae::SetTabAppoggio::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetTabAppoggio::" & Err.Message)
            Throw Err
        End Try
    End Function

    Public Function SetTabAppoggio(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByVal sPathNameMyFile As String) As Boolean
        Dim FncGen As New General
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim oMyFile As IO.StreamReader
        Dim sLineFile As String
        Dim dvAnagrafe As DataView
        Dim oRcAnag As ImportAnagICI
        Dim oRcDisp As ImportDisposizioneICI

        Try
            Log.Debug("GestDatiOPENae::SetTabAppoggio::ImportDaFile::inizio procedura")
            'oDBModel = cnn.DBConnection()
            'apro la transazione
            'oDBModel.BeginTransaction()

            sPathNameMyFile = sPathNameMyFile.Replace("HTTP://POMARANCE.RIBESINFORMATICA.IT/TMPOPENAE", "H:\SITI WEB LOCALI\OPENAE-AGENZIA ENTRATE")
            Log.Debug("GestDatiOPENae::SetTabAppoggio::ImportDaFile::file da acquisire " & sPathNameMyFile)
            'apro il file di testo per leggere le righe
            oMyFile = New IO.StreamReader(sPathNameMyFile)
            'leggo il file
            Do
                sLineFile = oMyFile.ReadLine
                If Not IsNothing(sLineFile) Then
                    If sLineFile <> "" Then
                        'controllo su che tipo di record sono posizionato
                        Select Case sLineFile.Substring(0, 2).Trim
                            Case "A"
                                oRcAnag = New ImportAnagICI
                                'Prelevo Cod_Contribuente
                                oRcAnag.nIdContribuente = Int(Mid(sLineFile, 3, 11))
                                'controllo che l'anagrafe non sia già esistente, in tal caso metto il record in un file txt dove l'operatore andrà a vedere a quale contribuente si riferisce
                                dvAnagrafe = oQueryManager.GetAnagrafe(oDBModel, sCodiceISTAT, oRcAnag.nIdContribuente)
                                If Not IsNothing(dvAnagrafe) Then
                                    If dvAnagrafe.Count = 0 Then
                                        'vuol dire che non esiste
                                        oRcAnag.sCodISTAT = sCodiceISTAT
                                        oRcAnag.sCFPIVA = Mid(sLineFile, 44, 16).Trim
                                        oRcAnag.sCognome = Mid(sLineFile, 60, 100).Trim
                                        oRcAnag.sNome = Mid(sLineFile, 160, 50).Trim
                                        oRcAnag.sSesso = Mid(sLineFile, 210, 1).Trim
                                        oRcAnag.sDataNascita = FncGen.ReplaceDataForDB(Mid(sLineFile, 211, 10).Trim)
                                        oRcAnag.sComuneNascita = Mid(sLineFile, 231, 50).Trim
                                        oRcAnag.sPVNascita = Mid(sLineFile, 281, 2).Trim
                                        oRcAnag.sNazionalita = Mid(sLineFile, 283, 35).Trim
                                        oRcAnag.sViaRes = Mid(sLineFile, 318, 50).Trim
                                        oRcAnag.sFrazioneRes = Mid(sLineFile, 368, 50).Trim
                                        oRcAnag.sCivicoRes = Mid(sLineFile, 418, 10).Trim
                                        oRcAnag.sCAPRes = Mid(sLineFile, 428, 10).Trim
                                        oRcAnag.sCittaRes = Mid(sLineFile, 438, 50).Trim
                                        oRcAnag.sPVRes = Mid(sLineFile, 488, 2).Trim
                                        oRcAnag.sNominativoInvio = Mid(sLineFile, 490, 50).Trim
                                        oRcAnag.sViaInvio = Mid(sLineFile, 540, 50).Trim
                                        oRcAnag.sCivicoInvio = Mid(sLineFile, 590, 10).Trim
                                        oRcAnag.sCAPInvio = Mid(sLineFile, 600, 10).Trim
                                        oRcAnag.sCittaInvio = Mid(sLineFile, 610, 50).Trim
                                        oRcAnag.sPVInvio = Mid(sLineFile, 660, 2).Trim
                                        'inserimento in tabella
                                        If oQueryManager.SetAnagrafe(oDBModel, oRcAnag) = -1 Then
                                            'resetto la transazione
                                            'oDBModel.RollBackTransaction()
                                            Return False
                                        End If
                                    End If
                                End If
                            Case "V"
                                oRcDisp = New ImportDisposizioneICI
                                oRcDisp.sCodISTAT = sCodiceISTAT
                                oRcDisp.nIdContribuente = Mid(sLineFile, 7, 11).Trim
                                oRcDisp.sAnno = Mid(sLineFile, 3, 4).Trim
                                oRcDisp.sAnnoRif = sAnnoRif
                                oRcDisp.sDataPagamento = FncGen.ReplaceDataForDB(Mid(sLineFile, 18, 10).Trim)
                                If Mid(sLineFile, 28, 1).Trim <> "" Then
                                    oRcDisp.sFlagAS = Mid(sLineFile, 28, 1).Trim
                                End If
                                If Mid(sLineFile, 59, 15).Trim <> "" Then
                                    oRcDisp.nImpAbiPrinc = Mid(sLineFile, 59, 15).Replace(",", ".")
                                End If
                                If Mid(sLineFile, 74, 15).Trim <> "" Then
                                    oRcDisp.nImpAltriFab = Mid(sLineFile, 74, 15).Replace(",", ".")
                                End If
                                If Mid(sLineFile, 44, 15).Trim <> "" Then
                                    oRcDisp.nImpAreeFab = Mid(sLineFile, 44, 15).Replace(",", ".")
                                End If
                                If Mid(sLineFile, 29, 15).Trim <> "" Then
                                    oRcDisp.nImpTerAgr = Mid(sLineFile, 29, 15).Replace(",", ".")
                                End If
                                If Mid(sLineFile, 89, 15).Trim <> "" Then
                                    oRcDisp.nImpDetrazione = Mid(sLineFile, 89, 15).Replace(",", ".")
                                End If
                                If Mid(sLineFile, 104, 15).Trim <> "" Then
                                    oRcDisp.nImpVersamento = Mid(sLineFile, 104, 15).Replace(",", ".")
                                End If
                                If Mid(sLineFile, 119, 4).Trim <> "" Then
                                    oRcDisp.nNumFab = Mid(sLineFile, 119, 4)
                                End If
                                If Mid(sLineFile, 123, 11).Trim <> "" Then
                                    oRcDisp.nNumQuietanza = Mid(sLineFile, 123, 11).Trim
                                End If
                                oRcDisp.sNumMovimento = Mid(sLineFile, 134, 10).Trim
                                oRcDisp.sDataAccredito = FncGen.ReplaceDataForDB(Mid(sLineFile, 144, 10).Trim)
                                oRcDisp.sFlagRavvOperoso = Mid(sLineFile, 154, 1).Trim
                                If oRcDisp.sFlagRavvOperoso = "2" Then
                                    oRcDisp.sFlagRavvOperoso = ""
                                    oRcDisp.sTipoBollettinoViolazioni = "1"
                                End If
                                oRcDisp.sNumSanzione = Mid(sLineFile, 155, 9).Trim
                                oRcDisp.sDataSanzione = FncGen.ReplaceDataForDB(Mid(sLineFile, 164, 10).Trim)
                                'Log.Debug("GestDatiOPENae::SetTabAppoggio(3parametri)::leggo data sanzione::" & Mid(sLineFile, 164, 10).Trim & "::" & oRcDisp.sDataSanzione)
                                'inserimento in tabella
                                If oQueryManager.SetDisposizioneICI(oDBModel, oRcDisp) = -1 Then
                                    'resetto la transazione
                                    'oDBModel.RollBackTransaction()
                                    Return False
                                End If
                        End Select
                    End If
                End If
            Loop Until sLineFile Is Nothing
            oMyFile.Close()
            Rename(sPathNameMyFile, sPathNameMyFile + "." + Format(Now, "yyyyMMdd_hhmmss"))
            'confermo la transazione
            'oDBModel.CommitTransaction()
            Log.Debug("GestDatiOPENae::SetTabAppoggio::ImportDaFile::fine procedura")
            Return True
        Catch Err As Exception
            'oDBModel.RollBackTransaction()
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetTabAppoggio::ImportDaFile::" & Err.Message)
            Throw Err
        End Try
    End Function

    Public Function SetTabAppoggio(ByVal sCodiceISTAT As String, ByVal sPathNameMyFile As String) As Boolean
        Dim FncGen As New General
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim oMyFile As IO.StreamReader
        Dim sLineFile As String
        Dim dvAnagrafe As DataView
        Dim oRcTesta As New ImportTestaICI
        Dim oRcAnag As ImportAnagICI
        Dim oRcDisp As ImportDisposizioneICI
        Dim oRcCoda As New ImportCodaICI
        Dim bFirstIns As Boolean = True

        Try
            Log.Debug("GestDatiOPENae::SetTabAppoggio::ImportDaFile::inizio procedura")
            'oDBModel = cnn.DBConnection()
            'apro la transazione
            'oDBModel.BeginTransaction()

            'apro il file di testo per leggere le righe
            oMyFile = New IO.StreamReader(sPathNameMyFile)
            'leggo il file
            Do
                sLineFile = oMyFile.ReadLine
                If Not IsNothing(sLineFile) Then
                    If sLineFile <> "" Then
                        'controllo su che tipo di record sono posizionato
                        Select Case sLineFile.Substring(0, 2)
                            Case "RT"  'sono in presenza del record di testa
                                oRcTesta.sCodISTAT = sCodiceISTAT
                                'Prelevo il codice flusso 
                                oRcTesta.nCodFlussoEnte = Int(Mid(sLineFile, 3, 7))
                                'Prelevo la data di creazione del file con formato MM/GG/AAAA
                                oRcTesta.sDataCreazione = FncGen.ReplaceDataForDB(Mid(sLineFile, 10, 8))
                                'Prelevo il nome del file che sto acquisendo
                                oRcTesta.sNomeTracciato = Mid(sLineFile, 18, 30).Trim 'Nome_File
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Letto RECORD TESTA")
                            Case "TF"  'sono in presenza del record di tab_flussi_pagamenti_ICI
                                'Prelevo Cod_Flusso_Pagamenti
                                oRcTesta.nCodFlussoPagamenti = Mid(sLineFile, 3, 7) 'COD_FLUSSO
                                'prelevo nome_file
                                oRcTesta.sNomeFlusso = Mid(sLineFile, 10, 50)  'NOME_FILE
                                'Prelevo la Data_Creazione con formato MM/GG/AAAA
                                oRcTesta.sDataCreazioneFlusso = FncGen.ReplaceDataForDB(Mid(sLineFile, 60, 8))
                                'Prelevo la divisa del flusso
                                oRcTesta.sDivisa = Mid(sLineFile, 68, 1) 'DIVISA FLUSSO
                                'prelevo Totale_Pagamenti_Acquisiti
                                oRcTesta.nTotPagamentiAcq = Int(Mid(sLineFile, 69, 15)) 'NUMERO_PAGAMENTI
                                'prelevo la Data_Inizio in Formato MM/GG/AAAA
                                oRcTesta.sDataInizio = FncGen.ReplaceDataForDB(Mid(sLineFile, 84, 8))
                                'prelevo la Data_Fine in Formato MM/GG/AAAA 
                                oRcTesta.sDataFine = FncGen.ReplaceDataForDB(Mid(sLineFile, 92, 8))
                                'prelevo l'ANNO
                                oRcTesta.sAnnoFlusso = Mid(sLineFile, 100, 4) 'Anno 
                                'totale importi positivi euro
                                oRcTesta.nTotImpPos = Mid(sLineFile, 104, 15) / 100
                                'totale_importi_negativi_euro
                                oRcTesta.nTotImpNeg = Mid(sLineFile, 119, 15) / 100
                                'totale_riversato
                                oRcTesta.nTotaleRiversato = Mid(sLineFile, 134, 15) / 100
                                'totale_importi_sanzioni_euro
                                oRcTesta.nTotaleImpSanzioni = Mid(sLineFile, 149, 15) / 100
                                'numero_sanzioni
                                oRcTesta.nTotSanzioni = Int(Mid(sLineFile, 164, 7))
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Letto RECORD di TAB_FLUSSI_PAGAMENTI_ICI")

                                'inserimento in tabella
                                If bFirstIns = True Then
                                    If oQueryManager.SetFlussiAcq(oDBModel, sCodiceISTAT, oRcTesta) = -1 Then
                                        'resetto la transazione
                                        'oDBModel.RollBackTransaction()
                                        Return False
                                    End If
                                    bFirstIns = False
                                End If
                            Case "AN" 'sono in presenza dei Record di Anagrafe
                                oRcAnag = New ImportAnagICI
                                'Prelevo Cod_Contribuente
                                oRcAnag.nIdContribuente = Int(Mid(sLineFile, 3, 7))
                                'controllo che l'anagrafe non sia già esistente, in tal caso metto il record in un file txt dove l'operatore andrà a vedere a quale contribuente si riferisce
                                dvAnagrafe = oQueryManager.GetAnagrafe(oDBModel, sCodiceISTAT, oRcAnag.nIdContribuente)
                                If Not IsNothing(dvAnagrafe) Then
                                    If dvAnagrafe.Count = 0 Then
                                        'vuol dire che non esiste
                                        oRcAnag.sCodISTAT = sCodiceISTAT
                                        'Prelevo Cod_Fiscale/Partita_Iva
                                        oRcAnag.sCFPIVA = Mid(sLineFile, 10, 16)
                                        'Prelevo Cognome
                                        oRcAnag.sCognome = Replace(Mid(sLineFile, 26, 50), "'", "''").Trim
                                        'Prelevo Nome
                                        oRcAnag.sNome = Replace(Mid(sLineFile, 76, 30), "'", "''").Trim
                                        'Prelevo Sesso
                                        oRcAnag.sSesso = Mid(sLineFile, 106, 1).Trim
                                        'Prelevo Data_Nascita in formato MM/GG/AAAA 
                                        oRcAnag.sDataNascita = FncGen.ReplaceDataForDB(Mid(sLineFile, 107, 8))
                                        If StrComp(oRcAnag.sDataNascita, "/  /") = 0 Then
                                            oRcAnag.sDataNascita = "00000000"
                                        End If
                                        'Prelevo Comune_nascita
                                        oRcAnag.sComuneNascita = Replace(Mid(sLineFile, 115, 30), "'", "''").Trim
                                        'Prelevo Prov_Nascita
                                        oRcAnag.sPVNascita = Mid(sLineFile, 145, 2).Trim
                                        'Prelevo Nazionalità
                                        oRcAnag.sNazionalita = Mid(sLineFile, 147, 35).Trim
                                        'Prelevo Via_Res
                                        oRcAnag.sViaRes = Replace(Mid(sLineFile, 182, 30), "'", "''").Trim
                                        'Prelevo Frazione_Res
                                        oRcAnag.sFrazioneRes = Replace(Mid(sLineFile, 212, 30), "'", "''").Trim
                                        'Prelevo Civico_Res
                                        oRcAnag.sCivicoRes = Mid(sLineFile, 242, 10).Trim
                                        'Prelevo Cap_Res
                                        oRcAnag.sCAPRes = Mid(sLineFile, 252, 5).Trim
                                        'Prelevo Citta_Res
                                        oRcAnag.sCittaRes = Replace(Mid(sLineFile, 257, 30), "'", "''").Trim
                                        'Prelevo Provincia_Res
                                        oRcAnag.sPVRes = Mid(sLineFile, 287, 2).Trim
                                        'Prelevo c/o 
                                        oRcAnag.sNominativoInvio = Replace(Mid(sLineFile, 289, 30), "'", "''").Trim
                                        'Prelevo Via_c/o
                                        oRcAnag.sViaInvio = Replace(Mid(sLineFile, 319, 30), "'", "''").Trim
                                        'Prelevo Civico_C_O
                                        oRcAnag.sCivicoInvio = Mid(sLineFile, 349, 10).Trim
                                        'Prelevo Cap_C_O_Anagrafe
                                        oRcAnag.sCAPInvio = Mid(sLineFile, 359, 5).Trim
                                        'Prelevo Citta_C_O_Anagrafe 
                                        oRcAnag.sCittaInvio = Replace(Mid(sLineFile, 364, 30), "'", "''").Trim
                                        'Prelevo Provincia_C_O_Anagrafe
                                        oRcAnag.sPVInvio = Mid(sLineFile, 394, 2).Trim
                                        Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Letto RECORD di ANAGRAFE")
                                        'inserimento in tabella
                                        If oQueryManager.SetAnagrafe(oDBModel, oRcAnag) = -1 Then
                                            'resetto la transazione
                                            'oDBModel.RollBackTransaction()
                                            Return False
                                        End If
                                    End If
                                End If
                            Case "PA" 'sono in presenza del record che descrive il pagamento
                                oRcDisp = New ImportDisposizioneICI
                                oRcDisp.sCodISTAT = sCodiceISTAT
                                oRcDisp.sAnnoRif = oRcTesta.sAnnoFlusso
                                'prelevo id_versamento
                                oRcDisp.nIdVersamento = Int(Mid(sLineFile, 3, 7))
                                'prelevo Cod_Flusso 
                                oRcDisp.nCodFlusso = Int(Mid(sLineFile, 10, 7))
                                'Prelevo CFPIVA
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere CFPIVA:" & Mid(sLineFile, 17, 16).Trim)
                                oRcDisp.sCFPIVA = Mid(sLineFile, 17, 16).Trim
                                'prelevo cognome_nome
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere NOMINATIVO:" & Mid(sLineFile, 33, 45).Trim)
                                oRcDisp.sNominativo = Replace(Mid(sLineFile, 33, 45), "'", "''").Trim
                                'Prelevo Data_accredito
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere DATA ACCREDITO:" & Mid(sLineFile, 78, 8).Trim)
                                oRcDisp.sDataAccredito = FncGen.ReplaceDataForDB(Mid(sLineFile, 78, 8))
                                'Prelevo Data_PAgamento
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere DATA PAGAMENTO:" & Mid(sLineFile, 86, 8).Trim)
                                oRcDisp.sDataPagamento = FncGen.ReplaceDataForDB(Mid(sLineFile, 86, 8))
                                'Prelevo Flag_acconto_Saldo
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere ACCONTO/SALDO:" & Mid(sLineFile, 94, 1).Trim)
                                oRcDisp.sFlagAS = Mid(sLineFile, 94, 1).Trim
                                'Prelevo Anno 
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere ANNO:" & Mid(sLineFile, 95, 4).Trim)
                                oRcDisp.sAnno = Mid(sLineFile, 95, 4)
                                'Prelevo N_Fab
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere N.FAB:" & Mid(sLineFile, 99, 7).Trim)
                                oRcDisp.nNumFab = Mid(sLineFile, 99, 7)
                                'Prelevo Importo
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere IMPORTO:" & Mid(sLineFile, 106, 15).Trim)
                                oRcDisp.nImpVersamento = Mid(sLineFile, 106, 15) / 100
                                'Prelevo Importo_Ter_Agr
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere TER AGR:" & Mid(sLineFile, 121, 15).Trim)
                                oRcDisp.nImpTerAgr = Mid(sLineFile, 121, 15) / 100
                                'Prelevo Importo_Aree_Fab
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere AREE FAB:" & Mid(sLineFile, 136, 15).Trim)
                                oRcDisp.nImpAreeFab = Mid(sLineFile, 136, 15) / 100
                                'Prelevo Importo_Altri_Fab
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere ALTRI FAB:" & Mid(sLineFile, 151, 15).Trim)
                                oRcDisp.nImpAltriFab = Mid(sLineFile, 151, 15) / 100
                                'Prelevo Importo_abi_prin
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere ABI PRIN:" & Mid(sLineFile, 166, 15).Trim)
                                oRcDisp.nImpAbiPrinc = Mid(sLineFile, 166, 15) / 100
                                'Prelevo Detrazione
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere DETRAZIONE:" & Mid(sLineFile, 181, 15).Trim)
                                oRcDisp.nImpDetrazione = Mid(sLineFile, 181, 15) / 100
                                'Prelevo Indirizzo_Res
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere INDIRIZZO:" & Mid(sLineFile, 196, 40).Trim)
                                oRcDisp.sIndirizzoRes = Replace(Mid(sLineFile, 196, 40), "'", "''").Trim
                                'prelevo Cap_Res
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere CAP:" & Mid(sLineFile, 236, 5).Trim)
                                oRcDisp.sCapRes = Mid(sLineFile, 236, 5).Trim
                                'Prelevo Citta_Res
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere CITTA:" & Mid(sLineFile, 241, 30).Trim)
                                oRcDisp.sCittaRes = Replace(Mid(sLineFile, 241, 30), "'", "''").Trim
                                'Prelevo Bollettino Ex Rurale
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere EX RURALE:" & Mid(sLineFile, 271, 1).Trim)
                                oRcDisp.sBollettinoEXRurale = Mid(sLineFile, 271, 1).Trim
                                'Prelevo Data_Sanzione
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere DATA SANZIONE:" & Mid(sLineFile, 272, 8).Trim)
                                oRcDisp.sDataSanzione = FncGen.ReplaceDataForDB(Mid(sLineFile, 272, 8))
                                'Prelevo N_Sanzione
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere N.SANZIONE:" & Mid(sLineFile, 280, 9).Trim)
                                oRcDisp.sNumSanzione = Mid(sLineFile, 280, 9).Trim
                                'Prelevo N_Movimento 
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere N.MOVIMENTO:" & Mid(sLineFile, 289, 7).Trim)
                                oRcDisp.sNumMovimento = Mid(sLineFile, 289, 7).Trim
                                'Prelevo Spazio_Libero
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere FILLER:" & Mid(sLineFile, 296, 13).Trim)
                                oRcDisp.sSpazioLibero = Mid(sLineFile, 296, 13).Trim
                                'Prelevo Data_Flusso_Rendicontazione in formato MM/GG/AAAA
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere DATA RENDICONTAZIONE:" & Mid(sLineFile, 309, 8).Trim)
                                oRcDisp.sDataFlussoRend = FncGen.ReplaceDataForDB(Mid(sLineFile, 309, 8))
                                'Prelevo Cod_Contribuente
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere COD CONTRIBUENTE:" & Mid(sLineFile, 317, 7).Trim)
                                oRcDisp.nIdContribuente = Mid(sLineFile, 317, 7)
                                'Prelevo Cod_Contribuente_Simile
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere COD CONTRIBUENTE SIMILE:" & Mid(sLineFile, 324, 7).Trim)
                                oRcDisp.nIdContribuenteSimile = Mid(sLineFile, 324, 7)
                                'Prelevo Flag_Trattato
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere FLAG TRATTATO:" & Mid(sLineFile, 331, 1).Trim)
                                oRcDisp.nFlagTrattato = Mid(sLineFile, 331, 1)
                                If oRcDisp.nFlagTrattato = "1" Then
                                    oRcDisp.nFlagTrattato = -1
                                Else
                                    oRcDisp.nFlagTrattato = 0
                                End If
                                'Prelevo Cod_Divisa 
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere DIVISA:" & Mid(sLineFile, 332, 1).Trim)
                                oRcDisp.sDivisa = Mid(sLineFile, 332, 1).Trim
                                'Prelevo Ravvedimento_Operoso
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere RAVVEDIMENTO OPEROSO:" & Mid(sLineFile, 333, 1).Trim)
                                oRcDisp.sFlagRavvOperoso = Mid(sLineFile, 333, 1).Trim
                                'Prelevo Cod_Flusso_Ap
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere COD FLUSSO AP:" & Mid(sLineFile, 334, 7).Trim)
                                oRcDisp.nCodFlussoAP = Mid(sLineFile, 334, 7)
                                'Prelevo Progressivo_Pagamento_Ap
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere PROG PAGAMENTO AP:" & Mid(sLineFile, 341, 7).Trim)
                                oRcDisp.nProgrPagAP = Mid(sLineFile, 341, 7)
                                'Prelevo Cod_Tipo_Pagamento
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere TIPO PAGAMENTO:" & Mid(sLineFile, 348, 2).Trim)
                                oRcDisp.sCodTipoPagamento = Mid(sLineFile, 348, 2).Trim
                                'Prelevo Cod Comunico
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere COD COMUNICO:" & Mid(sLineFile, 350, 14).Trim)
                                oRcDisp.sCodiceComunico = Mid(sLineFile, 350, 14).Trim
                                'Importo pagato euro
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere IMPORTO PAGATO:" & Mid(sLineFile, 364, 15).Trim)
                                oRcDisp.nImpVersato = Mid(sLineFile, 364, 15) / 100
                                'Nome Immagine Bollettino GIF
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere IMMAGINE:" & Mid(sLineFile, 379, 18).Trim)
                                oRcDisp.sNomeImmagine = Mid(sLineFile, 379, 18).Trim
                                'Provenienza Pagamento
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere PROVENIENZA:" & Mid(sLineFile, 397, 10).Trim)
                                oRcDisp.sProvenienza = Mid(sLineFile, 397, 10).Trim
                                'visualizzazione immagine Bollettino 
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Devo leggere VISUALIZZA IMMAGINE:" & Mid(sLineFile, 407, 1).Trim)
                                oRcDisp.sViewImmagine = Mid(sLineFile, 407, 1).Trim
                                '***28/06/2007***
                                'n progressivo rendicontazione
                                oRcDisp.nProgRendicontaz = Mid(sLineFile, 408, 5)
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Letto RECORD di PAGAMENTO")
                                'inserimento in tabella
                                If oQueryManager.SetDisposizioneICI(oDBModel, oRcDisp) = -1 Then
                                    'resetto la transazione
                                    'oDBModel.RollBackTransaction()
                                    Return False
                                End If
                            Case "RC"  'sono in presenza del record di coda
                                oRcCoda.nCodFlussoPagamenti = oRcTesta.nCodFlussoPagamenti
                                'Prelevo il numero delle posizioni nella tabella Pagamenti ICI
                                oRcCoda.nTotDisposizioni = Int(Mid(sLineFile, 10, 7))
                                'Prelevo il totale delle Posizioni Anagrafiche
                                oRcCoda.nTotAnagrafiche = Int(Mid(sLineFile, 17, 7))
                                'Prelevo il numero totale dei Versamenti
                                oRcCoda.nTotVersamenti = Int(Mid(sLineFile, 24, 7))
                                'Prelevo Totale_Importi_Versamenti
                                oRcCoda.nTotImpVersamenti = Mid(sLineFile, 46, 15) / 100
                                'Prelevo Totale_Importi_Terreni_Agr
                                oRcCoda.nTotImpTerAgr = Mid(sLineFile, 76, 15) / 100
                                'Prelevo Totale_Importi_Terreni_Fab
                                oRcCoda.nTotImpAreeFab = Mid(sLineFile, 106, 15) / 100
                                'Prelevo Totale_Importi_Abitazione_Principale
                                oRcCoda.nTotImpAbiPrin = Mid(sLineFile, 136, 15) / 100
                                'Prelevo Totale_Importi_Altri_Fab
                                oRcCoda.nTotImpAltriFab = Mid(sLineFile, 166, 15) / 100
                                'Prelevo Totale_Detrazione
                                oRcCoda.nTotImpDetrazione = Mid(sLineFile, 196, 15) / 100
                                'prelevo il numero delle violazioni
                                oRcCoda.nTotViolazioni = Mid(sLineFile, 211, 7)
                                'prelevo l'importo totale dell violazioni 
                                oRcCoda.nTotImpViolazioni = Mid(sLineFile, 218, 15) / 100
                                Log.Info("GestDatiOPENae::SetTabAppoggio::ImportDaFile::Letto RECORD di CODA")
                                'inserimento in tabella
                                If oQueryManager.SetFlussiAcq(oDBModel, oRcCoda) = -1 Then
                                    'resetto la transazione
                                    'oDBModel.RollBackTransaction()
                                    Return False
                                End If
                            Case "AV", "VE", "VI" 'NON GESTITI: Record di Anagrafe_Versamenti_ICI, di Versamenti ici, Violazioni ici
                        End Select
                    End If
                End If
            Loop Until sLineFile Is Nothing
            oMyFile.Close()
            Rename(sPathNameMyFile, sPathNameMyFile + "." + Format(Now, "yyyyMMdd_hhmmss"))
            'confermo la transazione
            'oDBModel.CommitTransaction()
            Log.Debug("GestDatiOPENae::SetTabAppoggio::ImportDaFile::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetTabAppoggio::ImportDaFile::" & Err.Message)
            oMyFile.Close()
            Throw Err
        End Try
    End Function

    Public Function CheckFile(ByVal sPathNameMyFile As String) As Boolean
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))

        Try
            Log.Debug("GestDatiOPENae::CheckFile::inizio procedura")
            'oDBModel = cnn.DBConnection()

            'controllo se il file esiste

            'controllo che file non sia già presente

            'controllo la presenza del record di testa

            'controllo la presenza del record di tab_flussi_pagamenti_ICI

            'controllo la presenza dei Record di Anagrafe

            'controllo la presenza del record di coda

            'controllo la presenza del record che descrive il pagamento

            Log.Debug("GestDatiOPENae::CheckFile::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::CheckFile::" & Err.Message)
            Throw Err
        End Try
    End Function

    Public Function DeleteTabAppoggio(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String) As Boolean
        Dim cnn As New DBConnection
        Dim oDBModel As New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(New System.Configuration.AppSettingsReader().GetValue("ConnectionStringDB", GetType(String)), String))
        Dim oConst As New General

        Try
            Log.Debug("GestDatiOPENae::DeleteTabAppoggio::inizio procedura")
            'oDBModel = cnn.DBConnection()

            '*******************************************
            'svuoto la tabella dei dati d'appoggio
            '*******************************************
            Select Case sTributo
                Case oConst.TRIBUTO_ICI
                    If oQueryManager.DeleteDisposizione(oDBModel, sCodiceISTAT, sAnnoRif) = -1 Then
                        Return False
                    End If
                Case Else
                    If oQueryManager.DeleteDisposizione(oDBModel, sTributo, sCodiceISTAT, sAnnoRif) = -1 Then
                        Return False
                    End If
            End Select

            '*******************************************
            'svuoto la tabella dei flussi
            '*******************************************
            If oQueryManager.DeleteFlusso(oDBModel, sTributo, sCodiceISTAT, sAnnoRif) = -1 Then
                Return False
            End If
            Log.Debug("GestDatiOPENae::DeleteTabAppoggio::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::DeleteTabAppoggio::" & Err.Message)
            Return False
        End Try
    End Function

    Private Function ListFlussi(ByVal dvFlussi As DataView, ByVal nList As Integer, ByRef oListFlussi() As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE) As Boolean
        Try
            Dim oFlusso As New AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE
            Dim FncGen As New General

            oFlusso.Anno = CStr(dvFlussi.Item(nList)("ANNO"))
            oFlusso.CodiceISTAT = CStr(dvFlussi.Item(nList)("CODICE_ISTAT"))
            If Not IsDBNull(dvFlussi.Item(nList)("data_estrazione")) Then
                oFlusso.DataEstrazione = FncGen.ReplaceDataForTXT(dvFlussi.Item(nList)("DATA_ESTRAZIONE"), "/")
            End If
            oFlusso.IdFlusso = CInt(dvFlussi.Item(nList)("id"))
            If Not IsDBNull(dvFlussi.Item(nList)("nome_file")) Then
                oFlusso.NomeFile = CStr(dvFlussi.Item(nList)("NOME_FILE"))
            End If
            If Not IsDBNull(dvFlussi.Item(nList)("narticoli")) Then
                oFlusso.NumeroArticoli = CInt(dvFlussi.Item(nList)("NARTICOLI"))
            End If
            If Not IsDBNull(dvFlussi.Item(nList)("nrecord")) Then
                oFlusso.NumeroRecords = (dvFlussi.Item(nList)("NRECORD"))
            End If
            If Not IsDBNull(dvFlussi.Item(nList)("nutenti")) Then
                oFlusso.NumeroUtenti = (dvFlussi.Item(nList)("NUTENTI"))
            End If
            ReDim Preserve oListFlussi(nList)
            oListFlussi(nList) = oFlusso
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetFlussiTracciati::ListFlussi::" & Err.Message)
            Return False
        End Try
    End Function
End Class
