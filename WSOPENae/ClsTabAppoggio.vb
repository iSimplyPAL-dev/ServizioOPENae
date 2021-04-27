Imports log4net
'Imports Utility

Public Class ClsTabAppoggio
    'Inherits ClsInterDB
    Private Shared Log As ILog = LogManager.GetLogger("ClsTabAppoggio")
    Private oQueryManager As ClsInterDB

    Public Function GetFlussiTracciati(ByVal sCodiceISTAT As String) As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE()
        'Dim cnn As New DBConnection
        'Dim oDBManager As DBManager
        'Dim dvMyDati As DataView
        'Dim x As Integer
        'Dim oListFlussi() As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE

        'Try
        '    Log.Debug("ServiceOPENae::GetFlussiTracciati::inizio procedura")
        '    oDBManager = cnn.DBConnection()
        '    'carico i dati da db
        '    dvMyDati = oQueryManager.GetFlussiCodIstat(oDBManager, sCodiceISTAT)
        '    If dvMyDati.Count > 0 Then
        '        For x = 0 To dvMyDati.Count
        '            'popolo l'oggetto elenco 
        '            If ListFlussi(dvMyDati, x, oListFlussi) = False Then
        '                Return Nothing
        '            End If
        '        Next
        '    End If
        '    Return oListFlussi
        '    Log.Debug("ServiceOPENae::GetFlussiTracciati::fine procedura")
        'Catch Err As Exception
        '    Log.Debug("Si è verificato un errore in ServiceOPENae::GetFlussiTracciati::" & Err.Message)
        '    Log.Warn("Si è verificato un errore in ServiceOPENae::GetFlussiTracciati::" & Err.Message)
        '    Return Nothing
        'End Try
    End Function

    Public Function SetTabAppoggio(ByVal sAnnoRif As String, ByVal sCodiceISTAT As String) As Boolean
        'Dim cnn As New DBConnection
        'Dim oDBManager As DBManager
        'Dim nMyIdFlusso As Integer
        'Dim nResult As Integer
        'Dim nUtenti As Integer
        'Dim nArticoli As Integer

        'Try
        '    Log.Debug("ServiceOPENae::SetTabAppoggio::inizio procedura")
        '    oDBManager = cnn.DBConnection()
        '    'apro la transazione
        '    oDBManager.BeginTransaction()
        '    '*********************************************************************
        '    'devo compilare la tabella dei flussi elaborati
        '    '*********************************************************************
        '    Log.Debug("ServiceOPENae::SetTabAppoggio::devo compilare la tabella dei flussi elaborati")
        '    nMyIdFlusso = oQueryManager.GetIdFlusso(oDBManager, sCodiceISTAT, sAnnoRif)
        '    If nMyIdFlusso <> -1 Then
        '        '*********************************************************************
        '        'inserisco i record che saranno da trasmettere
        '        '*********************************************************************
        '        Log.Debug("ServiceOPENae::SetTabAppoggio::inserisco i record che saranno da trasmettere")
        '        nResult = oQueryManager.SetDisposizione(oDBManager, sCodiceISTAT, sAnnoRif, nMyIdFlusso)
        '        If nResult = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '        '*********************************************************************
        '        'devo compilare i totalizzatori nella tabella dei flussi elaborati
        '        '*********************************************************************
        '        Log.Debug("ServiceOPENae::SetTabAppoggio::devo compilare i totalizzatori nella tabella dei flussi elaborati")
        '        'numero utenti
        '        '*********************************************************************
        '        nUtenti = oQueryManager.GetNUtenti(oDBManager, sCodiceISTAT, sAnnoRif)
        '        If nUtenti = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '        '*********************************************************************
        '        'numero articoli
        '        '*********************************************************************
        '        nArticoli = oQueryManager.GetNArticoli(oDBManager, sCodiceISTAT, sAnnoRif)
        '        If nArticoli = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '        '*********************************************************************
        '        'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
        '        '*********************************************************************
        '        nUtenti = oQueryManager.SetFlusso(oDBManager, sCodiceISTAT, sAnnoRif, nUtenti, nArticoli)
        '        If nUtenti = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '    Else
        '        'resetto la transazione
        '        oDBManager.RollBackTransaction()
        '        Return False
        '    End If
        '    'confermo la transazione
        '    oDBManager.CommitTransaction()
        '    Log.Debug("ServiceOPENae::SetTabAppoggio::fine procedura")
        'Catch Err As Exception
        '    Log.Debug("Si è verificato un errore in ServiceOPENae::SetTabAppoggio::" & Err.Message)
        '    Log.Warn("Si è verificato un errore in ServiceOPENae::SetTabAppoggio::" & Err.Message)
        '    Throw Err
        'End Try
    End Function

    Public Function SetTabAppoggio(ByVal oDisposizioni() As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Boolean
        'Dim cnn As New DBConnection
        'Dim oDBManager As DBManager
        'Dim nMyIdFlusso, x As Integer
        'Dim nResult As Integer
        'Dim nUtenti As Integer
        'Dim nArticoli As Integer

        'Try
        '    Log.Debug("ServiceOPENae::SetTabAppoggio::inizio procedura")
        '    oDBManager = cnn.DBConnection()
        '    'apro la transazione
        '    oDBManager.BeginTransaction()
        '    '*********************************************************************
        '    'devo compilare la tabella dei flussi elaborati
        '    '*********************************************************************
        '    Log.Debug("ServiceOPENae::SetTabAppoggio::devo compilare la tabella dei flussi elaborati")
        '    nMyIdFlusso = oQueryManager.GetIdFlusso(oDBManager, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno)
        '    If nMyIdFlusso <> -1 Then
        '        '*********************************************************************
        '        'inserisco i record che saranno da trasmettere
        '        '*********************************************************************
        '        Log.Debug("ServiceOPENae::SetTabAppoggio::inserisco i record che saranno da trasmettere")
        '        For x = 0 To oDisposizioni.GetUpperBound(0)
        '            oDisposizioni(x).nIDFlusso = nMyIdFlusso
        '            If oQueryManager.SetDisposizione(oDBManager, oDisposizioni(x)) = -1 Then
        '                'resetto la transazione
        '                oDBManager.RollBackTransaction()
        '                Return False
        '            End If
        '        Next
        '        '*********************************************************************
        '        'devo compilare i totalizzatori nella tabella dei flussi elaborati
        '        '*********************************************************************
        '        Log.Debug("ServiceOPENae::SetTabAppoggio::devo compilare i totalizzatori nella tabella dei flussi elaborati")
        '        'numero utenti
        '        '*********************************************************************
        '        nUtenti = oQueryManager.GetNUtenti(oDBManager, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno)
        '        If nUtenti = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '        '*********************************************************************
        '        'numero articoli
        '        '*********************************************************************
        '        nArticoli = oQueryManager.GetNArticoli(oDBManager, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno)
        '        If nArticoli = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '        '*********************************************************************
        '        'devo aggiornare la tabella AE_FLUSSI_ESTRATTI
        '        '*********************************************************************
        '        nUtenti = oQueryManager.SetFlusso(oDBManager, oDisposizioni(0).sCodISTAT, oDisposizioni(0).sAnno, nUtenti, nArticoli)
        '        If nUtenti = -1 Then
        '            'resetto la transazione
        '            oDBManager.RollBackTransaction()
        '            Return False
        '        End If
        '    Else
        '        'resetto la transazione
        '        oDBManager.RollBackTransaction()
        '        Return False
        '    End If
        '    'confermo la transazione
        '    oDBManager.CommitTransaction()
        '    Log.Debug("ServiceOPENae::SetTabAppoggio::fine procedura")
        'Catch Err As Exception
        '    Log.Debug("Si è verificato un errore in ServiceOPENae::SetTabAppoggio::" & Err.Message)
        '    Log.Warn("Si è verificato un errore in ServiceOPENae::SetTabAppoggio::" & Err.Message)
        '    Throw Err
        'End Try
    End Function

    Public Function DeleteTabAppoggio(ByVal sAnnoRif As String, ByVal sCodiceISTAT As String) As Boolean
        'Dim cnn As New DBConnection
        'Dim oDBManager As DBManager
        'Dim nResult As Integer

        'Try
        '    Log.Debug("ServiceOPENae::DeleteTabAppoggio::inizio procedura")
        '    oDBManager = cnn.DBConnection()

        '    '*******************************************
        '    'svuoto la tabella dei dati d'appoggio
        '    '*******************************************
        '    If oQueryManager.DeleteDisposizione(oDBManager, sCodiceISTAT, sAnnoRif) = -1 Then
        '        Return False
        '    End If

        '    '*******************************************
        '    'svuoto la tabella dei flussi
        '    '*******************************************
        '    If oQueryManager.DeleteFlusso(oDBManager, sCodiceISTAT, sAnnoRif) = -1 Then
        '        Return False
        '    End If
        '    Log.Debug("ServiceOPENae::DeleteTabAppoggio::fine procedura")
        '    Return True
        'Catch Err As Exception
        '    Log.Debug("Si è verificato un errore in ServiceOPENae::DeleteTabAppoggio::" & Err.Message)
        '    Log.Warn("Si è verificato un errore in ServiceOPENae::DeleteTabAppoggio::" & Err.Message)
        '    Return False
        'End Try
    End Function

    Private Function ListFlussi(ByVal dvFlussi As DataView, ByVal nList As Integer, ByRef oListFlussi() As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE) As Boolean
        Try
            Dim oFlusso As New AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE

            oFlusso.Anno = CStr(dvFlussi.Item(nList)("ANNO"))
            oFlusso.CodiceISTAT = CStr(dvFlussi.Item(nList)("CODICE_ISTAT"))
            If Not IsDBNull(dvFlussi.Item(nList)("data_estrazione")) Then
                oFlusso.DataEstrazione = dvFlussi.Item(nList)("DATA_ESTRAZIONE")
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
            Log.Debug("Si è verificato un errore in ServiceOPENae::GetFlussiTracciati::ListFlussi::" & Err.Message)
            Log.Warn("Si è verificato un errore in ServiceOPENae::GetFlussiTracciati::ListFlussi::" & Err.Message)
            Return False
        End Try
    End Function
End Class
