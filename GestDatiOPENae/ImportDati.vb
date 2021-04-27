Imports log4net

Public Class ImportDati
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ImportDati))
    Private AppReader As New System.Configuration.AppSettingsReader

    Public Sub New()
        log.Debug("Istanziata la classe ImportDati")
    End Sub

    Public Function PopolaTabAppoggioAE(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String) As Boolean
        Try
            Dim FncPopola As New ClsTabAppoggio
            log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaQuery::inizio procedura")
            'svuoto la tabella
            If FncPopola.DeleteTabAppoggio(sTributo, sAnnoRif, sCodiceISTAT) = False Then
                Return False
            End If
            'richiamo la procedura di inserimento dati
            If FncPopola.SetTabAppoggio(sTributo, sAnnoRif, sCodiceISTAT) = False Then
                Return False
            End If
            log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaQuery::fine procedura")
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::PopolaTabAppoggioAE::PopolaDaQuery::" & Err.Message)
            Return False
        End Try
    End Function

    Public Function PopolaTabAppoggioAE(ByVal oDati() As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Boolean
        Try
            Dim FncPopola As New ClsTabAppoggio
            log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti::inizio procedura")
            'controllo se ho dei dati da inserire
            If oDati.GetLength(0) > 0 Then
                'svuoto la tabella
                If FncPopola.DeleteTabAppoggio(oDati(0).sTributo, oDati(0).sAnno, oDati(0).sCodISTAT) = False Then
                    Return False
                End If
                'richiamo la procedura di inserimento dati
                If FncPopola.SetTabAppoggio(oDati) = False Then
                    Return False
                End If
            Else
                log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti:: array di dati vuoto")
                Return False
            End If
            log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti::fine procedura")
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti::" & Err.Message)
            Return False
        End Try
    End Function

    Public Function PopolaTabAppoggioAE(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByVal sFileImport As String, ByVal sProvenienza As String) As Boolean
        Try
            Dim FncPopola As New ClsTabAppoggio
            log.Debug("ImportDati::PopolaTabAppoggioAE::PopolaDaFile::inizio procedura")
            'sostituisco l'indirizzo http con il percorso fisico
            sFileImport = sFileImport.ToUpper.Replace(UCase(AppReader.GetValue("PathTracciatiForDownload", GetType(String))), UCase(AppReader.GetValue("PathTracciati", GetType(String))))
            log.Debug("GestDatiOPENae::PopolaTabAppoggioAE::PopolaDaFile:: file da acquisire::" & sFileImport)
            'controllo se ho dei dati da inserire
            If FncPopola.CheckFile(sFileImport) = True Then
                'svuoto la tabella
                If FncPopola.DeleteTabAppoggio(sTributo, sAnnoRif, sCodiceISTAT) = False Then
                    Return False
                End If
                Select Case sProvenienza
                    Case "O"
                        'richiamo la procedura di inserimento dati
                        If FncPopola.SetTabAppoggio(sCodiceISTAT, sFileImport) = False Then
                            Return False
                        End If
                    Case "R"
                        'richiamo la procedura di inserimento dati
                        If FncPopola.SetTabAppoggio(sTributo, sAnnoRif, sCodiceISTAT, sFileImport) = False Then
                            Return False
                        End If
                End Select
            Else
                log.Debug("GestDatiOPENae::PopolaTabAppoggioAE::PopolaDaFile:: file non correttamente formattato")
                Return False
            End If
            log.Debug("ImportDati::PopolaTabAppoggioAE::PopolaDaFile::fine procedura")
            Return True
        Catch Err As Exception
            log.Debug("Si è verificato un errore in ImportDati::PopolaTabAppoggioAE::PopolaDaFile::" & Err.Message)
            Return False
        End Try
    End Function
End Class
