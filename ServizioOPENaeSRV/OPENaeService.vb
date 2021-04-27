Imports log4net
Imports GestDatiOPENae

Namespace ServizioOPENae
    ''' <summary>
    ''' Classe di gestione dei metodi resi dispoibili dal servizio
    ''' </summary>
    Public Class OPENaeService
        Inherits MarshalByRefObject : Implements IGestDatiOPENae
        Private Shared ReadOnly log As ILog = LogManager.GetLogger("OPENaeService")
        Private FncGen As New General

        Public Sub New()
            log.Debug("Istanziata la classe OPENae Service")
        End Sub

        Public Function PopolaTabAppoggioAE(ByVal oDati() As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Boolean Implements GestDatiOPENae.IGestDatiOPENae.PopolaTabAppoggioAE
            Try
                log.Debug("OPENaeService::PopolaTabAppoggioAE:: inizio procedura")
                log.Info("OPENaeService::PopolaTabAppoggioAE:: inizio procedura")

                Dim FncElabDati As New ImportDati

                Return FncElabDati.PopolaTabAppoggioAE(oDati)
                log.Debug("OPENaeService::PopolaTabAppoggioAE:: fine procedura")
            Catch ex As Exception
                Throw New Exception(ex.Message & "::" & ex.StackTrace)
            End Try
        End Function
        Public Function PopolaTabAppoggioAE(ByVal sTributo As String, ByVal sCodiceISTAT As String, ByVal sAnnoRif As String, ByVal sFileImport As String, ByVal sProvenienza As String) As Boolean Implements GestDatiOPENae.IGestDatiOPENae.PopolaTabAppoggioAE
            Try
                log.Debug("OPENaeService::PopolaTabAppoggioAE:: inizio procedura")
                log.Info("OPENaeService::PopolaTabAppoggioAE:: inizio procedura")

                Dim FncElabDati As New ImportDati

                Select Case sTributo
                    Case FncGen.TRIBUTO_ICI
                        Return FncElabDati.PopolaTabAppoggioAE(sTributo, sAnnoRif, sCodiceISTAT, sFileImport, sProvenienza)
                    Case Else
                        Return FncElabDati.PopolaTabAppoggioAE(sTributo, sAnnoRif, sCodiceISTAT)
                End Select
                log.Debug("OPENaeService::PopolaTabAppoggioAE:: fine procedura")
            Catch ex As Exception
                Throw New Exception(ex.Message & "::" & ex.StackTrace)
            End Try
        End Function
        Public Function EstraiTracciato(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByRef sNomeFileTracciati As String) As String Implements IGestDatiOPENae.EstraiTracciato
            Try
                log.Debug("ServiceOPENae::EstraiTracciatoAE::inizio procedura")
                log.Info("ServiceOPENae::EstraiTracciatoAE::inizio procedura")

                Dim FncElabDati As New ExportDati
                Dim sFileEstratto As String = ""
                'richiamo la procedura di estrazione dati
                sFileEstratto = FncElabDati.EstraiTracciato(sTributo, sAnnoRif, sCodiceISTAT, sNomeFileTracciati)

                log.Debug("ServiceOPENae::EstraiTracciatoAE::fine procedura")
                Return sFileEstratto
            Catch ex As Exception
                Throw New Exception(ex.Message & "::" & ex.StackTrace)
            End Try
        End Function
        Public Function EstraiTracciato(ByVal sCodiceISTAT As String, ByVal sCodBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByRef sNomeFileTracciati As String) As String Implements IGestDatiOPENae.EstraiTracciato
            Try
                log.Debug("ServiceOPENae::EstraiTracciatoAE::inizio procedura")
                log.Info("ServiceOPENae::EstraiTracciatoAE::inizio procedura")

                Dim FncElabDati As New ExportDati
                Dim sFileEstratto As String = ""
                'richiamo la procedura di estrazione dati
                sFileEstratto = FncElabDati.EstraiTracciato(sCodiceISTAT, sCodBelfiore, sDescrEnte, sCAPEnte, sTributo, sAnnoRif, sDataScadenza, nProgInvio, sNomeFileTracciati)

                log.Debug("ServiceOPENae::EstraiTracciatoAE::fine procedura")
                Return sFileEstratto
            Catch ex As Exception
                Throw New Exception(ex.Message & "::" & ex.StackTrace)
            End Try
        End Function
        Public Function GetFlussiTracciati(ByVal sTributo As String, ByVal sCodiceISTAT As String) As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE() Implements GestDatiOPENae.IGestDatiOPENae.GetFlussiTracciati
            Try
                log.Debug("OPENaeService::GetFlussiTracciati:: inizio procedura")

                Dim FncElabDati As New ClsTabAppoggio

                Return FncElabDati.GetFlussiTracciati(sTributo, sCodiceISTAT)
                log.Debug("OPENaeService::GetFlussiTracciati:: fine procedura")
            Catch ex As Exception
                Throw New Exception(ex.Message & "::" & ex.StackTrace)
            End Try
        End Function
    End Class
End Namespace
