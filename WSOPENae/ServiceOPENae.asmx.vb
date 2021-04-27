Imports System.Web.Services
Imports System.Data.SqlClient
Imports log4net
Imports System.IO


<System.Web.Services.WebService(Namespace:="http://localhost/OPENaeWS/")> _
Public Class ServiceOPENae
    Inherits System.Web.Services.WebService
    Private Shared Log As ILog = LogManager.GetLogger("ServiceOPENae")

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    ' WEB SERVICE EXAMPLE
    ' The HelloWorld() example service returns the string Hello World.
    ' To build, uncomment the following lines then save and build the project.
    ' To test this web service, ensure that the .asmx file is the start page
    ' and press F5.
    '
    '<WebMethod()> _
    'Public Function HelloWorld() As String
    '   Return "Hello World"
    'End Function

    <WebMethod(Description:="Popola la tabella d'appoggio usando una query diretta", MessageName:="PopolaDaQuery")> _
    Public Function PopolaTabAppoggioAE(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String) As Boolean
        Try
            Log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaQuery::inizio procedura")
            'devo  richiamare il servizio per l'acquisizione
            Dim TypeOfRI As Type = GetType(GestDatiOPENae.IGestDatiOPENae)
            Dim oGestDatiAE As GestDatiOPENae.IGestDatiOPENae
            Dim AppReader As New System.Configuration.AppSettingsReader

            oGestDatiAE = Activator.GetObject(TypeOfRI, AppReader.GetValue("URLServizioGestDatiOPENae", GetType(String)))
            If oGestDatiAE.PopolaTabAppoggioAE(sTributo, sCodiceISTAT, sAnnoRif, "", "") = False Then
                Return False
            End If

            Log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaQuery::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::PopolaTabAppoggioAE::PopolaDaQuery::" & Err.Message)
            Return False
        End Try
    End Function

    <WebMethod(Description:="Popola la tabella d'appoggio ciclando sull'array di oggetti in input", MessageName:="PopolaDaOggetti")> _
    Public Function PopolaTabAppoggioAE(ByVal oDati() As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Boolean
        Try
            Log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti::inizio procedura")
            'devo  richiamare il servizio per l'acquisizione
            Dim TypeOfRI As Type = GetType(GestDatiOPENae.IGestDatiOPENae)
            Dim oGestDatiAE As GestDatiOPENae.IGestDatiOPENae
            Dim AppReader As New System.Configuration.AppSettingsReader

            oGestDatiAE = Activator.GetObject(TypeOfRI, AppReader.GetValue("URLServizioGestDatiOPENae", GetType(String)))
            If oGestDatiAE.PopolaTabAppoggioAE(oDati) = False Then
                Return False
            End If
            Log.Debug("ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti::fine procedura")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::PopolaTabAppoggioAE::PopolaDaOggetti::" & Err.Message)
            Return False
        End Try
    End Function

    <WebMethod()> _
    Public Function EstraiTracciatoAE(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByRef sNomeFileTracciati As String) As String
        Try
            Log.Debug("ServiceOPENae::EstraiTracciatoAE::inizio procedura")
            'devo  richiamare il servizio per l'estrazione
            Dim TypeOfRI As Type = GetType(GestDatiOPENae.IGestDatiOPENae)
            Dim oGestDatiAE As GestDatiOPENae.IGestDatiOPENae
            Dim AppReader As New System.Configuration.AppSettingsReader
            Dim sFileTracciato As String

            oGestDatiAE = Activator.GetObject(TypeOfRI, AppReader.GetValue("URLServizioGestDatiOPENae", GetType(String)))
            sFileTracciato = oGestDatiAE.EstraiTracciato(sTributo, sAnnoRif, sCodiceISTAT, sNomeFileTracciati)
            Log.Debug("ServiceOPENae::EstraiTracciatoAE::fine procedura")

            Return sFileTracciato
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciatoAE::" & Err.Message)
            Return ""
        End Try
    End Function

    <WebMethod(Description:="Preleva l'elenco dei flussi elaborati restituendoli in array di oggetti", MessageName:="GetFlussiTracciati")> _
    Public Function GetFlussiTracciatiAE(ByVal sTributo As String, ByVal sCodiceISTAT As String, ByRef sMyErr As String) As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE()
        Try
            Dim TypeOfRI As Type = GetType(GestDatiOPENae.IGestDatiOPENae)
            Dim oGestDatiAE As GestDatiOPENae.IGestDatiOPENae
            Dim AppReader As New System.Configuration.AppSettingsReader
            Dim oListFlussi() As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE

            Log.Debug("ServiceOPENae::GetFlussiTracciatiAE::inizio procedura")
            oGestDatiAE = Activator.GetObject(TypeOfRI, AppReader.GetValue("URLServizioGestDatiOPENae", GetType(String)))
            oListFlussi = oGestDatiAE.GetFlussiTracciati(sTributo, sCodiceISTAT)
            Log.Debug("ServiceOPENae::GetFlussiTracciatiAE::fine procedura")
            Return oListFlussi
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::GetFlussiTracciatiAE::" & Err.Message)
            Return Nothing
        End Try
    End Function

#Region "ICI"
    <WebMethod(Description:="Popola la tabella d'appoggio usando un file ICI", MessageName:="PopolaDaFile")> _
    Public Function PopolaTabAppoggioAE(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByVal sFileImport As String, ByVal sProvenienza As String) As Boolean
        Try
            Log.Debug("ServiceOPENae::PopolaTabAppoggioAE::inizio procedura")
            WriteFile("H:\SITI Web locali\WSOPENae\LOG\Debug\MyLog.log", "ServiceOPENae::PopolaTabAppoggioAE::inizio procedura", "")

            If sFileImport <> "" Then
                Dim TypeOfRI As Type = GetType(GestDatiOPENae.IGestDatiOPENae)
                Dim oGestDatiAE As GestDatiOPENae.IGestDatiOPENae
                Dim AppReader As New System.Configuration.AppSettingsReader

                oGestDatiAE = Activator.GetObject(TypeOfRI, AppReader.GetValue("URLServizioGestDatiOPENae", GetType(String)))
                If oGestDatiAE.PopolaTabAppoggioAE(sTributo, sCodiceISTAT, sAnnoRif, sFileImport, sProvenienza) = False Then
                    Return False
                End If
            End If
            Log.Debug("ServiceOPENae::PopolaTabAppoggioAE::fine procedura")
            WriteFile("H:\SITI Web locali\WSOPENae\LOG\Debug\MyLog.log", "ServiceOPENae::PopolaTabAppoggioAE::fine procedura", "")
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::PopolaTabAppoggioAE::" & Err.Message)
            Return False
        End Try
    End Function

    <WebMethod(Description:="Estrai Tracciato ICI restituendo path e nome del file creato", MessageName:="EstraiICI")> _
    Public Function EstraiTracciatoAE(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByVal sCodBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByRef sNomeFileTracciati As String) As String
        Try
            Log.Debug("ServiceOPENae::EstraiTracciatoAE::inizio procedura")
            'devo  richiamare il servizio per l'estrazione
            Dim TypeOfRI As Type = GetType(GestDatiOPENae.IGestDatiOPENae)
            Dim oGestDatiAE As GestDatiOPENae.IGestDatiOPENae
            Dim AppReader As New System.Configuration.AppSettingsReader
            Dim sFileTracciato As String

            oGestDatiAE = Activator.GetObject(TypeOfRI, AppReader.GetValue("URLServizioGestDatiOPENae", GetType(String)))
            sFileTracciato = oGestDatiAE.EstraiTracciato(sCodiceISTAT, sCodBelfiore, sDescrEnte, sCAPEnte, sTributo, sAnnoRif, sDataScadenza, nProgInvio, sNomeFileTracciati)
            Log.Debug("ServiceOPENae::EstraiTracciatoAE::fine procedura")

            Return sFileTracciato
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::EstraiTracciatoAE::" & Err.Message)
            Return ""
        End Try
    End Function
#End Region


    Private Function WriteFile(ByVal sFile As String, ByVal DatiToPrint As String, ByVal ErrWriteFile As String) As Integer
        Dim MyFileToWrite As IO.StreamWriter = IO.File.AppendText(sFile)
        Dim sDatiFile As String = ""

        Try
            sDatiFile = DatiToPrint

            MyFileToWrite.WriteLine(sDatiFile.ToUpper)
            MyFileToWrite.Flush()

            Return 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ServiceOPENae::WriteFile::" & Err.Message)
            ErrWriteFile = Err.Message
            Return 0
        Finally
            MyFileToWrite.Close()
        End Try
    End Function
End Class
