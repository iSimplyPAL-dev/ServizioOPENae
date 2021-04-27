Imports System.ServiceProcess
Imports log4net
Imports log4net.Config
Imports System
Imports System.Runtime.Remoting
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Collections
Imports System.IO
Imports System.Configuration

Namespace ServizioOPENae
    Public Class ServizioOPENae
        Inherits System.ServiceProcess.ServiceBase
        Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ServizioOPENae))
        'true --> quando si deve buildare il servizio
        'false --> quando si vuole lanciare in console per il debug
        Private Shared _runService As Boolean = True
        Private chan As TcpChannel
        Private httpChan As HttpChannel

#Region " Component Designer generated code "

        Public Sub New()
            MyBase.New()

            ' This call is required by the Component Designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call

        End Sub

        'UserService overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        ' The main entry point for the process
        <MTAThread()> _
        Shared Sub Main()
            Dim ServicesToRun() As System.ServiceProcess.ServiceBase

            ' More than one NT Service may run within the same process. To add
            ' another service to this process, change the following line to
            ' create a second service object. For example,
            '
            '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
            '
            If _runService = True Then
                ServicesToRun = New System.ServiceProcess.ServiceBase() {New ServizioOPENae}
                System.ServiceProcess.ServiceBase.Run(ServicesToRun)
            Else
                Dim oServizio As New ServizioOPENae
                oServizio.OnStart(Nothing)
                Console.WriteLine("andato...")
                Console.ReadLine()
            End If
        End Sub

        'Required by the Component Designer
        Private components As System.ComponentModel.IContainer

        ' NOTE: The following procedure is required by the Component Designer
        ' It can be modified using the Component Designer.  
        ' Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container
            Me.ServiceName = "Service1"
        End Sub

#End Region

        Protected Overrides Sub OnStart(ByVal args() As String)
            ' Add code here to start your service. This method should set things
            ' in motion so your service can do its work.
            Dim pathfileinfo As String = ConfigurationSettings.AppSettings("pathfileconflog4net").ToString()
            Dim fileconfiglog4net As New FileInfo(pathfileinfo)
            XmlConfigurator.ConfigureAndWatch(fileconfiglog4net)

            RegisterService()
        End Sub

        Protected Overrides Sub OnStop()
            ' Add code here to perform any tear-down necessary to stop your service.
            ChannelServices.UnregisterChannel(chan)
        End Sub

        Private Shared Sub RegisterService()
            ' Use the configuration file. 
            RemotingConfiguration.Configure(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)

            ' Check to see if we have full errors. 

            'string s = "Errore eccezioni"; 
            If RemotingConfiguration.CustomErrorsEnabled(False) = True Then
            End If

            Console.WriteLine("Inizializzazione Servizio Remoto")
            Dim clientProvider As New BinaryClientFormatterSinkProvider
            Dim serverProvider As New BinaryServerFormatterSinkProvider
            serverProvider.TypeFilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Full
            Dim props As IDictionary = New Hashtable
            props("port") = ConfigurationSettings.AppSettings("TCP_PORT").ToString()
            '50010; 
            'props["typeFilterLevel"] = TypeFilterLevel.Full; 
            Dim chan As New TcpChannel(props, clientProvider, serverProvider)

            props("port") = ConfigurationSettings.AppSettings("HTTP_PORT").ToString()
            ' 50011; 
            Dim clientProviderSoap As New SoapClientFormatterSinkProvider
            Dim serverProviderSoap As New SoapServerFormatterSinkProvider
            serverProviderSoap.TypeFilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Full


            Dim httpChan As New HttpChannel(props, Nothing, Nothing)

            log.Debug("Registrazione Canale")
            ChannelServices.RegisterChannel(chan)
            ChannelServices.RegisterChannel(httpChan)

            RemotingConfiguration.RegisterWellKnownServiceType(GetType(OPENaeService), "OPENaeService.rem", WellKnownObjectMode.SingleCall)
            log.Debug("Registrato OPENaeService")

        End Sub
    End Class
End Namespace