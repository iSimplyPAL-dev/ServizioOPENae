Imports RIBESFrameWork
Imports System.Web.HttpContext

Public Class CreateSessione

  Dim m_Parametro As String
  Dim m_UserName As String
  Dim m_IdentificativoApplicazione As String

    Public oSession As RIBESFrameWork.Session
    Public oSM As New RIBESFrameWork.SessionManager()
    Public oOM As New RIBESFrameWork.OperationManager()
  '**********************************************************
  'Costruttore della Classe
  'Parametri : 
  'Parametro: parametro del File Env dalla quale si accede al dataBase di WorkFlow
  'UserName: Utente che a i diritti di accesso al DataBase Applicativo
  '**********************************************************
  Public Sub New(ByVal parametro As String, ByVal username As String, ByVal IdentificativoApplicazione As String)

	m_Parametro = parametro
	m_UserName = username
	m_IdentificativoApplicazione = IdentificativoApplicazione

  End Sub



  Public Function CreaSessione(ByVal username As String, ByRef Errore As String) As Boolean
        oSM.sSessionManagerEnvSuffix = m_Parametro
        oOM.sOperationManagerEnvSuffix = m_Parametro

        If Not oSM.Initialize(username, m_Parametro) Then

            GoTo oSMInizialize

        End If

        oSession = oSM.CreateSession(m_IdentificativoApplicazione)

        If oSession Is Nothing Then

            GoTo ErrorSession

        End If

        If Not oOM.Initialize(oSession) Then

            GoTo oOMInitialize

        End If

        CreaSessione = True

        Exit Function

oSMInizialize:

        CreaSessione = False

        Errore = oSM.oErr.Description

        Exit Function

ErrorSession:

        CreaSessione = False

        Errore = oSM.oErr.Description

        Exit Function

oOMInitialize:

        CreaSessione = False

        Errore = oSession.oErr.Description

        Exit Function

  End Function

  Public Sub Kill()


  End Sub



End Class




