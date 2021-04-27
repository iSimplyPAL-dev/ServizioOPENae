Imports System
Imports RIBESFrameWork
Imports System.Data.SqlClient

Namespace DLL
  Public Class Utility

	Dim _Const As New Costanti()

	Public Shared Function GetParametro(ByRef strInput As Object) As String
	  GetParametro = ""
	  If Not IsDBNull(strInput) And Not IsNothing(strInput) Then
		GetParametro = CStr(strInput)
	  End If
	  Return GetParametro
	End Function

	Public Shared Function cTolng(ByRef objInput As Object) As Long
	  cTolng = 0
	  If Not IsDBNull(objInput) And Not IsNothing(objInput) Then
		If IsNumeric(objInput) Then
		  cTolng = CLng(objInput)
		End If
	  End If
	End Function

	Public Shared Function cToBool(ByRef objInput As Object) As Boolean

	  cToBool = False

	  If Not IsDBNull(objInput) And Not IsNothing(objInput) Then

		On Error Resume Next
		cToBool = CBool(objInput)
	  End If
	End Function


	Public Shared Function cToDate(ByRef objInput As Object) As Date

	  cToDate = System.DateTime.FromOADate(0)

	  If Not IsDBNull(objInput) And Not IsNothing(objInput) Then
		If IsDate(objInput) Then
		  cToDate = CDate(objInput)
		End If

	  End If

	End Function
	Public Function cDateToClient(ByRef vrnDateTime As Object) As String

	  Dim strData As String
	  Dim strTime As String

	  cDateToClient = ""


	  If Not IsDBNull(vrnDateTime) And Not IsNothing(vrnDateTime) Then
		strData = CStr(vrnDateTime)
		If IsDate(strData) Then
		  cDateToClient = Format(Day(CDate(strData)), "00") & "/" & Format(Month(CDate(strData)), "00") & "/" & Format(Year(CDate(strData)), "0000")
		End If
	  End If

	End Function
	Public Function CToStr(ByRef strInput As String) As Object

	  CToStr = System.DBNull.Value

	  If Len(strInput) > 0 Then
		CToStr = strInput
	  End If

	End Function


	Public Function CDateToDB(ByVal vInput As Object, Optional ByRef blnFormatoInputServer As Boolean = False) As String
	  Dim strDate As String
	  Dim intOra As Short
	  Dim intMinuti As Short
	  Dim intSecondi As Short
	  Dim strTime As String
	  Dim sTesto As String

	  CDateToDB = "Null"


	  If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
		sTesto = CStr(vInput)
		'verifica che sia una data valida (il formato non importa!!!)
		If Not IsDate(sTesto) Then
		  Exit Function
		End If
		If blnFormatoInputServer = True Then
		  'Formato in input in inglese, il DB lo vuole in italiano
		  strDate = Format(Day(CDate(sTesto)), "00") & "/" & Format(Month(CDate(sTesto)), "00") & "/" & Format(Year(CDate(sTesto)), "0000")

		Else
		  'Formato ITALIANO : applica solo una formattazione ai campi
		  Dim varAr() As String = Split(sTesto, "/")

		  Dim str As String
		  Dim strDate_test As DateTime

		  sTesto = Day(CDate(sTesto)) & "/" & Month(CDate(sTesto)) & "/" & Year(CDate(sTesto))

		  '    strDate_test = CDate(str)


		  '    strDate = String.Format("{0:dd}", varAr(0))




		  '  strDate = strDate & "/" & String.Format(varAr(1), "00")
		  '  strDate = strDate & "/" & Format(Year(CDate("01/01/" & varAr(2))), "0000")

		  '  strTime = String.Format(sTesto, "HH:mm:ss")

		  '  If Left(strTime, 2) = "00" And Mid(strTime, 4, 2) = "00" And Right(strTime, 2) = "00" Then
		  '    sTesto = strDate
		  '  Else
		  '    sTesto = strDate & " " & Left(strTime, 2) & ":" & Mid(strTime, 4, 2) & ":" & Right(strTime, 2)
		  '  End If
		End If

		'Verifica , dopo aver ricostruito la data , che sia una data valida
		If Not IsDate(sTesto) Then		  'NOTA: il test con IsDate() va comunque bene anche se gli vengono passati GG e MM invertiti purchè i valori siano validi
		  CDateToDB = "Null"
		Else
		  CDateToDB = "'" & sTesto & "'"
		End If
	  End If

	End Function
	Public Function CStrToDB(ByVal vInput As Object, Optional ByRef blnClearSpace As Boolean = False) As String
	  Dim sTesto As String

	  CStrToDB = "''"


	  If Not IsDBNull(vInput) And Not IsNothing(vInput) Then

		sTesto = CStr(vInput)
		If blnClearSpace Then
		  sTesto = Trim(sTesto)
		End If
		If Trim(sTesto) <> "" Then
		  CStrToDB = "'" & Replace(sTesto, "'", "''") & "'"
		End If
	  End If

	End Function

	Public Function CIdToDB(ByVal vInput As Object) As String

	  CIdToDB = "Null"

	  If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
		If IsNumeric(vInput) Then
		  If CDbl(vInput) > 0 Then
			CIdToDB = CStr(CDbl(vInput))
		  End If
		End If
	  End If

	End Function
	Public Function CToBit(ByRef vInput As Object) As Short

	  CToBit = 0

	  If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
		If CBool(vInput) Then
		  CToBit = 1
		Else
		  CToBit = 0
		End If
	  End If

	End Function
	Public Function CIdFromDB(ByVal vInput As Object) As String

	  CIdFromDB = "-1"

	  If Not IsDBNull(vInput) And Not IsNothing(vInput) And Not IsNothing(vInput) Then
		If IsNumeric(vInput) Then
		  If CDbl(vInput) > 0 Then
			CIdFromDB = CStr(CDbl(vInput))
		  End If
		End If
	  End If

	End Function
	Public Shared Function AddBackSlashToPath(ByRef sPath As Object, Optional ByRef blnRemoveInitialSlash As Boolean = True) As String

	  Dim lngErrNumber As Integer

	  'converte la variabile passata in una stringa
	  AddBackSlashToPath = ""
	  If IsDBNull(sPath) Or IsNothing(sPath) Then
		sPath = ""
	  End If

	  sPath = CStr(sPath)

	  AddBackSlashToPath = sPath

	  If Len(sPath) = 0 Then Exit Function

	  'Aggiunge la \ alla fine del path
	  If Right(sPath, 1) <> "\" And Right(sPath, 1) <> "/" Then	'And Len(sPath) > 1 Then

		sPath = sPath & "\"
	  End If
	  If blnRemoveInitialSlash Then

		If Left(sPath, 1) = "\" Or Left(sPath, 1) = "/" Then

		  sPath = Mid(sPath, 2)
		End If
	  End If
	  AddBackSlashToPath = sPath


    End Function


    Public Function NumberToChar(ByVal vInput As Object, ByVal strChar As String, ByVal lngLenght As Long) As String

      Dim mySTR As String
      NumberToChar = "-1"

      If Not IsDBNull(vInput) And Not IsNothing(vInput) And Not IsNothing(vInput) Then
        If Not IsDBNull(strChar) And Not IsNothing(strChar) And Not IsNothing(strChar) Then
          If Not IsDBNull(lngLenght) And Not IsNothing(lngLenght) And Not IsNothing(lngLenght) Then
            If IsNumeric(vInput) Then
              mySTR = vInput
              NumberToChar = mySTR.PadLeft(lngLenght, strChar)
            End If
          End If
        End If
      End If


    End Function


    '======================================================================================
    'FUNZIONE CHE RESTITUISCE UN ID NUMERICO DA UNA TABELLA CONTATORI
    '======================================================================================
    'Public Function GetNewId(ByRef strNomeTabella As String, ByVal objDBAccess As RIBESFrameWork.DBManager) As Long
    Public Function GetNewId(ByRef strNomeTabella As String, ByVal oSession As RIBESFrameWork.Session, ByVal Id_SottoAttivita As String) As Long

      Dim strSql As String
      Dim lngMaxId As Long
      Dim intRetVal As Integer
      Dim objCONST As New Costanti

      Dim objDBAccess As New RIBESFrameWork.DBManager
      Try
        objDBAccess = oSession.GetPrivateDBManager(Id_SottoAttivita)

        objDBAccess.BeginTransIsolationLevel()
        Try
          strSql = "SELECT MAXID FROM CONTATORI  WHERE NOME_TABELLA ='" & strNomeTabella & "'"
          Dim dr As SqlDataReader = objDBAccess.CmdCreateWithTransaction(strSql).ExecuteReader
          If dr.Read Then
            lngMaxId = dr.Item("MAXID")
            lngMaxId = lngMaxId + _Const.VALUE_INCREMENT
          End If
          dr.Close()
          strSql = "UPDATE CONTATORI SET MAXID=" & lngMaxId & " WHERE NOME_TABELLA ='" & strNomeTabella & "'"
          objDBAccess.CmdCreateWithTransaction(strSql)
          intRetVal = objDBAccess.CmdExec()
          If intRetVal = objCONST.INIT_VALUE_NUMBER Then
            Throw New Exception("UPDATE CONTATORI")
          End If
          objDBAccess.CommitTrans()
        Catch ex As Exception
          objDBAccess.RollbackTrans()
          Throw New Exception("Errore di accesso tabella CONTATORI")
        End Try

        GetNewId = lngMaxId
      Catch ex As Exception
        Throw New Exception("Anagrafica::GetNewId::" & ex.Message)
      Finally
        '********************Gestione Anagrafiche massive****************************
        objDBAccess.DisposeConnection()
        objDBAccess.Dispose()
        '********************Gestione Anagrafiche massive****************************
      End Try
    End Function
    '======================================================================================
    'FINE FUNZIONE CHE RESTITUISCE UN ID NUMERICO DA UNA TABELLA CONTATORI
    '======================================================================================
  End Class

End Namespace
