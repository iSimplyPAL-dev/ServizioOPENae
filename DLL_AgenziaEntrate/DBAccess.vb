Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Configuration

Namespace DLL

  Friend Class getDBobject
	Dim _const As New Costanti()
	Public Sub New()
	End Sub

	Private m_Connection As String

	Public Sub New(ByVal _ConnectionString As String)

	  'Se non si usa la stringa di connessione di  default del WebConfig
	  m_Connection = _ConnectionString

	End Sub


	Protected Function GetConnection() As SqlConnection

	  Dim ret_conn As SqlConnection

	  If m_Connection = String.Empty Then
		ret_conn = New SqlConnection(ConfigurationSettings.AppSettings("connectString"))
	  Else
		ret_conn = New SqlConnection(m_Connection)
	  End If

	  ret_conn.Open()
	  GetConnection = ret_conn

	End Function


	Protected Sub CloseConnection(ByVal conn As SqlConnection)
	  conn.Close()
	  conn = Nothing
	End Sub


	Public Function GetDataReader(ByVal strSQL As String) As SqlDataReader

	  Dim cn As SqlConnection = GetConnection()
	  Dim rdr As SqlDataReader

	  Dim cmd As New SqlCommand(strSQL, cn)
	  rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
	  cmd.Dispose()

	  Return rdr

	End Function

	Public Function RunActionQueryIdentiy(ByVal strSQL As String) As Integer

	  Dim cn As SqlConnection = GetConnection()
	  Dim cmd As New SqlCommand(strSQL, cn)
	  Dim IDValue As Integer
	  Try
		IDValue = cmd.ExecuteScalar()
		cmd.Dispose()
	  Finally
		CloseConnection(cn)
	  End Try
	  Return IDValue
	End Function

	Public Sub RunActionQuery(ByVal strSQL As String)

	  Dim cn As SqlConnection = GetConnection()
	  Dim cmd As New SqlCommand(strSQL, cn)

	  Try
		cmd.ExecuteNonQuery()
		cmd.Dispose()
	  Finally
		CloseConnection(cn)
	  End Try

	End Sub

	Public Sub RunSP(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter)

	  Dim cn As SqlConnection = GetConnection()
	  Dim retVal As Integer

	  Try

		Dim cmd As New SqlCommand(strSP, cn)
		cmd.CommandType = CommandType.StoredProcedure

		Dim p As SqlParameter
		For Each p In commandParameters
		  p = cmd.Parameters.Add(p)
		  p.Direction = ParameterDirection.Input
		Next


		cmd.ExecuteNonQuery()

		cmd.Dispose()

	  Finally
		CloseConnection(cn)
	  End Try
	End Sub
	Public Overloads Function RunSPReturnRS(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As SqlDataReader

	  Dim cn As SqlConnection = GetConnection()
	  Dim rdr As SqlDataReader

	  Dim cmd As New SqlCommand(strSP, cn)
	  cmd.CommandType = CommandType.StoredProcedure

	  Dim p As SqlParameter
	  For Each p In commandParameters
		p = cmd.Parameters.Add(p)
		p.Direction = ParameterDirection.Input
	  Next

	  rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
	  cmd.Dispose()

	  Return rdr

	End Function

	Public Function RunSPReturnInteger(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As Integer

	  Dim cn As SqlConnection = GetConnection()
	  Dim retVal As Integer

	  Try

		Dim cmd As New SqlCommand(strSP, cn)
		cmd.CommandType = CommandType.StoredProcedure

		Dim p As SqlParameter
		For Each p In commandParameters
		  p = cmd.Parameters.Add(p)
		  p.Direction = ParameterDirection.Input
		Next

		p = cmd.Parameters.Add(New SqlParameter("@RetVal", SqlDbType.Int))
		p.Direction = ParameterDirection.Output

		cmd.ExecuteNonQuery()
		retVal = cmd.Parameters("@RetVal").Value
		cmd.Dispose()

	  Finally
		CloseConnection(cn)
	  End Try

	  Return retVal

	End Function

	Public Function RunSPReturnRowCount(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As Integer

	  Dim cn As SqlConnection = GetConnection()
	  Dim retVal As Integer

	  Try

		Dim cmd As New SqlCommand(strSP, cn)
		cmd.CommandType = CommandType.StoredProcedure

		Dim p As SqlParameter
		For Each p In commandParameters
		  p = cmd.Parameters.Add(p)
		  p.Direction = ParameterDirection.Input
		Next

		p = cmd.Parameters.Add("@RowCount", SqlDbType.Int)
		p.Direction = ParameterDirection.ReturnValue



		cmd.ExecuteNonQuery()
		retVal = cmd.Parameters("@RowCount").Value
		cmd.Dispose()

	  Finally
		CloseConnection(cn)
	  End Try

	  Return retVal

	End Function
	Public Function RunSPReturnDataSet(ByVal strSP As String, ByVal DataTableName As String, ByVal ParamArray commandParameters() As SqlParameter) As DataSet

	  Dim cn As SqlConnection = GetConnection()

	  Dim ds As New DataSet()

	  Dim da As New SqlDataAdapter(strSP, cn)
	  da.SelectCommand.CommandType = CommandType.StoredProcedure

	  Dim p As SqlParameter
	  For Each p In commandParameters
		da.SelectCommand.Parameters.Add(p)
		p.Direction = ParameterDirection.Input
	  Next

	  da.Fill(ds, DataTableName)

	  CloseConnection(cn)
	  da.Dispose()

	  Return ds

	End Function

	Public Function RunSQLReturnDataSet(ByVal strSql As String) As DataSet

	  Dim cn As SqlConnection = GetConnection()

	  Dim ds As New DataSet()

	  Dim da As New SqlDataAdapter(strSql, cn)
	  da.SelectCommand.CommandType = CommandType.Text

	  da.Fill(ds)

	  CloseConnection(cn)
	  da.Dispose()

	  Return ds

	End Function
	Public Function RunSQLReturnDataAdapter(ByVal strSql As String) As SqlDataAdapter

	  Dim cn As SqlConnection = GetConnection()

	  Dim da As New SqlDataAdapter(strSql, cn)
	  da.SelectCommand.CommandType = CommandType.Text
	  CloseConnection(cn)

	  Return da

	End Function

	Public Function GetConnectionGrid() As SqlConnection

	  Dim ret_conn As SqlConnection

	  If m_Connection = String.Empty Then
		ret_conn = New SqlConnection(ConfigurationSettings.AppSettings("connectString"))
	  Else
		ret_conn = New SqlConnection(m_Connection)
	  End If

	  ret_conn.Open()
	  GetConnectionGrid = ret_conn

	End Function

	Public Function GetNewId(ByRef strNomeTabella As String) As Long

	  'ANTONELLO
	  'funzione che estrae il nuovo ID
	  Dim strSql As String
	  Dim sqlTrans As SqlTransaction
	  Dim lngMaxId As Long
	  Dim oComm As SqlCommand
	  Dim ret_conn As SqlConnection
	  ret_conn = New SqlConnection(ConfigurationSettings.AppSettings("connectString"))

	  ret_conn.Open()
	  sqlTrans = ret_conn.BeginTransaction(IsolationLevel.Serializable)
	  Try

		strSql = "SELECT MAXID FROM CONTATORI  WHERE NOME_TABELLA ='" & strNomeTabella & "'"
		oComm = New SqlCommand(strSql, ret_conn, sqlTrans)
		Dim dr As SqlDataReader = oComm.ExecuteReader
		If dr.Read Then
		  lngMaxId = dr.Item("MAXID")
		  lngMaxId = lngMaxId + _const.VALUE_INCREMENT
		End If
		dr.Close()
		strSql = "UPDATE CONTATORI SET MAXID=" & lngMaxId & " WHERE NOME_TABELLA ='" & strNomeTabella & "'"
		oComm = New SqlCommand(strSql, ret_conn, sqlTrans)
		oComm.ExecuteNonQuery()
		sqlTrans.Commit()
	  Catch ex As Exception
		sqlTrans.Rollback()
		Throw
	  Finally
		oComm.Dispose()
		ret_conn.Close()
	  End Try

	  GetNewId = lngMaxId

	End Function


	Public Function RunSPReturnToGrid(ByVal strSP As String, _
	  ByRef oConn As SqlConnection, _
	  ByRef oComm As SqlCommand, _
	  ByVal ParamArray commandParameters() As SqlParameter) As Integer
	  '///Utilizzata per popolare una griglia da una storedprocedure
	  '///Deve tornare un oggettocommand,e un  oggetto connection

	  Dim cn As SqlConnection = GetConnection()
	  Dim retVal As Integer

	  oConn = cn
	  Dim cmd As New SqlCommand(strSP, cn)
	  cmd.CommandType = CommandType.StoredProcedure

	  Dim p As SqlParameter
	  For Each p In commandParameters
		p = cmd.Parameters.Add(p)
		p.Direction = ParameterDirection.Input
	  Next

	  p = cmd.Parameters.Add("@RowCount", SqlDbType.Int)
	  p.Direction = ParameterDirection.ReturnValue



	  cmd.ExecuteNonQuery()
	  oComm = cmd
	  retVal = cmd.Parameters("@RowCount").Value

	  Return retVal


	End Function
  End Class

End Namespace
