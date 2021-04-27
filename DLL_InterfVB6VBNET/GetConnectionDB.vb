
Imports System.Data.OleDb
Imports System.Configuration
Imports System

Imports System.ComponentModel
Imports System.Data
Imports System.Text



Public Class GetConnectionDB

    Public StrConnDBOpen As String
    Public MyCnString As String

    Public Sub New()
        'connessione al DB di RDB
        'AppReader = New System.Configuration.AppSettingsReader

        'MyCnString = ConfigurationSettings.AppSettings("ConnessioneACCESS").ToString()
        'MyCnString = MyCnString & CType(AppReader.GetValue("PathACCESS", GetType(String)), String)
        'MyCnString = MyCnString & CType(AppReader.GetValue("Database", GetType(String)), String)

    End Sub


    Protected Function GetConnectionACCESS(ByVal myconn As String) As OleDbConnection
        Dim MyConnection As OleDbConnection

        MyCnString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="
        MyCnString = MyCnString & "C:\RDB\DATABASE\rdb.mdb"

        MyConnection = New OleDbConnection(myconn)

        MyConnection.Open()
        GetConnectionACCESS = MyConnection
    End Function

    Protected Sub CloseConnectionACCESS(ByVal MyConn As OleDbConnection)
        MyConn.Close()
        MyConn = Nothing
    End Sub

    Public Function GetDataReaderACCESS(ByVal MySQL As String, ByVal myconn As String, ByRef StrErrore As String) As OleDbDataReader
        Dim MyCn As OleDbConnection = GetConnectionACCESS(myconn)
        Dim MyDataReader As OleDbDataReader
        Dim MyCm As New OleDbCommand(MySQL, MyCn)
        MyCm.CommandTimeout = 900

        Try
            MyDataReader = MyCm.ExecuteReader(CommandBehavior.CloseConnection)
            MyCm.Dispose()
            Return MyDataReader
        Catch ex As Exception
            StrErrore = ex.Message
        End Try

    End Function

    Public Sub RunExecuteQueryACCESS(ByVal MySQL As String, ByVal myconn As String)
        Dim MyCn As OleDbConnection = GetConnectionACCESS(myconn)
        Dim MyCm As New OleDbCommand(MySQL, MyCn)
        MyCm.CommandTimeout = 900

        Try
            MyCm.ExecuteNonQuery()
            MyCm.Dispose()
        Catch EX As Exception
            MessageBox.Show("Errore: " & EX.Message, "RUN EXECUTE QUERY", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw EX
        Finally
            CloseConnectionACCESS(MyCn)
        End Try
    End Sub

End Class

