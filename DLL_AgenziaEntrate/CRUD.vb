Imports System
Imports System.Data.SqlClient

Namespace DLL

  <Serializable()> Public Class CRUD
#Region "Abstract methods"

#End Region

#Region "Properties"

	Protected _id As Integer
	Protected _concurrency As DateTime

	Public Property ID() As Integer
	  Get
		Return _id
	  End Get
	  Set(ByVal Value As Integer)
		_id = Value
	  End Set
	End Property

	Public Property Concurrency() As DateTime
	  Get
		Return _concurrency
	  End Get

	  Set(ByVal Value As DateTime)
		_concurrency = Value
	  End Set
	End Property

#End Region
  End Class
End Namespace