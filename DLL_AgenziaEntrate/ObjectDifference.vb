Imports System
Imports System.Reflection
Imports System.Collections

Namespace DLL
  Public Class ObjectDifference

	

	Public Sub New(ByVal propertyName As String, ByVal firstValue As Object, ByVal secondValue As Object)
	  _propertyName = propertyName
	  _firstValue = firstValue
	  _secondValue = secondValue
	End Sub

	Public Shared Function GetDifferences(ByVal firstObject As Object, ByVal secondObject As Object) As IList

	  Dim firstType As Type = firstObject.GetType()
	  Dim secondType As Type = secondObject.GetType()
	  Dim firstValue As Object
	  Dim secondValue As Object
	  Dim differences As New ArrayList()


	  If Not firstType.Equals(secondType) Then
		Throw New Exception("Gli oggetti devono essere eguali.[GetDifferences] class ObjectDifference.vb")
	  End If
	  Dim prop As PropertyInfo
	  For Each prop In firstType.GetProperties()

		firstValue = prop.GetValue(firstObject, Nothing)
		secondValue = prop.GetValue(secondObject, Nothing)

		If Not IsNothing(firstValue) Then
		  If Not firstValue.Equals(secondValue) Then
			differences.Add(New ObjectDifference(prop.Name, firstValue, secondValue))
		  End If
		End If
	  Next prop

	  Return differences

	End Function


	Private _propertyName As String
	Private _firstValue As Object
	Private _secondValue As Object


	Public Property PropertyName() As String
	  Get
		Return _propertyName
	  End Get
	  Set(ByVal Value As String)
		_propertyName = Value
	  End Set
	End Property

	Public Property FirstValue() As Object
	  Get
		Return _firstValue
	  End Get
	  Set(ByVal Value As Object)
		_firstValue = Value
	  End Set
	End Property

	Public Property SecondValue() As Object
	  Get
		Return _secondValue
	  End Get
	  Set(ByVal Value As Object)
		_secondValue = Value
	  End Set
	End Property
  End Class
End Namespace