Namespace DLL
  Friend Class Costanti

	Public Const INIT_VALUE_NUMBER As Integer = -1
	Public Const INIT_VALUE_STRING As String = ""
	Public Const INIT_VALUE_BOOL As Boolean = False
	Public Const VALUE_NUMBER_ZERO As Integer = 0
	Public Const VALUE_NUMBER_UNO As Integer = 1
    Public Const VALUE_INCREMENT As Integer = 1
    Public Const VALUE_VIRTUALCF_DEFAULT As String = "CFV"

	Public Enum DBOperation
	  DB_INSERT = 1	'NUOVA ANAGRAFICA
	  DB_UPDATE = 0	'MODIFICA ANAGRAFICA
	End Enum

  End Class
End Namespace