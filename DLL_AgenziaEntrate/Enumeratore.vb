Imports System
Namespace DLL
  Public Enum UpdateRecordStatus
	Updated = 0
	Deleted = 1
	Concurrency = 2

  End Enum
  Public Enum InsertRecordStatus
	Insert = 2
	NoInsert = 0
  End Enum
End Namespace