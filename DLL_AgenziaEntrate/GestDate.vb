Namespace DLL

  Friend Class GestDate
	Public Function GiraData(ByVal data As String) As String
	  'leggo la data nel formato gg/mm/aaaa e la metto nel formato aaaammgg
	  Dim Giorno As String
	  Dim Mese As String
	  Dim Anno As String

	  If data <> "" Then
		Giorno = Mid(data, 1, 2)
		Mese = Mid(data, 4, 2)
		Anno = Mid(data, 7, 4)
		GiraData = Anno & Mese & Giorno
	  Else
		GiraData = ""
	  End If
	End Function


	Public Function GiraDataFromDB(ByVal data As String) As String
	  'leggo la data nel formato aaaammgg  e la metto nel formato gg/mm/aaaa
	  Dim Giorno As String
	  Dim Mese As String
	  Dim Anno As String
	  If data <> "" Then
		Giorno = Mid(data, 7, 2)
		Mese = Mid(data, 5, 2)
		Anno = Mid(data, 1, 4)
		GiraDataFromDB = Giorno & "/" & Mese & "/" & Anno
	  Else
		GiraDataFromDB = ""
	  End If
	End Function
  End Class
End Namespace