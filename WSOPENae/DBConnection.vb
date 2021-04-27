Imports System
Imports System.Configuration
Imports System.Collections
Imports System.Text
Imports System.Data.SqlClient
'Imports Utility

Public Class DBConnection

    Private AppReader As New System.Configuration.AppSettingsReader

    'Dim oDBManager As DBManager

    Public Sub New()

    End Sub

    'Public Function DBConnection() As DBManager
    '    'se non viene passata la connessione utilizzo quella presente nel Web.Config
    '    oDBManager = New DBManager(CType(AppReader.GetValue("ConnectionStringDB", GetType(String)), String))
    '    Return oDBManager
    'End Function

    'Public Function DBConnection(ByVal sConnectionString As String) As DBManager
    '    oDBManager = New DBManager(sConnectionString)
    '    Return oDBManager
    'End Function

End Class

Public Class General
    Public Const TIPORCTESTA As String = "0"
    Public Const TIPORCDETTAGLIO As String = "1"
    Public Const TIPORCCODA As String = "9"
    Public Const TARSU_IDFORNITURA As String = "SMRIF"
    Public Const TARSU_CODNUMFORNITURA As String = "34"
    Public Const H2O_IDFORNITURA As String = "NWIDR"
    Public Const H2O_CODNUMFORNITURA As String = "24"
    Public Const CHRCONTROLLO As String = "A"
    Public Const CHRASCIIFINERIGA As String = vbCrLf

    Public Function ReplaceDataForTXT(ByVal myString As String) As String
        'leggo la data nel formato aaaammgg  e la metto nel formato GGMMAAAA
        Dim sGiorno As String
        Dim sMese As String
        Dim sAnno As String
        If myString <> "" Then
            sGiorno = Mid(myString, 7, 2)
            sMese = Mid(myString, 5, 2)
            sAnno = Mid(myString, 1, 4)
            ReplaceDataForTXT = sGiorno & sMese & sAnno
        Else
            ReplaceDataForTXT = ""
        End If
    End Function

    Public Function ReplaceDataForDB(ByVal myString As String) As String
        'leggo la data nel formato gg/mm/aaaa e la metto nel formato aaaammgg
        Dim sGiorno As String
        Dim sMese As String
        Dim sAnno As String

        If myString <> "" Then
            sGiorno = Mid(myString, 1, 2)
            sMese = Mid(myString, 4, 2)
            sAnno = Mid(myString, 7, 4)
            ReplaceDataForDB = sAnno & sMese & sGiorno
        Else
            ReplaceDataForDB = ""
        End If
    End Function

    Public Shared Function FormattaPerXML(ByVal sCampoDaTrattare As String) As String
        'dovendo inserire in un file dei caratteri particolari, occorre fare riferimento
        'al loro corrispondente entity-name
        If sCampoDaTrattare <> "" Then
            sCampoDaTrattare = Replace(sCampoDaTrattare, "&", "&amp;")
            sCampoDaTrattare = Replace(sCampoDaTrattare, "<", "&lt;")
            sCampoDaTrattare = Replace(sCampoDaTrattare, ">", "&gt;")
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(34), "&quot;") '"
            sCampoDaTrattare = Replace(sCampoDaTrattare, "'", "&apos;")
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(176), "&#x176;") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(124), "&#x124;") '|
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(248), "&#x248;") '�

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(224), "a" & "&apos;") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(225), "a" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(226), "a") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(227), "a") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(228), "a") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(229), "a") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(230), "ae") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(192), "A" & "&apos;") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(193), "A" & "&apos;") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(194), "A") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(195), "A") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(196), "A") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(197), "A") '�
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(198), "AE") '�

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(233), "e" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(232), "e" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(234), "e")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(235), "e")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(200), "E" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(201), "E" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(202), "E")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(203), "E")  '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(236), "i" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(237), "i" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(238), "i")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(239), "i")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(204), "I" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(205), "I" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(206), "I")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(207), "I")  '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(242), "o" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(243), "o" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(244), "o")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(245), "o")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(246), "o")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(210), "O" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(211), "O" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(212), "O")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(213), "O")  '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(214), "O")  '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(249), "u" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(250), "u" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(251), "u") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(252), "u") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(217), "U" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(218), "U" & "&apos;") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(219), "U") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(220), "U") '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(199), "C") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(231), "c") '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(209), "N") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(241), "n") '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(221), "Y") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(253), "y") '�()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(255), "y") '�()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(223), "ss") '�()
        End If
        FormattaPerXML = sCampoDaTrattare
    End Function
End Class
