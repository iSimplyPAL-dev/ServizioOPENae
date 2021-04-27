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
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(176), "&#x176;") '°
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(124), "&#x124;") '|
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(248), "&#x248;") 'ø

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(224), "a" & "&apos;") 'à
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(225), "a" & "&apos;") 'á()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(226), "a") 'â()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(227), "a") 'ã()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(228), "a") 'ä()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(229), "a") 'å()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(230), "ae") 'æ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(192), "A" & "&apos;") 'À
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(193), "A" & "&apos;") 'Á
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(194), "A") 'Â
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(195), "A") 'Ã
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(196), "A") 'Ä
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(197), "A") 'Å
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(198), "AE") 'Æ

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(233), "e" & "&apos;") 'é()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(232), "e" & "&apos;") 'è()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(234), "e")  'ê()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(235), "e")  'ë()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(200), "E" & "&apos;") 'È()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(201), "E" & "&apos;") 'É()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(202), "E")  'Ê()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(203), "E")  'Ë()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(236), "i" & "&apos;") 'ì()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(237), "i" & "&apos;") 'í()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(238), "i")  'î()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(239), "i")  'ï()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(204), "I" & "&apos;") 'Ì()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(205), "I" & "&apos;") 'Í()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(206), "I")  'Î()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(207), "I")  'Ï()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(242), "o" & "&apos;") 'ò()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(243), "o" & "&apos;") 'ó()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(244), "o")  'ô()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(245), "o")  'õ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(246), "o")  'ö()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(210), "O" & "&apos;") 'Ò()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(211), "O" & "&apos;") 'Ó()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(212), "O")  'Ô()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(213), "O")  'Õ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(214), "O")  'Ö()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(249), "u" & "&apos;") 'ù()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(250), "u" & "&apos;") 'ú()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(251), "u") 'û()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(252), "u") 'ü()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(217), "U" & "&apos;") 'Ú()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(218), "U" & "&apos;") 'Ù()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(219), "U") 'Û()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(220), "U") 'Ü()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(199), "C") 'Ç()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(231), "c") 'ç()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(209), "N") 'Ñ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(241), "n") 'ñ()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(221), "Y") 'Ý()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(253), "y") 'ý()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(255), "y") 'ÿ()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(223), "ss") 'ß()
        End If
        FormattaPerXML = sCampoDaTrattare
    End Function
End Class
