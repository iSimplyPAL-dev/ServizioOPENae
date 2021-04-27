Imports System
Imports System.Configuration
Imports System.Collections
Imports System.Text
Imports System.Data.SqlClient
Imports Utility
Imports log4net

Public Class DBConnection
    Private Shared Log As ILog = LogManager.GetLogger("DBConnection")
    Private AppReader As New System.Configuration.AppSettingsReader

    Dim oDBManager As DBModel

    Public Sub New()

    End Sub

    Public Function DBConnection() As DBModel
        'se non viene passata la connessione utilizzo quella presente nel Web.Config
        Log.Debug("DBConnection::stringa di connessione" & CType(AppReader.GetValue("ConnectionStringDB", GetType(String)), String))
        oDBManager = New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, CType(AppReader.GetValue("ConnectionStringDB", GetType(String)), String))
        Return oDBManager
    End Function

    Public Function DBConnection(ByVal sConnectionString As String) As DBModel
        Log.Debug("DBConnection::stringa di connessione" & sConnectionString)
        oDBManager = New DBModel(AgenziaEntrateDLL.AgenziaEntrate.Generale.DBType, sConnectionString)
        Return oDBManager
    End Function

End Class

Public Class General
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(General))
    Public Const TIPORCTESTA As String = "0"
    Public Const TIPORCDETTAGLIO As String = "1"
    Public Const TIPORCCODA As String = "9"
    Public Const ICI_TIPORCTESTA As String = "ICI0"
    Public Const ICI_TIPORCCODA As String = "ICI9"
    Public Const ICI_TIPORCCONTABILEORD As String = "3"
    Public Const ICI_TIPORCCONTABILEVIOL As String = "6"
    Public Const ICI_TIPORCANAGRAFICOFIS As String = "4"
    Public Const ICI_TIPORCANAGRAFICOGIUR As String = "5"
    Public Const ICI_CODCONCESSIONE As String = "1"
    Public Const ICI_CODTESORERIA As String = "0"
    Public Const ICI_RIVERSAMENTOCONTOCORRENTE As String = "POSTE"
    Public Const TARSU_IDFORNITURA As String = "SMRIF"
    Public Const TARSU_CODNUMFORNITURA As String = "34"
    ' Modifica per la variazione del tracciato record
    Public Const H2O_IDFORNITURA As String = "IDR00"
    ' tipo di comunicazione

    Public Const H2O_TIPOCOMUNICAZIONE As String = "0"
    Public Const H2O_CFINTERMEDIARIO As String = ""
    Public Const H2O_PROTCOMUNICAZIONE As String = " "
    Public Const H2O_NUMCAF As String = " "
    Public Const H2O_IMPEGNOTRASMISSIONE As String = " "
    Public Const H2O_DATAIMPEGNO As String = " "

    Public Const H2O_CODNUMFORNITURA As String = "24"
    Public Const CHRCONTROLLO As String = "A"
    Public Const CHRASCIIFINERIGA As String = vbCrLf
    Public Const TRIBUTO_TARSU As String = "0434"
    Public Const TRIBUTO_TIA As String = "0465"
    Public Const TRIBUTO_H2O As String = "9000"
    Public Const TRIBUTO_ICI As String = "8852"


    Public Function GetValParamCmd(ByVal MyCMD As SqlClient.SqlCommand) As String
        Dim sReturn As String
        Dim x As Integer

        For x = 0 To MyCMD.Parameters.Count - 1
            If MyCMD.Parameters(x).DbType = DbType.String Or MyCMD.Parameters(x).DbType = DbType.DateTime Then
                sReturn += "'" + MyCMD.Parameters(x).Value & "',"
            Else
                sReturn += MyCMD.Parameters(x).Value & ","
            End If
        Next
        Return sReturn
    End Function

Public Function ReplaceChar(ByVal myString As String) As String
    Dim sReturn As String

    sReturn = Replace(myString, "'", "''")
    sReturn = Replace(sReturn, "*", "%")
    sReturn = Replace(sReturn, "&nbsp;", " ")
    sReturn = Trim(sReturn)
    Return sReturn
End Function

    Public Function ReplaceDataForTXT(ByVal myString As String, Optional ByVal sCharSep As String = "") As String
        'leggo la data nel formato aaaammgg  e la metto nel formato GGMMAAAA
        Dim sGiorno As String
        Dim sMese As String
        Dim sAnno As String
        If myString <> "" Then
            sGiorno = Mid(myString, 7, 2)
            sMese = Mid(myString, 5, 2)
            sAnno = Mid(myString, 1, 4)
            ReplaceDataForTXT = sGiorno & sCharSep & sMese & sCharSep & sAnno
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
            myString = myString.Replace("/", "")
            sGiorno = Mid(myString, 1, 2)
            sMese = Mid(myString, 3, 2)
            sAnno = Mid(myString, 5, 4)
            ReplaceDataForDB = sAnno & sMese & sGiorno
        Else
            ReplaceDataForDB = ""
        End If
    End Function

    Public Shared Function FormattaPerTXT(ByVal sCampoDaTrattare As String, ByVal nCheckLen As Integer) As String
        'dovendo inserire in un file dei caratteri particolari, occorre fare riferimento
        'al loro corrispondente entity-name
        If sCampoDaTrattare <> "" Then
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(224), "a'") 'à
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(225), "a'") 'á()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(226), "a") 'â()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(227), "a") 'ã()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(228), "a") 'ä()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(229), "a") 'å()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(230), "ae") 'æ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(192), "A'") 'À
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(193), "A'") 'Á
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(194), "A") 'Â
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(195), "A") 'Ã
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(196), "A") 'Ä
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(197), "A") 'Å
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(198), "AE") 'Æ

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(233), "e'") 'é()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(232), "e'") 'è()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(234), "e")  'ê()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(235), "e")  'ë()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(200), "E'") 'È()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(201), "E'") 'É()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(202), "E")  'Ê()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(203), "E")  'Ë()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(236), "i'") 'ì()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(237), "i'") 'í()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(238), "i")  'î()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(239), "i")  'ï()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(204), "I'") 'Ì()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(205), "I'") 'Í()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(206), "I")  'Î()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(207), "I")  'Ï()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(242), "o'") 'ò()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(243), "o'") 'ó()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(244), "o")  'ô()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(245), "o")  'õ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(246), "o")  'ö()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(210), "O'") 'Ò()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(211), "O'") 'Ó()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(212), "O")  'Ô()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(213), "O")  'Õ()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(214), "O")  'Ö()

            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(249), "u'") 'ù()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(250), "u'") 'ú()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(251), "u") 'û()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(252), "u") 'ü()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(217), "U'") 'Ú()
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(218), "U'") 'Ù()
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
            sCampoDaTrattare = Replace(sCampoDaTrattare, Chr(176), "") '°()

            sCampoDaTrattare = Replace(sCampoDaTrattare, "''", "'") '°()
        End If
        If nchecklen > -1 Then
            If sCampoDaTrattare.Length > nchecklen Then
                sCampoDaTrattare = sCampoDaTrattare.Substring(0, nCheckLen)
            End If
        End If
        FormattaPerTXT = sCampoDaTrattare
    End Function

    Public Function DecimalEuro(ByVal ImportoEuro As String) As String
        Dim StrImporto As String
        Dim SegnoNegativo As Boolean = False

        If InStr(1, ImportoEuro, ",") = 0 Then
            StrImporto = CStr(ImportoEuro)
            If Len(StrImporto) < 3 Then
                StrImporto = Left("000" & StrImporto, 3)
                StrImporto = Left(StrImporto, (Len(StrImporto) - 2)) & "," & Right(StrImporto, 2)
            Else
                StrImporto = Left(StrImporto, (Len(StrImporto) - 2)) & "," & Right(StrImporto, 2)
            End If
        Else
            StrImporto = CStr(ImportoEuro)
        End If
        If InStr(StrImporto, "-") <> 0 Then
            StrImporto = Right(StrImporto, Len(StrImporto) - InStr(StrImporto, "-"))
            SegnoNegativo = True
        End If
        If SegnoNegativo = True Then
            DecimalEuro = CDbl(StrImporto) * (-1)
        Else
            DecimalEuro = CDbl(StrImporto)
        End If
    End Function

    Public Function ControlloCinCF(ByVal sCodiceFiscale As String) As Boolean
        Dim Mesi As String
        Dim Vocali As String
        Dim Consonanti As String
        Dim Numeri As String
        Dim PesoVocaliPari As String
        Dim PesoVocaliDispari As String
        Dim PesoConsonantiPari As String
        Dim PesoConsonantiDispari As String
        Dim PesoNumeriPari As String
        Dim PesoNumeriDispari As String
        Dim CFControllo As String
        Dim SumPari As Integer
        Dim SumDispari As Integer
        Dim IDChr As Integer
        Dim ChrControllo As String
        Dim PosChr As Integer
        Dim Resto As Integer
        Dim u As String

        'restituisce TRUE se il CIN è corretto
        Try
            Mesi = "ABCDEHLMPRST"
            Vocali = "AEIOU"
            Consonanti = "BCDFGHJKLMNPQRSTVWXYZ"
            Numeri = "0123456789"
            PesoVocaliPari = Chr(0) + Chr(4) + Chr(8) + Chr(14) + Chr(20)
            PesoConsonantiPari = Chr(1) + Chr(2) + Chr(3) + Chr(5) + Chr(6) + Chr(7) + Chr(9) + Chr(10) + Chr(11) + Chr(12) + Chr(13) + Chr(15) + Chr(16) + Chr(17) + Chr(18) + Chr(19) + Chr(21) + Chr(22) + Chr(23) + Chr(24) + Chr(25)
            PesoNumeriPari = Chr(0) + Chr(1) + Chr(2) + Chr(3) + Chr(4) + Chr(5) + Chr(6) + Chr(7) + Chr(8) + Chr(9)
            PesoVocaliDispari = Chr(1) + Chr(9) + Chr(19) + Chr(11) + Chr(16)
            PesoConsonantiDispari = Chr(0) + Chr(5) + Chr(7) + Chr(13) + Chr(15) + Chr(17) + Chr(21) + Chr(2) + Chr(4) + Chr(18) + Chr(20) + Chr(3) + Chr(6) + Chr(8) + Chr(12) + Chr(14) + Chr(10) + Chr(22) + Chr(25) + Chr(24) + Chr(23)
            PesoNumeriDispari = Chr(1) + Chr(0) + Chr(5) + Chr(7) + Chr(9) + Chr(13) + Chr(15) + Chr(17) + Chr(19) + Chr(21)

            CFControllo = sCodiceFiscale.Substring(0, 15)

            'Somma pari
            SumPari = 0
            For IDChr = 2 To 14 Step 2
                ChrControllo = Mid(CFControllo, IDChr, 1)
                PosChr = InStr(Vocali, ChrControllo)
                If PosChr > 0 Then
                    SumPari = SumPari + Asc(Mid(PesoVocaliPari, PosChr))
                Else
                    PosChr = InStr(Consonanti, ChrControllo)
                    If PosChr > 0 Then
                        SumPari = SumPari + Asc(Mid(PesoConsonantiPari, PosChr))
                    Else
                        PosChr = InStr(Numeri, ChrControllo)
                        If PosChr > 0 Then
                            SumPari = SumPari + Asc(Mid(PesoNumeriPari, PosChr))
                        End If
                    End If
                End If
            Next

            'Somma Dispari
            SumDispari = 0
            For IDChr = 1 To 15 Step 2
                ChrControllo = Mid(CFControllo, IDChr, 1)
                PosChr = InStr(Vocali, ChrControllo)
                If PosChr > 0 Then
                    SumDispari = SumDispari + Asc(Mid(PesoVocaliDispari, PosChr))
                Else
                    PosChr = InStr(Consonanti, ChrControllo)
                    If PosChr > 0 Then
                        SumDispari = SumDispari + Asc(Mid(PesoConsonantiDispari, PosChr))
                    Else
                        PosChr = InStr(Numeri, ChrControllo)
                        If PosChr > 0 Then
                            SumDispari = SumDispari + Asc(Mid(PesoNumeriDispari, PosChr))
                        End If
                    End If
                End If
            Next

            Resto = SumPari + SumDispari
            Do While Resto >= 26
                Resto = Resto - 26
            Loop

            u = "."
            PosChr = InStr(1, PesoVocaliPari, Chr(Resto), 0)
            If PosChr > 0 Then
                u = Mid(Vocali, PosChr, 1)
            Else
                PosChr = InStr(1, PesoConsonantiPari, Chr(Resto), 0)
                If PosChr > 0 Then
                    u = Mid(Consonanti, PosChr, 1)
                End If
            End If

            If u <> Right(sCodiceFiscale, 1) Then
                Return False
            Else
                Return True
            End If

        Catch Err As Exception
            log.Debug("Si è verificato un errore in ServiceOPENae::ControlloCinCF::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ControlloCinCF::" & Err.Message)
            Return False
        End Try
    End Function

    Public Function ControlloCinPIVA(ByVal sPartitaIVA As String) As Boolean
        Dim s As Integer
        Dim s1 As Integer
        Dim s2 As Integer
        Dim c As Integer
        Dim i As Integer
        Dim Chars As Integer
        'funzione scaricata dal sito http://www.icosaedro.it/cf-pi/index.html il 21/11/2003
        '1:      .s = 0
        '2. sommare ad s le cifre di posto dispari (dalla prima alla nona)
        '3. per ogni cifra di posto pari (dalla seconda alla decima),
        '   moltiplicare la cifra per due e, se risulta piu' di 9,
        '   sottrarre 9; quindi aggiungere il risultato a s;
        '4. si calcola il resto della divisione di s per 10:
        '   r=s%10 cioe' r=s-10*int(s/10); risulta un numero tra 0 e 9;
        '5. se r=0 si pone c=0, altrimenti si pone c=10-r
        '6:      .l() 'ultima cifra del cod. fisc. deve valere c.


        'restituisce TRUE se il CIN è corretto
        Try
            s1 = 0
            For i = 0 To 9
                i = i + 1
                Chars = CInt(Mid(sPartitaIVA, i, 1))
                s1 = s1 + Asc(CStr(Chars)) - Asc("0")
            Next

            For i = 1 To 9
                i = i + 1
                Chars = CInt(Mid(sPartitaIVA, i, 1))
                c = 2 * (Asc(CStr(Chars)) - Asc("0"))
                If c > 9 Then
                    c = c - 9
                    s2 = s2 + c
                Else
                    s2 = s2 + c
                End If
            Next
            s = s1 + s2
            If ((10 - s Mod 10) Mod 10 <> Asc(Mid(sPartitaIVA, 11, 1)) - Asc("0")) Then
                Return False
            Else
                Return True
            End If
        Catch Err As Exception ' Catch the error.
            log.Debug("Si è verificato un errore in ServiceOPENae::ControlloCinPIVA::" & Err.Message)
            log.Warn("Si è verificato un errore in ServiceOPENae::ControlloCinPIVA::" & Err.Message)
            Return False
        End Try
    End Function
End Class
