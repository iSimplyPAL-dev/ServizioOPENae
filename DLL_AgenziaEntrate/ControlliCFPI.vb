Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml

Namespace DLL

  Public Class ControlliCFPI

	'===============================================================================
	'WorkFlow
	'===============================================================================
	Dim m_oSession As New RIBESFrameWork.Session()
	Dim m_IDSottoAttivita As String
	Public Sub New(ByVal oSession As RIBESFrameWork.Session, ByVal IDSottoAttivita As String)
	  m_oSession = oSession
	  m_IDSottoAttivita = IDSottoAttivita
	End Sub
	'===============================================================================
	'WorkFlow
	'===============================================================================


	Public Function Calcolo_Codice_Fiscale(ByVal Cognome As String, ByVal Nome As String, ByVal DataNascita As String, ByVal sesso As String, _
	ByVal Comune As String, ByRef message As String) As String

	  Dim mesi As String
	  Dim Vocali As String
	  Dim Consonanti As String
	  Dim Numeri As String
	  Dim CDCognome(4) As String
	  Dim CDNome(4) As String

	  Dim SQLString As String

	  Dim CODICE As String

	  Dim CodiceFiscaleCognome As String
	  Dim CodiceFiscaleNome As String
	  Dim CodiceFiscaleComune As String

	  Dim PesoVocaliPari As String
	  Dim PesoconsonantiPari As String
	  Dim PesoNumeriPari As String
	  Dim PesoVocaliDispari As String
	  Dim PesoconsonantiDispari As String
	  Dim PesoNumeriDispari As String


	  Dim DataRBelfiore As SqlDataReader

	  '===============================================================================
	  'WorkFlow
	  '===============================================================================
	  Dim objDBAccess As New RIBESFrameWork.DBManager()
      Try
        objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
        '===============================================================================
        'WorkFlow
        '===============================================================================


        mesi = "ABCDEHLMPRST"
        Vocali = "AEIOU"
        Consonanti = "BCDFGHJKLMNPQRSTVWXYZ"
        Numeri = "0123456789"
        PesoVocaliPari = Chr(0) + Chr(4) + Chr(8) + Chr(14) + Chr(20)
        PesoconsonantiPari = Chr(1) + Chr(2) + Chr(3) + Chr(5) + Chr(6) + Chr(7) + Chr(9) + Chr(10) + Chr(11) + Chr(12) + Chr(13) + Chr(15) + Chr(16) + Chr(17) + Chr(18) + Chr(19) + Chr(21) + Chr(22) + Chr(23) + Chr(24) + Chr(25)
        PesoNumeriPari = Chr(0) + Chr(1) + Chr(2) + Chr(3) + Chr(4) + Chr(5) + Chr(6) + Chr(7) + Chr(8) + Chr(9)
        PesoVocaliDispari = Chr(1) + Chr(9) + Chr(19) + Chr(11) + Chr(16)
        PesoconsonantiDispari = Chr(0) + Chr(5) + Chr(7) + Chr(13) + Chr(15) + Chr(17) + Chr(21) + Chr(2) + Chr(4) + Chr(18) + Chr(20) + Chr(3) + Chr(6) + Chr(8) + Chr(12) + Chr(14) + Chr(10) + Chr(22) + Chr(25) + Chr(24) + Chr(23)
        PesoNumeriDispari = Chr(1) + Chr(0) + Chr(5) + Chr(7) + Chr(9) + Chr(13) + Chr(15) + Chr(17) + Chr(19) + Chr(21)

        '---------------------------------------
        '   Dati prelevabili dal cognome
        '---------------------------------------
        Dim l%
        Dim N%
        Dim i%
        Dim Consonante$
        Dim Vocale$

        Cognome = Trim$(Cognome)
        l% = Len(Cognome)
        N% = 0
        For i% = 1 To l%
          Consonante$ = Mid(Cognome, i%, 1)
          If InStr(Consonanti, Consonante$) > 0 Then
            N% = N% + 1
            CDCognome(N%) = Consonante$
            If N% = 3 Then Exit For
          End If
        Next
        If N% < 3 Then
          For i% = 1 To l%
            Vocale$ = Mid(Cognome, i%, 1)
            If InStr(Vocali, Vocale$) > 0 Then
              N% = N% + 1
              CDCognome(N%) = Vocale$
              If N% = 3 Then Exit For
            End If
          Next
        End If
        If N% < 3 Then
          For i% = N% + 1 To 3
            CDCognome(i%) = "X"
          Next
        End If
        CodiceFiscaleCognome = UCase(CDCognome(1)) + UCase(CDCognome(2)) + UCase(CDCognome(3))

        '----------------------------------
        '   Dati prelevabili dal Nome
        '----------------------------------
        Nome = Trim$(Nome)
        l% = Len(Nome)
        N% = 0
        For i% = 1 To l%
          Consonante$ = Mid(Nome, i%, 1)
          If InStr(Consonanti, Consonante$) > 0 Then
            N% = N% + 1
            CDNome(N%) = Consonante$
            If N% = 4 Then
              CDNome(2) = CDNome(3)
              CDNome(3) = CDNome(4)
              CDNome(4) = ""
              Exit For
            End If
          End If
        Next
        If N% < 3 Then
          For i% = 1 To l%
            Vocale$ = Mid(Nome, i%, 1)
            If InStr(Vocali, Vocale$) > 0 Then
              N% = N% + 1
              CDNome(N%) = Vocale$
              If N% = 3 Then Exit For
            End If
          Next
        End If
        If N% < 3 Then
          For i% = N% + 1 To 3
            CDNome(i%) = "X"
          Next
        End If
        CodiceFiscaleNome = UCase(CDNome(1)) + UCase(CDNome(2)) + UCase(CDNome(3))
        '-------------------------

        '-----------------------------------
        '   Dati prelevabili dalla Data
        '-----------------------------------
        Dim CodiceFiscaleData As String
        CodiceFiscaleData = Right$(Format$(Year(DataNascita), "00"), 2) + Mid(mesi, Month(DataNascita), 1)
        If sesso = "M" Then
          CodiceFiscaleData = CodiceFiscaleData + Right("00" & Left(DataNascita, 2), 2)
        Else
          CodiceFiscaleData = CodiceFiscaleData + Right("00" & (CInt(Left(DataNascita, 2)) + 40), 2)
        End If

        '------------------------------------------------
        '   Dati prelevabili dalla Tabella Comuni
        '------------------------------------------------


        SQLString = "SELECT * From COMUNI"
        SQLString = SQLString + " WHERE ((([COMUNI].[COMUNE] + [COMUNI].[PV])= '" & Replace(Comune, "'", "''") & "'));"
        '===============================================================================
        'WorkFlow
        '===============================================================================
        DataRBelfiore = objDBAccess.GetPrivateDataReader(SQLString)
        '===============================================================================
        'WorkFlow
        '===============================================================================

        If DataRBelfiore.Read = True Then
          CodiceFiscaleComune = DataRBelfiore.Item("IDENTIFICATIVO")    'objrecordset("COD_BELFIORE")
        Else
          message = "Codice Fiscale Sconosciuto"
          CodiceFiscaleComune = "X000"
        End If

        DataRBelfiore.Close()

        '------------------------------------------------------------------------------------------------------------------------------
        CODICE = CodiceFiscaleCognome + CodiceFiscaleNome + CodiceFiscaleData + CodiceFiscaleComune
        '------------------------------------------------------------------------------------------------------------------------------

        '-------------------
        '   Somma pari
        '-------------------
        Dim SommaPari%
        Dim Carattere$
        Dim Pos%

        SommaPari% = 0
        For i% = 2 To 14 Step 2
          Carattere$ = Mid(CODICE, i%, 1)
          Pos% = InStr(Vocali, Carattere$)
          If Pos% > 0 Then
            SommaPari% = SommaPari% + Asc(Mid(PesoVocaliPari, Pos%))
          Else
            Pos% = InStr(Consonanti, Carattere$)
            If Pos% > 0 Then
              SommaPari% = SommaPari% + Asc(Mid(PesoconsonantiPari, Pos%))
            Else
              Pos% = InStr(Numeri, Carattere$)
              If Pos% > 0 Then
                SommaPari% = SommaPari% + Asc(Mid(PesoNumeriPari, Pos%))
              End If
            End If
          End If
        Next

        '------------------------------------------------
        '   Somma Dispari
        '------------------------------------------------
        Dim SommaDispari%

        SommaDispari% = 0
        For i% = 1 To 15 Step 2
          Carattere$ = Mid(CODICE, i%, 1)
          Pos% = InStr(Vocali, Carattere$)
          If Pos% > 0 Then
            SommaDispari% = SommaDispari% + Asc(Mid(PesoVocaliDispari, Pos%))
          Else
            Pos% = InStr(Consonanti, Carattere$)
            If Pos% > 0 Then
              SommaDispari% = SommaDispari% + Asc(Mid(PesoconsonantiDispari, Pos%))
            Else
              Pos% = InStr(Numeri, Carattere$)
              If Pos% > 0 Then
                SommaDispari% = SommaDispari% + Asc(Mid(PesoNumeriDispari, Pos%))
              End If
            End If
          End If
        Next

        Dim RESTO%

        RESTO% = SommaPari% + SommaDispari%

        Do While RESTO% >= 26
          RESTO% = RESTO% - 26
        Loop

        Dim U$

        U$ = "."
        Pos% = InStr(1, PesoVocaliPari, Chr(RESTO%), 0)
        If Pos% > 0 Then
          U$ = Mid(Vocali, Pos%, 1)
        Else
          Pos% = InStr(1, PesoconsonantiPari, Chr(RESTO%), 0)
          If Pos% > 0 Then
            U$ = Mid(Consonanti, Pos%, 1)
          End If
        End If

        '-------------------------

        Calcolo_Codice_Fiscale = CODICE + U$
      Catch ex As Exception
        Throw New Exception("Anagrafica::Calcolo_Codice_Fiscale::" & ex.Message)
      Finally
        '********************Gestione Anagrafiche massive****************************
        objDBAccess.DisposeConnection()
        objDBAccess.Dispose()
        '********************Gestione Anagrafiche massive****************************
      End Try

    End Function

    Public Function Data_Nascita_da_CodFiscale(ByVal CodFiscale As String, ByVal Msg As Boolean, ByRef message As String) As String

      Dim anno As String
      Dim Mese As String
      Dim Giorno As String
      Dim LetteraMese As String
      Dim GiornoNascita As String

      'CALCOLO LA DATA DI NASCITA DAL CODICE FISCALE.
      'CONTROLLO SE I PRIMI DUE NUMERI CHE SI TROVANO NELLE POSIZIONI 10 E 11
      'SONO < 40, SE SI` ALLORA SO CHE E` UN MASCHIO E RICAVO IL GIORNO DI NASCITA, ALTRIMENTI E' FEMMINA
      'E DEVO SOTTRARRE 40 AL GIORNO DI NASCITA, POI ATTRAVERSO LA LETTERA CHE SI TROVA IN POSIZIONE 9
      'E LE DUE CIFRE PRECEDENTI OTTENGO IL MESE E L'ANNO DI NASCITA.


      'ripulisco le variabili
      Giorno = "" : Mese = "" : anno = ""

      'controllo che la lunghezza del codice fiscale sia corretta
      If Len(Trim(CodFiscale)) <> 16 And Len(Trim(CodFiscale)) <> 0 Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Verificarne l'esattezza e riprovare. CALCOLO DATA DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      End If
      'controllo che abbiamo passato veramente un codice fiscale
      If IsNumeric(Left(CodFiscale, 6)) = True Then
        If Msg = True Then
          message = "Non si può effettuare il calcolo dei dati di nascita su una partita IVA.CALCOLO DATA DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      End If

      'L'ANNO CORRISPONDE ALLE PRIME DUE CIFRE (CHE SI TROVANO NELLE POSIZIONI 7 E 8) E DAVANTI METTO FISSO 19
      'controllo che i caratteri che si trovano dalla posizione 7 alla 8 siano numerici
      If IsNumeric(Mid(CodFiscale, 7, 2)) = False Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire alla data di nascita.CALCOLO DATA DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      Else
        anno = "19" & Mid(CodFiscale, 7, 2)
      End If

      'IL MESE CORRISPONDE ALLA LETTERA CHE SI TROVA IN POSIZIONE 9 SAPENDO CHE I MESI
      'NEL CF HANNO LE SEGUENTI LETTERE "ABCDEHLMPRST" EFFETTUO UN SELECT CASE
      'controllo che il carattere che si trova alla posizione 10 non sia numerico
      If IsNumeric(Mid(CodFiscale, 9, 1)) = True Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire alla data di nascita.CALCOLO DATA DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      Else
        LetteraMese = Mid(CodFiscale, 9, 1)
        Select Case LetteraMese
          Case "A"   'GENNAIO
            Mese = "01"
          Case "B"   'FEBBRAIO
            Mese = "02"
          Case "C"   'MARZO
            Mese = "03"
          Case "D"   'APRILE
            Mese = "04"
          Case "E"   'MAGGIO
            Mese = "05"
          Case "H"   'GIUGNO
            Mese = "06"
          Case "L"   'LUGLIO
            Mese = "07"
          Case "M"   'AGOSTO
            Mese = "08"
          Case "P"   'SETTEMBRE
            Mese = "09"
          Case "R"   'OTTOBRE
            Mese = "10"
          Case "S"   'NOVEMBRE
            Mese = "11"
          Case "T"   'DICEMBRE
            Mese = "12"
        End Select
      End If
      'I NUMERI DI POSIZIONE 10 E 11 CORRISPONDONO AL GIORNO DI NASCITA SE IL NUMERO E`
      '> 40 E` UNA FEMMINA E QUINDI SOTTRAGGO 40 ALTRIMENTI NON SOTTRAGGO NIENTE
      'controllo che i caratteri che si trovano dalla posizione 10 alla 11 non siano numerici
      If IsNumeric(Mid(CodFiscale, 10, 2)) = False Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire alla data di nascita.CALCOLO DATA DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      Else
        GiornoNascita = Mid(CodFiscale, 10, 2)
        If CInt(GiornoNascita) > 40 Then
          Giorno = CInt(GiornoNascita) - 40
        Else
          Giorno = GiornoNascita
        End If
        Giorno = Right("00" & Giorno, 2)
      End If
      'RICOMPONGO LA DATA DI NASCITA
      If Trim(Giorno) <> "" And Trim(Mese) <> "" And Trim(anno) <> "" Then
        'controllo che la data ottenuta abbia senso
        '///////////////////////////////////////////////////////
        If controlla_data(Giorno & "/" & Mese & "/" & anno) = True Then
          Data_Nascita_da_CodFiscale = Giorno & "/" & Mese & "/" & anno
        Else
          If Msg = True Then
            message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire alla data di nascita.CALCOLO DATA DI NASCITA DA CODICE FISCALE"
          End If
        End If
        '//////////////////////////////////////////
      End If

    End Function

    Public Function Luogo_Nascita_da_CodFiscale(ByVal CodFiscale As String, ByVal Msg As Boolean, ByRef message As String, _
    ByRef Identificativo As String, ByRef Provincia As String) As String
      Dim CodCom As String

      'CALCOLO IL LUOGO DI NASCITA.
      'SAPENDO CHE LA LETTERA CHE SI TROVA IN POSIZIONE 12 E LE TRE CIFRE SUCCESSIVE CORRISPONDO AL
      'CODICE BELFIORE DEL COMUNE DI NASCITA, OTTENGO IL LUOGO

      'controllo che la lunghezza del codice fiscale sia corretta
      If Len(Trim(CodFiscale)) <> 16 And Len(Trim(CodFiscale)) <> 0 Then
        If Msg = True Then
          MsgBox("Il codice fiscale risulta errato." & Chr(10) & "Verificarne l'esattezza e riprovare.", 16, "CALCOLO LUOGO DI NASCITA DA CODICE FISCALE")
        End If
        Exit Function
      End If
      'controllo che abbiamo passato veramente un codice fiscale
      If IsNumeric(Mid(CodFiscale, 1, 6)) = True Then
        If Msg = True Then
          message = "Non si può effettuare il calcolo dei dati di nascita su una partita IVA.CALCOLO LUOGO DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      End If

      'PER OTTENERE IL COMUNE DI NASCITA UTILIZZO IL CODICE CHE SI TROVA DALLA POSIZIONE 12 ALLA POSIZIONE 15
      'controllo che il carattere che si trova alla 12 posizione non sia numerico
      If IsNumeric(Mid(CodFiscale, 12, 1)) = True Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire al luogo di nascita.CALCOLO LUOGO DI NASCITA DA CODICE FISCALE"
        End If
        Exit Function
      Else
        'controllo che ii caratteri che si trovano dalla posizione 13 alla 15 posizione siano numerici
        If IsNumeric(Mid(CodFiscale, 13, 3)) = False Then
          If Msg = True Then
            message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire al luogo di nascita.CALCOLO LUOGO DI NASCITA DA CODICE FISCALE"
          End If
          Exit Function
        Else
          Dim SQL As String
          Dim CmdComune As New SqlCommand
          Dim DrComune As SqlDataReader
          '===============================================================================
          'WorkFlow
          '===============================================================================
          Dim objDBAccess As New RIBESFrameWork.DBManager
          Try
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================

            CodCom = Mid(CodFiscale, 12, 4)
            Identificativo = CodCom
            SQL = "SELECT COMUNI.COMUNE, COMUNI.PV "
            SQL = SQL + " FROM COMUNI WHERE COMUNI.IDENTIFICATIVO='" & CodCom & "'"


            DrComune = objDBAccess.GetPrivateDataReader(SQL)

            If DrComune.Read = True Then
              Luogo_Nascita_da_CodFiscale = Utility.GetParametro(DrComune.Item("COMUNE"))
              Provincia = Utility.GetParametro(DrComune.Item("PV"))
            Else
              If Msg = True Then
                message = "Codice Belfiore inesistente." & Chr(10) & "Impossibile risalire al Comune di Nascita.CALCOLO LUOGO DI NASCITA DA CODICE FISCALE"
              End If
            End If

            DrComune.Close()

          Catch ex As Exception
            Throw New Exception("Anagrafica::Luogo_Nascita_da_CodFiscale::" & ex.Message)
          Finally
            '********************Gestione Anagrafiche massive****************************
            objDBAccess.DisposeConnection()
            objDBAccess.Dispose()
            '********************Gestione Anagrafiche massive****************************
          End Try
        End If
      End If

    End Function

    Public Function Sesso_da_CodFiscale(ByVal CodFiscale As String, ByVal Msg As Boolean, ByRef message As String) As String
      Dim Giorno As String

      'CALCOLO IL SESSO DAL CODICE FISCALE.
      'CONTROLLO SE I PRIMI DUE NUMERI CHE SI TROVANO NELLE POSIZIONI 10 E 11
      'SONO < 40, SE SI` ALLORA SO CHE E` UN MASCHIO ALTRIMENTI E' UNA FEMMINA

      'controllo che la lunghezza del codice fiscale sia corretta
      If Len(Trim(CodFiscale)) <> 16 And Len(Trim(CodFiscale)) <> 0 Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Verificarne l'esattezza e riprovare.CALCOLO SESSO DA CODICE FISCALE"
        End If
        Exit Function
      End If
      'controllo che abbiamo passato veramente un codice fiscale
      If IsNumeric(Mid(CodFiscale, 1, 6)) = True Then
        If Msg = True Then
          message = "Non si può effettuare il calcolo dei dati di nascita su una partita IVA.CALCOLO SESSO DA CODICE FISCALE"
        End If
        Exit Function
      End If

      'I NUMERI DI POSIZIONE 10 E 11 CORRISPONDONO AL GIORNO DI NASCITA SE IL NUMERO E`
      '> 40 E` UNA FEMMINA E A SESSO ASSEGNO "F" ALTRIMENTI E` UN MASCHIO A SESSO ASSEGNO "M"
      'controllo che i caratteri che si trovano dalla posizione 10 alla 11 non siano numerici
      If IsNumeric(Mid(CodFiscale, 10, 2)) = False Then
        If Msg = True Then
          message = "Il codice fiscale risulta errato." & Chr(10) & "Impossibile risalire alla data di nascita.CALCOLO SESSO DA CODICE FISCALE"
        End If
        Exit Function
      Else
        Giorno = Mid(CodFiscale, 10, 2)
        If Giorno > 40 Then
          Sesso_da_CodFiscale = "F"
        Else
          Sesso_da_CodFiscale = "M"
        End If
      End If

    End Function

    Public Function ControlloCinCF(ByVal CodiceFiscale As String) As Boolean
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
      ControlloCinCF = True

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

        CFControllo = Left(CodiceFiscale, 15)

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

        If u <> Right(CodiceFiscale, 1) Then
          ControlloCinCF = False
        End If

      Catch ex As Exception ' Catch the error.
        ControlloCinCF = False
      End Try
    End Function

    Public Function ControlloCinPI(ByVal PIVA As String) As Boolean
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
      ControlloCinPI = True
      Try
        s1 = 0
        For i = 0 To 9
          i = i + 1
          Chars = CInt(Mid(PIVA, i, 1))
          s1 = s1 + Asc(CStr(Chars)) - Asc("0")
        Next

        For i = 1 To 9
          i = i + 1
          Chars = CInt(Mid(PIVA, i, 1))
          c = 2 * (Asc(CStr(Chars)) - Asc("0"))
          If c > 9 Then
            c = c - 9
            s2 = s2 + c
          Else
            s2 = s2 + c
          End If
        Next
        s = s1 + s2
        If ((10 - s Mod 10) Mod 10 <> Asc(Mid(PIVA, 11, 1)) - Asc("0")) Then
          ControlloCinPI = False
        Else
          ControlloCinPI = True
        End If
      Catch ex As Exception ' Catch the error.
        ControlloCinPI = False
      End Try
    End Function

    Function controlla_data(ByRef datacontrollo As String) As Boolean
      'SUB CHE CONTROLLA SE E' STATA INSERITA UNA DATA CORRETTA
      Dim controllo_data As Object
      controlla_data = True
      Try

        Dim Mese, Giorno, Anno As Integer
        Dim bisestile As Integer



        Giorno = CInt(Microsoft.VisualBasic.Left(datacontrollo, 2))
        Mese = CInt(Mid(datacontrollo, 4, 2))
        Anno = CInt(Mid(datacontrollo, 7, 4))
        If Len(Anno) = 4 Then
          bisestile = CInt(Anno) Mod 4
        Else
          Throw New Exception("Anno Errato")
        End If

        'controllo del giorno
        If Mese = 2 And bisestile = 0 Then    'controllo giorni di feb. quando Anno(bisestile)
          If Giorno < 1 Or Giorno > 29 Then
            Throw New Exception("Giorno Errato")
          End If
        ElseIf Mese = 2 And bisestile <> 0 Then  'controllo giorni di feb.quando anno non bisestile
          If Giorno < 1 Or Giorno > 28 Then
            Throw New Exception("Giorno Errato")
          End If
        ElseIf Mese = 11 Or Mese = 4 Or Mese = 6 Or Mese = 9 Then
          'controllo giorni se il mese ne deve avere 30
          If Giorno < 1 Or Giorno > 30 Then
            Throw New Exception("Giorno Errato")
          End If
        ElseIf Mese <> 11 And Mese <> 4 And Mese <> 6 And Mese <> 9 Then
          'altri mesi
          If Giorno < 1 Or Giorno > 31 Then
            Throw New Exception("Giorno Errato")
          End If
        End If

        'controllo mese
        If Mese < 1 Or Mese > 12 Then
          Throw New Exception("Mese Errato")
        End If

        Exit Function

      Catch ex As Exception
        controlla_data = False
        Throw ex
      End Try

    End Function


    Public Function ControlloChrFuoriPostoCF(ByVal CF As String) As Boolean
      Dim IDChar As Integer
      Dim ChrControllo As String

      ControlloChrFuoriPostoCF = True
      IDChar = 1
      For IDChar = 1 To 16
        ChrControllo = Mid(CF, IDChar, 1)
        Select Case IDChar
          Case 1, 2, 3, 4, 5, 6, 9, 12, 16
            If (Asc(ChrControllo) < 65 Or Asc(ChrControllo) > 90) And (Asc(ChrControllo) < 97 Or Asc(ChrControllo) > 122) Then
              ControlloChrFuoriPostoCF = False
              Exit Function
            End If
          Case 7, 8, 10, 11, 13, 14, 15
            If Not IsNumeric(ChrControllo) Then
              ControlloChrFuoriPostoCF = False
              Exit Function
            End If
        End Select
      Next IDChar
    End Function


    Public Function ControlloChrFuoriPostoPI(ByVal PI As String) As Boolean
      Dim IDChar As Integer
      Dim ChrControllo As String

      ControlloChrFuoriPostoPI = True
      IDChar = 1
      For IDChar = 1 To 11
        ChrControllo = Mid(PI, IDChar, 1)
        Select Case IDChar
          Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
            If Not IsNumeric(ChrControllo) Then
              ControlloChrFuoriPostoPI = False
            End If
        End Select
      Next IDChar
    End Function
  End Class
End Namespace