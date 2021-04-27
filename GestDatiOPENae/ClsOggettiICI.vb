Public Class ImportTestaICI
    Private _nCodFlussoEnte As Integer = -1
    Private _nCodFlussoPagamenti As Integer = -1
    Private _nTotPagamentiAcq As Integer = 0
    Private _nTotSanzioni As Integer = 0
    Private _sCodISTAT As String = ""
    Private _sDataCreazione As String = ""
    Private _sNomeTracciato As String = ""
    Private _sNomeFlusso As String = ""
    Private _sDataCreazioneFlusso As String = ""
    Private _sDivisa As String = ""
    Private _sDataInizio As String = ""
    Private _sDataFine As String = ""
    Private _sAnnoFlusso As String = ""
    Private _nTotImpPos As Double = 0
    Private _nTotImpNeg As Double = 0
    Private _nTotaleRiversato As Double = 0
    Private _nTotaleImpSanzioni As Double = 0

    Public Property nCodFlussoEnte() As Integer
        Get
            Return _nCodFlussoEnte
        End Get

        Set(ByVal Value As Integer)
            _nCodFlussoEnte = Value
        End Set
    End Property
    Public Property nCodFlussoPagamenti() As Integer
        Get
            Return _nCodFlussoPagamenti
        End Get

        Set(ByVal Value As Integer)
            _nCodFlussoPagamenti = Value
        End Set
    End Property
    Public Property nTotPagamentiAcq() As Integer
        Get
            Return _nTotPagamentiAcq
        End Get

        Set(ByVal Value As Integer)
            _nTotPagamentiAcq = Value
        End Set
    End Property
    Public Property nTotSanzioni() As Integer
        Get
            Return _nTotSanzioni
        End Get

        Set(ByVal Value As Integer)
            _nTotSanzioni = Value
        End Set
    End Property
    Public Property sCodISTAT() As String
        Get
            Return _sCodISTAT
        End Get

        Set(ByVal Value As String)
            _sCodISTAT = Value
        End Set
    End Property
    Public Property sDataCreazione() As String
        Get
            Return _sDataCreazione
        End Get

        Set(ByVal Value As String)
            _sDataCreazione = Value
        End Set
    End Property
    Public Property sNomeTracciato() As String
        Get
            Return _sNomeTracciato
        End Get

        Set(ByVal Value As String)
            _sNomeTracciato = Value
        End Set
    End Property
    Public Property sNomeFlusso() As String
        Get
            Return _sNomeFlusso
        End Get

        Set(ByVal Value As String)
            _sNomeFlusso = Value
        End Set
    End Property
    Public Property sDataCreazioneFlusso() As String
        Get
            Return _sDataCreazioneFlusso
        End Get

        Set(ByVal Value As String)
            _sDataCreazioneFlusso = Value
        End Set
    End Property
    Public Property sDivisa() As String
        Get
            Return _sDivisa
        End Get

        Set(ByVal Value As String)
            _sDivisa = Value
        End Set
    End Property
    Public Property sDataInizio() As String
        Get
            Return _sDataInizio
        End Get

        Set(ByVal Value As String)
            _sDataInizio = Value
        End Set
    End Property
    Public Property sDataFine() As String
        Get
            Return _sDataFine
        End Get

        Set(ByVal Value As String)
            _sDataFine = Value
        End Set
    End Property
    Public Property sAnnoFlusso() As String
        Get
            Return _sAnnoFlusso
        End Get

        Set(ByVal Value As String)
            _sAnnoFlusso = Value
        End Set
    End Property
    Public Property nTotImpPos() As Double
        Get
            Return _nTotImpPos
        End Get

        Set(ByVal Value As Double)
            _nTotImpPos = Value
        End Set
    End Property
    Public Property nTotImpNeg() As Double
        Get
            Return _nTotImpNeg
        End Get

        Set(ByVal Value As Double)
            _nTotImpNeg = Value
        End Set
    End Property
    Public Property nTotaleRiversato() As Double
        Get
            Return _nTotaleRiversato
        End Get

        Set(ByVal Value As Double)
            _nTotaleRiversato = Value
        End Set
    End Property
    Public Property nTotaleImpSanzioni() As Double
        Get
            Return _nTotaleImpSanzioni
        End Get

        Set(ByVal Value As Double)
            _nTotaleImpSanzioni = Value
        End Set
    End Property
End Class

Public Class ImportCodaICI
    Private _nCodFlussoPagamenti As Integer = -1
    Private _nTotAnagrafiche As Integer = 0
    Private _nTotDisposizioni As Integer = 0
    Private _nTotVersamenti As Integer = 0
    Private _nTotViolazioni As Integer = 0
    Private _nTotImpVersamenti As Double = 0
    Private _nTotImpTerAgr As Double = 0
    Private _nTotImpAreeFab As Double = 0
    Private _nTotImpAbiPrin As Double = 0
    Private _nTotImpAltriFab As Double = 0
    Private _nTotImpDetrazione As Double = 0
    Private _nTotImpViolazioni As Double = 0

    Public Property nCodFlussoPagamenti() As Integer
        Get
            Return _nCodFlussoPagamenti
        End Get

        Set(ByVal Value As Integer)
            _nCodFlussoPagamenti = Value
        End Set
    End Property
    Public Property nTotAnagrafiche() As Integer
        Get
            Return _nTotAnagrafiche
        End Get

        Set(ByVal Value As Integer)
            _nTotAnagrafiche = Value
        End Set
    End Property
    Public Property nTotDisposizioni() As Integer
        Get
            Return _nTotDisposizioni
        End Get

        Set(ByVal Value As Integer)
            _nTotDisposizioni = Value
        End Set
    End Property
    Public Property nTotVersamenti() As Integer
        Get
            Return _nTotVersamenti
        End Get

        Set(ByVal Value As Integer)
            _nTotVersamenti = Value
        End Set
    End Property
    Public Property nTotViolazioni() As Integer
        Get
            Return _nTotViolazioni
        End Get

        Set(ByVal Value As Integer)
            _nTotViolazioni = Value
        End Set
    End Property
    Public Property nTotImpVersamenti() As Double
        Get
            Return _nTotImpVersamenti
        End Get

        Set(ByVal Value As Double)
            _nTotImpVersamenti = Value
        End Set
    End Property
    Public Property nTotImpTerAgr() As Double
        Get
            Return _nTotImpTerAgr
        End Get

        Set(ByVal Value As Double)
            _nTotImpTerAgr = Value
        End Set
    End Property
    Public Property nTotImpAreeFab() As Double
        Get
            Return _nTotImpAreeFab
        End Get

        Set(ByVal Value As Double)
            _nTotImpAreeFab = Value
        End Set
    End Property
    Public Property nTotImpAbiPrin() As Double
        Get
            Return _nTotImpAbiPrin
        End Get

        Set(ByVal Value As Double)
            _nTotImpAbiPrin = Value
        End Set
    End Property
    Public Property nTotImpAltriFab() As Double
        Get
            Return _nTotImpAltriFab
        End Get

        Set(ByVal Value As Double)
            _nTotImpAltriFab = Value
        End Set
    End Property
    Public Property nTotImpDetrazione() As Double
        Get
            Return _nTotImpDetrazione
        End Get

        Set(ByVal Value As Double)
            _nTotImpDetrazione = Value
        End Set
    End Property
    Public Property nTotImpViolazioni() As Double
        Get
            Return _nTotImpViolazioni
        End Get

        Set(ByVal Value As Double)
            _nTotImpViolazioni = Value
        End Set
    End Property
End Class

Public Class ImportAnagICI
    Private _nIdContribuente As Integer = -1
    Private _sCodISTAT As String = ""
    Private _sCFPIVA As String = ""
    Private _sCognome As String = ""
    Private _sNome As String = ""
    Private _sSesso As String = ""
    Private _sDataNascita As String = ""
    Private _sComuneNascita As String = ""
    Private _sPVNascita As String = ""
    Private _sNazionalita As String = ""
    Private _sViaRes As String = ""
    Private _sFrazioneRes As String = ""
    Private _sCivicoRes As String = ""
    Private _sCAPRes As String = ""
    Private _sCittaRes As String = ""
    Private _sPVRes As String = ""
    Private _sNominativoInvio As String = ""
    Private _sViaInvio As String = ""
    Private _sCivicoInvio As String = ""
    Private _sCAPInvio As String = ""
    Private _sCittaInvio As String = ""
    Private _sPVInvio As String = ""

    Public Property nIdContribuente() As Integer
        Get
            Return _nIdContribuente
        End Get

        Set(ByVal Value As Integer)
            _nIdContribuente = Value
        End Set
    End Property
    Public Property sCodISTAT() As String
        Get
            Return _sCodISTAT
        End Get

        Set(ByVal Value As String)
            _sCodISTAT = Value
        End Set
    End Property
    Public Property sCFPIVA() As String
        Get
            Return _sCFPIVA
        End Get

        Set(ByVal Value As String)
            _sCFPIVA = Value
        End Set
    End Property
    Public Property sCognome() As String
        Get
            Return _sCognome
        End Get

        Set(ByVal Value As String)
            _sCognome = Value
        End Set
    End Property
    Public Property sNome() As String
        Get
            Return _sNome
        End Get

        Set(ByVal Value As String)
            _sNome = Value
        End Set
    End Property
    Public Property sSesso() As String
        Get
            Return _sSesso
        End Get

        Set(ByVal Value As String)
            _sSesso = Value
        End Set
    End Property
    Public Property sDataNascita() As String
        Get
            Return _sDataNascita
        End Get

        Set(ByVal Value As String)
            _sDataNascita = Value
        End Set
    End Property
    Public Property sComuneNascita() As String
        Get
            Return _sComuneNascita
        End Get

        Set(ByVal Value As String)
            _sComuneNascita = Value
        End Set
    End Property
    Public Property sPVNascita() As String
        Get
            Return _sPVNascita
        End Get

        Set(ByVal Value As String)
            _sPVNascita = Value
        End Set
    End Property
    Public Property sNazionalita() As String
        Get
            Return _sNazionalita
        End Get

        Set(ByVal Value As String)
            _sNazionalita = Value
        End Set
    End Property
    Public Property sViaRes() As String
        Get
            Return _sViaRes
        End Get

        Set(ByVal Value As String)
            _sViaRes = Value
        End Set
    End Property
    Public Property sFrazioneRes() As String
        Get
            Return _sFrazioneRes
        End Get

        Set(ByVal Value As String)
            _sFrazioneRes = Value
        End Set
    End Property
    Public Property sCivicoRes() As String
        Get
            Return _sCivicoRes
        End Get

        Set(ByVal Value As String)
            _sCivicoRes = Value
        End Set
    End Property
    Public Property sCAPRes() As String
        Get
            Return _sCAPRes
        End Get

        Set(ByVal Value As String)
            _sCAPRes = Value
        End Set
    End Property
    Public Property sCittaRes() As String
        Get
            Return _sCittaRes
        End Get

        Set(ByVal Value As String)
            _sCittaRes = Value
        End Set
    End Property
    Public Property sPVRes() As String
        Get
            Return _sPVRes
        End Get

        Set(ByVal Value As String)
            _sPVRes = Value
        End Set
    End Property
    Public Property sNominativoInvio() As String
        Get
            Return _sNominativoInvio
        End Get

        Set(ByVal Value As String)
            _sNominativoInvio = Value
        End Set
    End Property
    Public Property sViaInvio() As String
        Get
            Return _sViaInvio
        End Get

        Set(ByVal Value As String)
            _sViaInvio = Value
        End Set
    End Property
    Public Property sCivicoInvio() As String
        Get
            Return _sCivicoInvio
        End Get

        Set(ByVal Value As String)
            _sCivicoInvio = Value
        End Set
    End Property
    Public Property sCAPInvio() As String
        Get
            Return _sCAPInvio
        End Get

        Set(ByVal Value As String)
            _sCAPInvio = Value
        End Set
    End Property
    Public Property sCittaInvio() As String
        Get
            Return _sCittaInvio
        End Get

        Set(ByVal Value As String)
            _sCittaInvio = Value
        End Set
    End Property
    Public Property sPVInvio() As String
        Get
            Return _sPVInvio
        End Get

        Set(ByVal Value As String)
            _sPVInvio = Value
        End Set
    End Property
End Class

Public Class ImportDisposizioneICI
    Private _nIdVersamento As Integer = -1
    Private _nCodFlusso As Integer = -1
    Private _nIdContribuente As Integer = -1
    Private _nIdContribuenteSimile As Integer = -1
    Private _nNumFab As Integer = 0
    Private _nNumQuietanza As Integer = 0
    Private _nFlagTrattato As Integer = 0
    Private _nCodFlussoAP As Integer = 0
    Private _nProgrPagAP As Integer = 0
    Private _nProgRendicontaz As Integer = 0
    Private _sCodISTAT As String = ""
    Private _sCFPIVA As String = ""
    Private _sNominativo As String = ""
    Private _sDataAccredito As String = ""
    Private _sDataPagamento As String = ""
    Private _sFlagAS As String = "0"
    Private _sAnno As String = ""
    Private _sAnnoRif As String = ""
    Private _sIndirizzoRes As String = ""
    Private _sCapRes As String = ""
    Private _sCittaRes As String = ""
    Private _sDataSanzione As String = ""
    Private _sNumSanzione As String = ""
    Private _sNumMovimento As String = ""
    Private _sSpazioLibero As String = ""
    Private _sDataFlussoRend As String = ""
    Private _sDivisa As String = ""
    Private _sNomeImmagine As String = ""
    Private _sProvenienza As String = ""
    Private _sViewImmagine As String = ""
    Private _sCodiceComunico As String = ""
    Private _sCodTipoPagamento As String = ""
    Private _sFlagRavvOperoso As String = ""
    Private _sTipoBollettinoViolazioni As String = ""
    Private _sBollettinoEXRurale As String = ""
    Private _nImpVersamento As Double = 0
    Private _nImpTerAgr As Double = 0
    Private _nImpAreeFab As Double = 0
    Private _nImpAltriFab As Double = 0
    Private _nImpAbiPrinc As Double = 0
    Private _nImpDetrazione As Double = 0
    Private _nImpVersato As Double = 0

    Public Property nIdVersamento() As Integer
        Get
            Return _nIdVersamento
        End Get

        Set(ByVal Value As Integer)
            _nIdVersamento = Value
        End Set
    End Property
    Public Property nCodFlusso() As Integer
        Get
            Return _nCodFlusso
        End Get

        Set(ByVal Value As Integer)
            _nCodFlusso = Value
        End Set
    End Property
    Public Property nIdContribuente() As Integer
        Get
            Return _nIdContribuente
        End Get

        Set(ByVal Value As Integer)
            _nIdContribuente = Value
        End Set
    End Property
    Public Property nIdContribuenteSimile() As Integer
        Get
            Return _nIdContribuenteSimile
        End Get

        Set(ByVal Value As Integer)
            _nIdContribuenteSimile = Value
        End Set
    End Property
    Public Property nNumFab() As Integer
        Get
            Return _nNumFab
        End Get

        Set(ByVal Value As Integer)
            _nNumFab = Value
        End Set
    End Property
    Public Property nNumQuietanza() As Integer
        Get
            Return _nNumQuietanza
        End Get

        Set(ByVal Value As Integer)
            _nNumQuietanza = Value
        End Set
    End Property
    Public Property nFlagTrattato() As Integer
        Get
            Return _nFlagTrattato
        End Get

        Set(ByVal Value As Integer)
            _nFlagTrattato = Value
        End Set
    End Property
    Public Property nCodFlussoAP() As Integer
        Get
            Return _nCodFlussoAP
        End Get

        Set(ByVal Value As Integer)
            _nCodFlussoAP = Value
        End Set
    End Property
    Public Property nProgrPagAP() As Integer
        Get
            Return _nProgrPagAP
        End Get

        Set(ByVal Value As Integer)
            _nProgrPagAP = Value
        End Set
    End Property
    Public Property nProgRendicontaz() As Integer
        Get
            Return _nProgRendicontaz
        End Get

        Set(ByVal Value As Integer)
            _nProgRendicontaz = Value
        End Set
    End Property
    Public Property sCodISTAT() As String
        Get
            Return _sCodISTAT
        End Get

        Set(ByVal Value As String)
            _sCodISTAT = Value
        End Set
    End Property
    Public Property sCFPIVA() As String
        Get
            Return _sCFPIVA
        End Get

        Set(ByVal Value As String)
            _sCFPIVA = Value
        End Set
    End Property
    Public Property sNominativo() As String
        Get
            Return _sNominativo
        End Get

        Set(ByVal Value As String)
            _sNominativo = Value
        End Set
    End Property
    Public Property sDataAccredito() As String
        Get
            Return _sDataAccredito
        End Get

        Set(ByVal Value As String)
            _sDataAccredito = Value
        End Set
    End Property
    Public Property sDataPagamento() As String
        Get
            Return _sDataPagamento
        End Get

        Set(ByVal Value As String)
            _sDataPagamento = Value
        End Set
    End Property
    Public Property sFlagAS() As String
        Get
            Return _sFlagAS
        End Get

        Set(ByVal Value As String)
            _sFlagAS = Value
        End Set
    End Property
    Public Property sAnno() As String
        Get
            Return _sAnno
        End Get

        Set(ByVal Value As String)
            _sAnno = Value
        End Set
    End Property
    Public Property sAnnoRif() As String
        Get
            Return _sAnnoRif
        End Get

        Set(ByVal Value As String)
            _sAnnoRif = Value
        End Set
    End Property
    Public Property sIndirizzoRes() As String
        Get
            Return _sIndirizzoRes
        End Get

        Set(ByVal Value As String)
            _sIndirizzoRes = Value
        End Set
    End Property
    Public Property sCapRes() As String
        Get
            Return _sCapRes
        End Get

        Set(ByVal Value As String)
            _sCapRes = Value
        End Set
    End Property
    Public Property sCittaRes() As String
        Get
            Return _sCittaRes
        End Get

        Set(ByVal Value As String)
            _sCittaRes = Value
        End Set
    End Property
    Public Property sDataSanzione() As String
        Get
            Return _sDataSanzione
        End Get

        Set(ByVal Value As String)
            _sDataSanzione = Value
        End Set
    End Property
    Public Property sNumSanzione() As String
        Get
            Return _sNumSanzione
        End Get

        Set(ByVal Value As String)
            _sNumSanzione = Value
        End Set
    End Property
    Public Property sNumMovimento() As String
        Get
            Return _sNumMovimento
        End Get

        Set(ByVal Value As String)
            _sNumMovimento = Value
        End Set
    End Property
    Public Property sSpazioLibero() As String
        Get
            Return _sSpazioLibero
        End Get

        Set(ByVal Value As String)
            _sSpazioLibero = Value
        End Set
    End Property
    Public Property sDataFlussoRend() As String
        Get
            Return _sDataFlussoRend
        End Get

        Set(ByVal Value As String)
            _sDataFlussoRend = Value
        End Set
    End Property
    Public Property sDivisa() As String
        Get
            Return _sDivisa
        End Get

        Set(ByVal Value As String)
            _sDivisa = Value
        End Set
    End Property
    Public Property sNomeImmagine() As String
        Get
            Return _sNomeImmagine
        End Get

        Set(ByVal Value As String)
            _sNomeImmagine = Value
        End Set
    End Property
    Public Property sProvenienza() As String
        Get
            Return _sProvenienza
        End Get

        Set(ByVal Value As String)
            _sProvenienza = Value
        End Set
    End Property
    Public Property sViewImmagine() As String
        Get
            Return _sViewImmagine
        End Get

        Set(ByVal Value As String)
            _sViewImmagine = Value
        End Set
    End Property
    Public Property sCodiceComunico() As String
        Get
            Return _sCodiceComunico
        End Get

        Set(ByVal Value As String)
            _sCodiceComunico = Value
        End Set
    End Property
    Public Property sCodTipoPagamento() As String
        Get
            Return _sCodTipoPagamento
        End Get

        Set(ByVal Value As String)
            _sCodTipoPagamento = Value
        End Set
    End Property
    Public Property sFlagRavvOperoso() As String
        Get
            Return _sFlagRavvOperoso
        End Get

        Set(ByVal Value As String)
            _sFlagRavvOperoso = Value
        End Set
    End Property
    Public Property sTipoBollettinoViolazioni() As String
        Get
            Return _sTipoBollettinoViolazioni
        End Get

        Set(ByVal Value As String)
            _sTipoBollettinoViolazioni = Value
        End Set
    End Property
    Public Property sBollettinoEXRurale() As String
        Get
            Return _sBollettinoEXRurale
        End Get

        Set(ByVal Value As String)
            _sBollettinoEXRurale = Value
        End Set
    End Property
    Public Property nImpVersamento() As Double
        Get
            Return _nImpVersamento
        End Get

        Set(ByVal Value As Double)
            _nImpVersamento = Value
        End Set
    End Property
    Public Property nImpTerAgr() As Double
        Get
            Return _nImpTerAgr
        End Get

        Set(ByVal Value As Double)
            _nImpTerAgr = Value
        End Set
    End Property
    Public Property nImpAreeFab() As Double
        Get
            Return _nImpAreeFab
        End Get

        Set(ByVal Value As Double)
            _nImpAreeFab = Value
        End Set
    End Property
    Public Property nImpAltriFab() As Double
        Get
            Return _nImpAltriFab
        End Get

        Set(ByVal Value As Double)
            _nImpAltriFab = Value
        End Set
    End Property
    Public Property nImpAbiPrinc() As Double
        Get
            Return _nImpAbiPrinc
        End Get

        Set(ByVal Value As Double)
            _nImpAbiPrinc = Value
        End Set
    End Property
    Public Property nImpDetrazione() As Double
        Get
            Return _nImpDetrazione
        End Get

        Set(ByVal Value As Double)
            _nImpDetrazione = Value
        End Set
    End Property
    Public Property nImpVersato() As Double
        Get
            Return _nImpVersato
        End Get

        Set(ByVal Value As Double)
            _nImpVersato = Value
        End Set
    End Property
End Class

Public Class ExportTestaICI
    Dim _nProgressivoInvio As Integer = 1
    Dim _nNumSupporti As Integer = 0
    Dim _nNumOrdineSupporto As Integer = 0
    Dim _sCodConcessione As String = "1"
    Dim _sPeriodoRifRiscossioni As String = ""
    Dim _sDataScadenza As String = ""
    Dim _sFiller As String = "0"

    Public Property nProgressivoInvio() As Integer
        Get
            Return _nProgressivoInvio
        End Get

        Set(ByVal Value As Integer)
            _nProgressivoInvio = Value
        End Set
    End Property
    Public Property nNumSupporti() As Integer
        Get
            Return _nNumSupporti
        End Get

        Set(ByVal Value As Integer)
            _nNumSupporti = Value
        End Set
    End Property
    Public Property nNumOrdineSupporto() As Integer
        Get
            Return _nNumOrdineSupporto
        End Get

        Set(ByVal Value As Integer)
            _nNumOrdineSupporto = Value
        End Set
    End Property
    Public Property sCodConcessione() As String
        Get
            Return _sCodConcessione
        End Get

        Set(ByVal Value As String)
            _sCodConcessione = Value
        End Set
    End Property
    Public Property sPeriodoRifRiscossioni() As String
        Get
            Return _sPeriodoRifRiscossioni
        End Get

        Set(ByVal Value As String)
            _sPeriodoRifRiscossioni = Value
        End Set
    End Property
    Public Property sDataScadenza() As String
        Get
            Return _sDataScadenza
        End Get

        Set(ByVal Value As String)
            _sDataScadenza = Value
        End Set
    End Property
    Public Property sFiller() As String
        Get
            Return _sFiller
        End Get

        Set(ByVal Value As String)
            _sFiller = Value
        End Set
    End Property
End Class

Public Class ExportCodaICI
    Dim _nProgressivoInvio As Integer = 1
    Dim _nTotRc1 As Integer = 0
    Dim _nTotRc3 As Integer = 0
    Dim _nTotRc45 As Integer = 0
    Dim _nTotRc6 As Integer = 0
    Dim _sCodConcessione As String = "1"
    Dim _sPeriodoRifRiscossioni As String = ""
    Dim _sDataScadenza As String = ""
    Dim _sFiller As String = "0"

    Public Property nProgressivoInvio() As Integer
        Get
            Return _nProgressivoInvio
        End Get

        Set(ByVal Value As Integer)
            _nProgressivoInvio = Value
        End Set
    End Property
    Public Property nTotRc1() As Integer
        Get
            Return _nTotRc1
        End Get

        Set(ByVal Value As Integer)
            _nTotRc1 = Value
        End Set
    End Property
    Public Property nTotRc3() As Integer
        Get
            Return _nTotRc3
        End Get

        Set(ByVal Value As Integer)
            _nTotRc3 = Value
        End Set
    End Property
    Public Property nTotRc45() As Integer
        Get
            Return _nTotRc45
        End Get

        Set(ByVal Value As Integer)
            _nTotRc45 = Value
        End Set
    End Property
    Public Property nTotRc6() As Integer
        Get
            Return _nTotRc6
        End Get

        Set(ByVal Value As Integer)
            _nTotRc6 = Value
        End Set
    End Property
    Public Property sCodConcessione() As String
        Get
            Return _sCodConcessione
        End Get

        Set(ByVal Value As String)
            _sCodConcessione = Value
        End Set
    End Property
    Public Property sPeriodoRifRiscossioni() As String
        Get
            Return _sPeriodoRifRiscossioni
        End Get

        Set(ByVal Value As String)
            _sPeriodoRifRiscossioni = Value
        End Set
    End Property
    Public Property sDataScadenza() As String
        Get
            Return _sDataScadenza
        End Get

        Set(ByVal Value As String)
            _sDataScadenza = Value
        End Set
    End Property
    Public Property sFiller() As String
        Get
            Return _sFiller
        End Get

        Set(ByVal Value As String)
            _sFiller = Value
        End Set
    End Property
End Class

Public Class ExportRiversamentoICI
    Dim _sCodConcessione As String = "1"
    Dim _sCodBelfiore As String = ""
    Dim _sDataRiversamento As String = ""
    Dim _sTipoRiscossioni As String = "M"
    Dim _sFiller As String = "0"
    Dim _nNumQuietanza As Integer = 0
    Dim _nProgressivoRc As Integer = 0
    Dim _nCodTesoreria As Integer = 0
    Dim _nCommissione As Integer = 0
    Dim _nTotNumRiscossioni As Integer = 0
    Dim _nTipoRiversamento As Integer = 0
    Dim _nTotImpRiversato As Double = 0

    Public Property sCodConcessione() As String
        Get
            Return _sCodConcessione
        End Get

        Set(ByVal Value As String)
            _sCodConcessione = Value
        End Set
    End Property
    Public Property sCodBelfiore() As String
        Get
            Return _sCodBelfiore
        End Get

        Set(ByVal Value As String)
            _sCodBelfiore = Value
        End Set
    End Property
    Public Property sDataRiversamento() As String
        Get
            Return _sDataRiversamento
        End Get

        Set(ByVal Value As String)
            _sDataRiversamento = Value
        End Set
    End Property
    Public Property sTipoRiscossioni() As String
        Get
            Return _sTipoRiscossioni
        End Get

        Set(ByVal Value As String)
            _sTipoRiscossioni = Value
        End Set
    End Property
    Public Property sFiller() As String
        Get
            Return _sFiller
        End Get

        Set(ByVal Value As String)
            _sFiller = Value
        End Set
    End Property
    Public Property nNumQuietanza() As Integer
        Get
            Return _nNumQuietanza
        End Get

        Set(ByVal Value As Integer)
            _nNumQuietanza = Value
        End Set
    End Property
    Public Property nProgressivoRc() As Integer
        Get
            Return _nProgressivoRc
        End Get

        Set(ByVal Value As Integer)
            _nProgressivoRc = Value
        End Set
    End Property
    Public Property nCodTesoreria() As Integer
        Get
            Return _nCodTesoreria
        End Get

        Set(ByVal Value As Integer)
            _nCodTesoreria = Value
        End Set
    End Property
    Public Property nCommissione() As Integer
        Get
            Return _nCommissione
        End Get

        Set(ByVal Value As Integer)
            _nCommissione = Value
        End Set
    End Property
    Public Property nTotNumRiscossioni() As Integer
        Get
            Return _nTotNumRiscossioni
        End Get

        Set(ByVal Value As Integer)
            _nTotNumRiscossioni = Value
        End Set
    End Property
    Public Property nTipoRiversamento() As Integer
        Get
            Return _nTipoRiversamento
        End Get

        Set(ByVal Value As Integer)
            _nTipoRiversamento = Value
        End Set
    End Property
    Public Property nTotImpRiversato() As Double
        Get
            Return _nTotImpRiversato
        End Get

        Set(ByVal Value As Double)
            _nTotImpRiversato = Value
        End Set
    End Property
End Class

Public Class ExportContabileICI
    Dim _sCodConcessione As String = "1"
    Dim _sCodBelfiore As String = ""
    Dim _sDataVersamento As String = ""
    Dim _sCFPIVA As String = ""
    Dim _sDataAccredito As String = ""
    Dim _sComune As String = ""
    Dim _sCAP As String = ""
    Dim _sNumProvLiq As String = ""
    Dim _sDataProvLiq As String = ""
    Dim _sNumRifQuietanza As String = ""
    Dim _sFiller As String = "0"
    Dim _nNumQuietanza As Integer = 0
    Dim _nProgressivoRc As Integer = 0
    Dim _nAnnoImposta As Integer = 0
    Dim _nFlagQuadratura As Integer = 0
    Dim _nQuadratura As Integer = 0
    Dim _nReperibilita As Integer = 0
    Dim _nTipoVersamento As Integer = 0
    Dim _nCompentenza As Integer = 0
    Dim _nNumFab As Integer = 0
    Dim _nFlagAS As Integer = 0
    Dim _nIdentificazione As Integer = 0
    Dim _nRavvedimento As Integer = 0
    Dim _nImpVersato As Double = 0
    Dim _nImpTerAgr As Double = 0
    Dim _nImpAreeFab As Double = 0
    Dim _nImpAbiPrinc As Double = 0
    Dim _nImpAltriFab As Double = 0
    Dim _nImpDetrazione As Double = 0
    Dim _bIsViolazione As Boolean = False

    Public Property sCodConcessione() As String
        Get
            Return _sCodConcessione
        End Get

        Set(ByVal Value As String)
            _sCodConcessione = Value
        End Set
    End Property
    Public Property sCodBelfiore() As String
        Get
            Return _sCodBelfiore
        End Get

        Set(ByVal Value As String)
            _sCodBelfiore = Value
        End Set
    End Property
    Public Property sDataVersamento() As String
        Get
            Return _sDataVersamento
        End Get

        Set(ByVal Value As String)
            _sDataVersamento = Value
        End Set
    End Property
    Public Property sDataAccredito() As String
        Get
            Return _sDataAccredito
        End Get

        Set(ByVal Value As String)
            _sDataAccredito = Value
        End Set
    End Property
    Public Property sCFPIVA() As String
        Get
            Return _sCFPIVA
        End Get

        Set(ByVal Value As String)
            _sCFPIVA = Value
        End Set
    End Property
    Public Property sComune() As String
        Get
            Return _sComune
        End Get

        Set(ByVal Value As String)
            _sComune = Value
        End Set
    End Property
    Public Property sCAP() As String
        Get
            Return _sCAP
        End Get

        Set(ByVal Value As String)
            _sCAP = Value
        End Set
    End Property
    Public Property sNumProvLiq() As String
        Get
            Return _sNumProvLiq
        End Get

        Set(ByVal Value As String)
            _sNumProvLiq = Value
        End Set
    End Property
    Public Property sDataProvLiq() As String
        Get
            Return _sDataProvLiq
        End Get

        Set(ByVal Value As String)
            _sDataProvLiq = Value
        End Set
    End Property
    Public Property sNumRifQuietanza() As String
        Get
            Return _sNumRifQuietanza
        End Get

        Set(ByVal Value As String)
            _sNumRifQuietanza = Value
        End Set
    End Property
    Public Property sFiller() As String
        Get
            Return _sFiller
        End Get

        Set(ByVal Value As String)
            _sFiller = Value
        End Set
    End Property
    Public Property nNumQuietanza() As Integer
        Get
            Return _nNumQuietanza
        End Get

        Set(ByVal Value As Integer)
            _nNumQuietanza = Value
        End Set
    End Property
    Public Property nProgressivoRc() As Integer
        Get
            Return _nProgressivoRc
        End Get

        Set(ByVal Value As Integer)
            _nProgressivoRc = Value
        End Set
    End Property
    Public Property nAnnoImposta() As Integer
        Get
            Return _nAnnoImposta
        End Get

        Set(ByVal Value As Integer)
            _nAnnoImposta = Value
        End Set
    End Property
    Public Property nQuadratura() As Integer
        Get
            Return _nQuadratura
        End Get

        Set(ByVal Value As Integer)
            _nQuadratura = Value
        End Set
    End Property
    Public Property nReperibilita() As Integer
        Get
            Return _nReperibilita
        End Get

        Set(ByVal Value As Integer)
            _nReperibilita = Value
        End Set
    End Property
    Public Property nTipoVersamento() As Integer
        Get
            Return _nTipoVersamento
        End Get

        Set(ByVal Value As Integer)
            _nTipoVersamento = Value
        End Set
    End Property
    Public Property nCompentenza() As Integer
        Get
            Return _nCompentenza
        End Get

        Set(ByVal Value As Integer)
            _nCompentenza = Value
        End Set
    End Property
    Public Property nNumFab() As Integer
        Get
            Return _nNumFab
        End Get

        Set(ByVal Value As Integer)
            _nNumFab = Value
        End Set
    End Property
    Public Property nFlagAS() As Integer
        Get
            Return _nFlagAS
        End Get

        Set(ByVal Value As Integer)
            _nFlagAS = Value
        End Set
    End Property
    Public Property nIdentificazione() As Integer
        Get
            Return _nIdentificazione
        End Get

        Set(ByVal Value As Integer)
            _nIdentificazione = Value
        End Set
    End Property
    Public Property nRavvedimento() As Integer
        Get
            Return _nRavvedimento
        End Get

        Set(ByVal Value As Integer)
            _nRavvedimento = Value
        End Set
    End Property
    Public Property nImpVersato() As Double
        Get
            Return _nImpVersato
        End Get

        Set(ByVal Value As Double)
            _nImpVersato = Value
        End Set
    End Property
    Public Property nImpTerAgr() As Double
        Get
            Return _nImpTerAgr
        End Get

        Set(ByVal Value As Double)
            _nImpTerAgr = Value
        End Set
    End Property
    Public Property nImpAreeFab() As Double
        Get
            Return _nImpAreeFab
        End Get

        Set(ByVal Value As Double)
            _nImpAreeFab = Value
        End Set
    End Property
    Public Property nImpAbiPrinc() As Double
        Get
            Return _nImpAbiPrinc
        End Get

        Set(ByVal Value As Double)
            _nImpAbiPrinc = Value
        End Set
    End Property
    Public Property nImpAltriFab() As Double
        Get
            Return _nImpAltriFab
        End Get

        Set(ByVal Value As Double)
            _nImpAltriFab = Value
        End Set
    End Property
    Public Property nImpDetrazione() As Double
        Get
            Return _nImpDetrazione
        End Get

        Set(ByVal Value As Double)
            _nImpDetrazione = Value
        End Set
    End Property
    Public Property bIsViolazione() As Boolean
        Get
            Return _bIsViolazione
        End Get

        Set(ByVal Value As Boolean)
            _bIsViolazione = Value
        End Set
    End Property
End Class

Public Class ExportAnagraficoICI
    Dim _sCodConcessione As String = "1"
    Dim _sCodBelfiore As String = ""
    Dim _sSesso As String = "F"
    Dim _sCognomeRagSoc As String = ""
    Dim _sNome As String = ""
    Dim _sComune As String = ""
    Dim _sFiller As String = "0"
    Dim _nNumQuietanza As Integer = 0
    Dim _nProgressivoRc As Integer = 0

    Public Property sCodConcessione() As String
        Get
            Return _sCodConcessione
        End Get

        Set(ByVal Value As String)
            _sCodConcessione = Value
        End Set
    End Property
    Public Property sCodBelfiore() As String
        Get
            Return _sCodBelfiore
        End Get

        Set(ByVal Value As String)
            _sCodBelfiore = Value
        End Set
    End Property
    Public Property sSesso() As String
        Get
            Return _sSesso
        End Get

        Set(ByVal Value As String)
            _sSesso = Value
        End Set
    End Property
    Public Property sCognomeRagSoc() As String
        Get
            Return _sCognomeRagSoc
        End Get

        Set(ByVal Value As String)
            _sCognomeRagSoc = Value
        End Set
    End Property
    Public Property sNome() As String
        Get
            Return _sNome
        End Get

        Set(ByVal Value As String)
            _sNome = Value
        End Set
    End Property
    Public Property sComune() As String
        Get
            Return _sComune
        End Get

        Set(ByVal Value As String)
            _sComune = Value
        End Set
    End Property
    Public Property sFiller() As String
        Get
            Return _sFiller
        End Get

        Set(ByVal Value As String)
            _sFiller = Value
        End Set
    End Property
    Public Property nNumQuietanza() As Integer
        Get
            Return _nNumQuietanza
        End Get

        Set(ByVal Value As Integer)
            _nNumQuietanza = Value
        End Set
    End Property
    Public Property nProgressivoRc() As Integer
        Get
            Return _nProgressivoRc
        End Get

        Set(ByVal Value As Integer)
            _nProgressivoRc = Value
        End Set
    End Property
End Class