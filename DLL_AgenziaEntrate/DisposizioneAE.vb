Namespace AgenziaEntrate
    <Serializable()> _
        Public Class DisposizioneAE
        Private _nIDDisposizione As Integer
        Private _nIDCollegamento As Integer
        Private _nIDFlusso As Integer
        Private _sCodISTAT As String
        Private _sTributo As String
        Private _sAnno As String
        Private _sCodFiscaleEnte As String
        Private _sCognomeEnte As String
        Private _sNomeEnte As String
        Private _sSessoEnte As String
        Private _sDataNascitaEnte As String
        Private _sComuneNascitaSedeEnte As String
        Private _sPVNascitaSedeEnte As String
        Private _nIDContribuente As Integer
        Private _sCodFiscale As String
        Private _sCognome As String
        Private _sNome As String
        Private _sSesso As String
        Private _sDataNascita As String
        Private _sComuneNascitaSede As String
        Private _sPVNascitaSede As String
        Private _sComuneDomFisc As String
        Private _sPVDomFisc As String
        Private _nIDTitoloOccupazione As AgenziaEntrate.Generale.TitoloOccupazione
        Private _nIDTipoOccupante As AgenziaEntrate.Generale.NaturaOccupazione
        Private _sEstremiContratto As String
        Private _sTipoContratto As String
        Private _sDataInizio As String
        Private _sDataFine As String
        Private _nIDDestinazioneUso As AgenziaEntrate.Generale.DestinazioneUso
        Private _nIDTipoUtenza As AgenziaEntrate.Generale.TipologiaUtenza
        Private _sComuneAmmUbicazione As String
        Private _sPVAmmUbicazione As String
        Private _sComuneCatastUbicazione As String
        Private _sCodComuneUbicazioneCatast As String
        Private _sIDTipoUnita As String
        Private _sSezione As String
        Private _sFoglio As String
        Private _sParticella As String
        Private _sEstensioneParticella As String
        Private _sIDTipoParticella As String
        Private _sSubalterno As String
        Private _sIndirizzo As String
        Private _sCivico As String
        Private _sInterno As String
        Private _sScala As String
        Private _nIDAssenzaDatiCatastali As AgenziaEntrate.Generale.AssenzaDatiCatastali
        Private _nMesiFatturazione As Integer
        Private _sSegno As String
        Private _nConsumo As Integer
        Private _nImportoFatturato As Double

        Public Sub New()
            _nIDDisposizione = -1
            _nIDCollegamento = -1
            _nIDFlusso = -1
            _sCodISTAT = ""
            _sTributo = ""
            _sAnno = ""
            _sCodFiscaleEnte = ""
            _sCognomeEnte = ""
            _sNomeEnte = ""
            _sSessoEnte = ""
            _sDataNascitaEnte = ""
            _sComuneNascitaSedeEnte = ""
            _sPVNascitaSedeEnte = ""
            _nIDContribuente = -1
            _sCodFiscale = ""
            _sCognome = ""
            _sNome = ""
            _sSesso = ""
            _sDataNascita = ""
            _sComuneNascitaSede = ""
            _sPVNascitaSede = ""
            _sComuneDomFisc = ""
            _sPVDomFisc = ""
            _nIDTitoloOccupazione = AgenziaEntrate.Generale.TitoloOccupazione.NonSpecificato
            _nIDTipoOccupante = AgenziaEntrate.Generale.NaturaOccupazione.NonSpecificato
            _sEstremiContratto = ""
            _sTipoContratto = ""
            _sDataInizio = ""
            _sDataFine = ""
            _nIDDestinazioneUso = AgenziaEntrate.Generale.DestinazioneUso.NonSpecificato
            _nIDTipoUtenza = AgenziaEntrate.Generale.TipologiaUtenza.NonSpecificato
            _sComuneAmmUbicazione = ""
            _sPVAmmUbicazione = ""
            _sComuneCatastUbicazione = ""
            _sCodComuneUbicazioneCatast = ""
            _sIDTipoUnita = ""
            _sSezione = ""
            _sFoglio = ""
            _sParticella = ""
            _sEstensioneParticella = ""
            _sIDTipoParticella = ""
            _sSubalterno = ""
            _sIndirizzo = ""
            _sCivico = ""
            _sInterno = ""
            _sScala = ""
            _nIDAssenzaDatiCatastali = AgenziaEntrate.Generale.AssenzaDatiCatastali.NonSpecificato
            _nMesiFatturazione = -1
            _sSegno = ""
            _nConsumo = -1
            _nImportoFatturato = 0
        End Sub

        Public Property nIDDisposizione() As Integer
            Get
                Return _nIDDisposizione
            End Get
            Set(ByVal Value As Integer)
                _nIDDisposizione = Value
            End Set
        End Property
        Public Property nIDCollegamento() As Integer
            Get
                Return _nIDCollegamento
            End Get
            Set(ByVal Value As Integer)
                _nIDCollegamento = Value
            End Set
        End Property
        Public Property nIDFlusso() As Integer
            Get
                Return _nIDFlusso
            End Get
            Set(ByVal Value As Integer)
                _nIDFlusso = Value
            End Set
        End Property
        Public Property nIDContribuente() As Integer
            Get
                Return _nIDContribuente
            End Get
            Set(ByVal Value As Integer)
                _nIDContribuente = Value
            End Set
        End Property
        Public Property nMesiFatturazione() As Integer
            Get
                Return _nMesiFatturazione
            End Get
            Set(ByVal Value As Integer)
                _nMesiFatturazione = Value
            End Set
        End Property
        Public Property nConsumo() As Integer
            Get
                Return _nConsumo
            End Get
            Set(ByVal Value As Integer)
                _nConsumo = Value
            End Set
        End Property
        Public Property nImportoFatturato() As Double
            Get
                Return _nImportoFatturato
            End Get
            Set(ByVal Value As Double)
                _nImportoFatturato = Value
            End Set
        End Property
        Public Property nIDTitoloOccupazione() As AgenziaEntrate.Generale.TitoloOccupazione
            Get
                Return _nIDTitoloOccupazione
            End Get
            Set(ByVal Value As AgenziaEntrate.Generale.TitoloOccupazione)
                _nIDTitoloOccupazione = Value
            End Set
        End Property
        Public Property nIDTipoOccupante() As AgenziaEntrate.Generale.NaturaOccupazione
            Get
                Return _nIDTipoOccupante
            End Get
            Set(ByVal Value As AgenziaEntrate.Generale.NaturaOccupazione)
                _nIDTipoOccupante = Value
            End Set
        End Property
        Public Property nIDDestinazioneUso() As AgenziaEntrate.Generale.DestinazioneUso
            Get
                Return _nIDDestinazioneUso
            End Get
            Set(ByVal Value As AgenziaEntrate.Generale.DestinazioneUso)
                _nIDDestinazioneUso = Value
            End Set
        End Property
        Public Property nIDAssenzaDatiCatastali() As AgenziaEntrate.Generale.AssenzaDatiCatastali
            Get
                Return _nIDAssenzaDatiCatastali
            End Get
            Set(ByVal Value As AgenziaEntrate.Generale.AssenzaDatiCatastali)
                _nIDAssenzaDatiCatastali = Value
            End Set
        End Property
        Public Property nIDTipoUtenza() As AgenziaEntrate.Generale.TipologiaUtenza
            Get
                Return _nIDTipoUtenza
            End Get
            Set(ByVal Value As AgenziaEntrate.Generale.TipologiaUtenza)
                _nIDTipoUtenza = Value
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
        Public Property sTributo() As String
            Get
                Return _sTributo
            End Get
            Set(ByVal Value As String)
                _sTributo = Value
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
        Public Property sCodFiscaleEnte() As String
            Get
                Return _sCodFiscaleEnte
            End Get
            Set(ByVal Value As String)
                _sCodFiscaleEnte = Value
            End Set
        End Property
        Public Property sCognomeEnte() As String
            Get
                Return _sCognomeEnte
            End Get
            Set(ByVal Value As String)
                _sCognomeEnte = Value
            End Set
        End Property
        Public Property sNomeEnte() As String
            Get
                Return _sNomeEnte
            End Get
            Set(ByVal Value As String)
                _sNomeEnte = Value
            End Set
        End Property
        Public Property sSessoEnte() As String
            Get
                Return _sSessoEnte
            End Get
            Set(ByVal Value As String)
                _sSessoEnte = Value
            End Set
        End Property
        Public Property sDataNascitaEnte() As String
            Get
                Return _sDataNascitaEnte
            End Get
            Set(ByVal Value As String)
                _sDataNascitaEnte = Value
            End Set
        End Property
        Public Property sComuneNascitaSedeEnte() As String
            Get
                Return _sComuneNascitaSedeEnte
            End Get
            Set(ByVal Value As String)
                _sComuneNascitaSedeEnte = Value
            End Set
        End Property
        Public Property sPVNascitaSedeEnte() As String
            Get
                Return _sPVNascitaSedeEnte
            End Get
            Set(ByVal Value As String)
                _sPVNascitaSedeEnte = Value
            End Set
        End Property
        Public Property sCodFiscale() As String
            Get
                Return _sCodFiscale
            End Get
            Set(ByVal Value As String)
                _sCodFiscale = Value
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
        Public Property sComuneNascitaSede() As String
            Get
                Return _sComuneNascitaSede
            End Get
            Set(ByVal Value As String)
                _sComuneNascitaSede = Value
            End Set
        End Property
        Public Property sPVNascitaSede() As String
            Get
                Return _sPVNascitaSede
            End Get
            Set(ByVal Value As String)
                _sPVNascitaSede = Value
            End Set
        End Property
        Public Property sComuneDomFisc() As String
            Get
                Return _sComuneDomFisc
            End Get
            Set(ByVal Value As String)
                _sComuneDomFisc = Value
            End Set
        End Property
        Public Property sPVDomFisc() As String
            Get
                Return _sPVDomFisc
            End Get
            Set(ByVal Value As String)
                _sPVDomFisc = Value
            End Set
        End Property
        Public Property sEstremiContratto() As String
            Get
                Return _sEstremiContratto
            End Get
            Set(ByVal Value As String)
                _sEstremiContratto = Value
            End Set
        End Property
        Public Property sTipoContratto() As String
            Get
                Return _sTipoContratto
            End Get
            Set(ByVal Value As String)
                _sTipoContratto = Value
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
        Public Property sComuneAmmUbicazione() As String
            Get
                Return _sComuneAmmUbicazione
            End Get
            Set(ByVal Value As String)
                _sComuneAmmUbicazione = Value
            End Set
        End Property
        Public Property sPVAmmUbicazione() As String
            Get
                Return _sPVAmmUbicazione
            End Get
            Set(ByVal Value As String)
                _sPVAmmUbicazione = Value
            End Set
        End Property
        Public Property sComuneCatastUbicazione() As String
            Get
                Return _sComuneCatastUbicazione
            End Get
            Set(ByVal Value As String)
                _sComuneCatastUbicazione = Value
            End Set
        End Property
        Public Property sCodComuneUbicazioneCatast() As String
            Get
                Return _sCodComuneUbicazioneCatast
            End Get
            Set(ByVal Value As String)
                _sCodComuneUbicazioneCatast = Value
            End Set
        End Property
        Public Property sIDTipoUnita() As String
            Get
                Return _sIDTipoUnita
            End Get
            Set(ByVal Value As String)
                _sIDTipoUnita = Value
            End Set
        End Property
        Public Property sSezione() As String
            Get
                Return _sSezione
            End Get
            Set(ByVal Value As String)
                _sSezione = Value
            End Set
        End Property
        Public Property sFoglio() As String
            Get
                Return _sFoglio
            End Get
            Set(ByVal Value As String)
                _sFoglio = Value
            End Set
        End Property
        Public Property sParticella() As String
            Get
                Return _sParticella
            End Get
            Set(ByVal Value As String)
                _sParticella = Value
            End Set
        End Property
        Public Property sEstensioneParticella() As String
            Get
                Return _sEstensioneParticella
            End Get
            Set(ByVal Value As String)
                _sEstensioneParticella = Value
            End Set
        End Property
        Public Property sIDTipoParticella() As String
            Get
                Return _sIDTipoParticella
            End Get
            Set(ByVal Value As String)
                _sIDTipoParticella = Value
            End Set
        End Property
        Public Property sSubalterno() As String
            Get
                Return _sSubalterno
            End Get
            Set(ByVal Value As String)
                _sSubalterno = Value
            End Set
        End Property
        Public Property sIndirizzo() As String
            Get
                Return _sIndirizzo
            End Get
            Set(ByVal Value As String)
                _sIndirizzo = Value
            End Set
        End Property
        Public Property sCivico() As String
            Get
                Return _sCivico
            End Get
            Set(ByVal Value As String)
                _sCivico = Value
            End Set
        End Property
        Public Property sInterno() As String
            Get
                Return _sInterno
            End Get
            Set(ByVal Value As String)
                _sInterno = Value
            End Set
        End Property
        Public Property sScala() As String
            Get
                Return _sScala
            End Get
            Set(ByVal Value As String)
                _sScala = Value
            End Set
        End Property
        Public Property sSegno() As String
            Get
                Return _sSegno
            End Get
            Set(ByVal Value As String)
                _sSegno = Value
            End Set
        End Property
    End Class
End Namespace