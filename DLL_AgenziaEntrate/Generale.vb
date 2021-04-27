Namespace AgenziaEntrate
    <Serializable()>
    Public Class Generale
        Public Const DBType As String = "SQL"

        Public Enum TitoloOccupazione
            NonSpecificato = -1
            Proprieta = 1
            Usufrutto = 2
            Locazione_AltroDiritto = 3
            AltroDiritto_Rappresentante = 4
        End Enum

        Public Enum NaturaOccupazione
            NonSpecificato = -1
            Singolo = 1
            NucleoFamiliare = 2
            AttivitaCommerciale = 3
            AltraTipologia = 4
        End Enum

        Public Enum DestinazioneUso
            NonSpecificato = -1
            Abitativo = 1
            TenutoADisposizione = 2
            Commerciale = 3
            Box = 4
            AltriUsi = 5
        End Enum

        Public Enum AssenzaDatiCatastali
            NonSpecificato = -1
            NonAccatastato = 1
            NonAccatastabile = 2
            DatiNonDisponibili_Preesistente = 3
            OmessaDichiarazione = 4
            FornitureTemporanee = 5
            Condominii = 6
        End Enum

        Public Enum TipologiaUtenza
            NonSpecificato = -1
            DomesticaRes = 0
            DomesticaNonRes = 1
            NonDomestica = 2
            GrandeUtenza = 3
        End Enum
    End Class
End Namespace