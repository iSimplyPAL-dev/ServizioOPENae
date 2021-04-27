Imports System
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports RIBESFrameWork
Imports System.Xml

Namespace DLL

    '=============================================
    '
    'Gestione Eccezioni Personalizzata
    '
    '=============================================
    Public Class ExceptionAnagrafica

        Inherits Exception
        Public Sub New(ByVal strMessage As String)
            MyBase.New(strMessage)
        End Sub

    End Class

    Public Class GetListaContatti

        Public oConn As SqlConnection
        Public Query As String
        Public QueryCount As String
        Public RecordCount As Long
        Public TableName As String

    End Class

    Public Class GetList

        Public oConn As SqlConnection
        Public oComm As SqlCommand
        Public lngRecordCount As Integer

    End Class

    <Serializable()> Public Class DettaglioAnagraficaReturn

        Private m_COD_CONTRIBUENTE As String = ""
        Private m_CODICEFISCALE As String = ""

        '****************************************************************************
        'COD_CONTRIBUENTE
        '****************************************************************************
        Public Property COD_CONTRIBUENTE() As String
            Get
                Return m_COD_CONTRIBUENTE
            End Get
            Set(ByVal Value As String)
                m_COD_CONTRIBUENTE = Value
            End Set
        End Property
        '****************************************************************************
        'CODICEFISCALE
        '****************************************************************************
        Public Property CODICEFISCALE() As String
            Get
                Return m_CODICEFISCALE
            End Get
            Set(ByVal Value As String)
                m_CODICEFISCALE = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************


    End Class

    <Serializable()> Public Class RicercaComune


        Private m_COMUNE As String = ""
        Private m_PROVINCIA As String = ""
        Private m_CAP As String = ""


        '****************************************************************************
        'COMUNE
        '****************************************************************************
        Public Property COMUNE() As String
            Get
                Return m_COMUNE
            End Get
            Set(ByVal Value As String)
                m_COMUNE = Value
            End Set
        End Property
        '****************************************************************************
        'PROVINCIA
        '****************************************************************************
        Public Property PROVINCIA() As String
            Get
                Return m_PROVINCIA
            End Get
            Set(ByVal Value As String)
                m_PROVINCIA = Value
            End Set
        End Property
        '****************************************************************************
        'CAP
        '****************************************************************************
        Public Property CAP() As String
            Get
                Return m_CAP
            End Get
            Set(ByVal Value As String)
                m_CAP = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************


    End Class

    <Serializable()> Public Class DettaglioAnagrafica
        Inherits CRUD

        Public m_COD_CONTRIBUENTE As String = ""
        Public m_ID_DATA_ANAGRAFICA As String = ""
        Public m_ID_DATA_SPEDIZIONE As String = ""
        '//**************DATI NASCITA*******************
        Public m_Cognome As String = ""
        Public m_Nome As String = ""
        Public m_CodiceFiscale As String = ""
        Public m_PartitaIva As String = ""
        Public m_CodiceComuneNascita As String = ""
        Public m_ComuneNascita As String = ""
        Public m_ProvinciaNascita As String = ""
        Public m_DataNascita As String = ""
        Public m_NazionalitaNascita As String = ""
        '//*************FINE DATI NASCITA****************
        Public m_Sesso As String = ""
        '//**************DATI RESIDENZA*******************
        Public m_CodiceComuneResidenza As String = ""
        Public m_ComuneResidenza As String = ""
        Public m_ProvinciaResidenza As String = ""
        Public m_CapResidenza As String = ""
        Public m_CodViaResidenza As String = ""
        Public m_ViaResidenza As String = ""
        Public m_PosizioneCivicoResidenza As String = ""
        Public m_CivicoResidenza As String = ""
        Public m_EsponenteCivicoResidenza As String = ""
        Public m_ScalaCivicoResidenza As String = ""
        Public m_InternoCivicoResidenza As String = ""
        Public m_FrazioneResidenza As String = ""
        Public m_NazionalitaResidenza As String = ""
        '//**************FINE DATI RESIDENZA***************
        '//**************DATI GENERICI***************
        Public m_NucleoFamiliare As String = ""
        Public m_RappresentanteLegale As String = ""
        Public m_DataMorte As String = ""
        Public m_Professione As String = ""
        Public m_Note As String = ""
        Public m_DaRicontrollare As Boolean = False
        Public m_Data_Inizio_Validita As String = ""
        Public m_Data_Fine_Validita As String = ""
        Public m_Data_Ultima_Modifica As String = ""
        Public m_Operatore As String = ""
        Public m_Cod_Contribuente_Rapp_Legale As String = ""
        Public m_CodEnte As String = ""
        Public m_CodIndividuale As String = ""
        Public m_CodFamiglia As String = ""
        Public m_NC_Tributari As String = ""
        Public m_Data_Ultimo_Agg_Tributi As String = ""
        Public m_NC_Anagrafica_Res As String = ""
        Public m_Data_Ultimo_Agg_Anagrafe As String = ""
        Public m_COD_TRIBUTO As String = ""

        '//**************FINE DATI GENERICI***************
        '//**************DATI SPEDIZIONE*****************
        Public m_COGNOME_INVIO As String = ""
        Public m_NOME_INVIO As String = ""
        Public m_COD_COMUNE_RCP As String = ""
        Public m_COMUNE_RCP As String = ""
        Public m_LOC_RCP As String = ""
        Public m_PROVINCIA_RCP As String = ""
        Public m_CAP_RCP As String = ""
        Public m_COD_VIA_RCP As String = ""
        Public m_VIA_RCP As String = ""
        Public m_POSIZIONE_CIV_RCP As String = ""
        Public m_CIVICO_RCP As String = ""
        Public m_ESPONENTE_CIVICO_RCP As String = ""
        Public m_SCALA_CIVICO_RCP As String = ""
        Public m_INTERNO_CIVICO_RCP As String = ""
        Public m_FRAZIONE_RCP As String = ""
        Public m_DATA_INIZIO_VALIDITA_SPED As String = ""
        Public m_DATA_FINE_VALIDITA_SPED As String = ""
        Public m_DATA_ULTIMA_MODIFICA_SPED As String = ""
        Public m_OPERATORE_SPEDIZIONE As String = ""
        '//**************FINE DATI SPEDIZIONE***************
        '//**************DATI CONTATTI*****************
        Public m_TipoRiferimento As String = ""
        Public m_DatiRiferimento As String = ""
        Public m_dsContatti As dsContatti = Nothing
        Public m_dsTipiContatti As DataSet = Nothing
        '//**************FINE DATI CONTATTI************

        '======================PROPRIETA'======================================
        '****************************CODICE CONTRIBUENTE****************************************
        Public Property COD_CONTRIBUENTE() As String
            Get
                Return m_COD_CONTRIBUENTE
            End Get
            Set(ByVal Value As String)
                m_COD_CONTRIBUENTE = Value
            End Set
        End Property
        '****************************CODICE CONTRIBUENTE****************************************
        '***************************ID_DATA_ANAGRAFICA****************************************

        Public Property ID_DATA_SPEDIZIONE() As String
            Get
                Return m_ID_DATA_SPEDIZIONE
            End Get
            Set(ByVal Value As String)
                m_ID_DATA_SPEDIZIONE = Value
            End Set

        End Property
        '***************************ID_DATA_ANAGRAFICA****************************************
        '***************************ID_DATA_SPEDIZIONE***************************************

        Public Property ID_DATA_ANAGRAFICA() As String
            Get
                Return m_ID_DATA_ANAGRAFICA
            End Get
            Set(ByVal Value As String)
                m_ID_DATA_ANAGRAFICA = Value
            End Set
        End Property
        '***************************ID_DATA_SPEDIZIONE****************************************
        '*************************************DATI NASCITA***************************************
        '****************************************************************************************
        Public Property Cognome() As String
            Get
                Return m_Cognome
            End Get
            Set(ByVal Value As String)
                m_Cognome = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property RappresentanteLegale() As String
            Get
                Return m_RappresentanteLegale
            End Get
            Set(ByVal Value As String)
                m_RappresentanteLegale = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property Nome() As String
            Get
                Return m_Nome
            End Get
            Set(ByVal Value As String)
                m_Nome = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property CodiceFiscale() As String
            Get
                Return m_CodiceFiscale
            End Get
            Set(ByVal Value As String)
                m_CodiceFiscale = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property PartitaIva() As String
            Get
                Return m_PartitaIva
            End Get
            Set(ByVal Value As String)
                m_PartitaIva = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property CodiceComuneNascita() As String
            Get
                Return m_CodiceComuneNascita
            End Get
            Set(ByVal Value As String)
                m_CodiceComuneNascita = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property ComuneNascita() As String
            Get
                Return m_ComuneNascita
            End Get
            Set(ByVal Value As String)
                m_ComuneNascita = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property ProvinciaNascita() As String
            Get
                Return m_ProvinciaNascita
            End Get
            Set(ByVal Value As String)
                m_ProvinciaNascita = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property DataNascita() As String
            Get
                Return m_DataNascita
            End Get
            Set(ByVal Value As String)
                m_DataNascita = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property NazionalitaNascita() As String
            Get
                Return m_NazionalitaNascita
            End Get
            Set(ByVal Value As String)
                m_NazionalitaNascita = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property Sesso() As String
            Get
                Return m_Sesso
            End Get
            Set(ByVal Value As String)
                m_Sesso = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************

        '***************************************************************************
        '*****************************DATI RESIDENZA*******************************
        '***************************************************************************
        '***************************************************************************
        '***************************************************************************
        Public Property CodiceComuneResidenza() As String
            Get
                Return m_CodiceComuneResidenza
            End Get
            Set(ByVal Value As String)
                m_CodiceComuneResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property ComuneResidenza() As String
            Get
                Return m_ComuneResidenza
            End Get
            Set(ByVal Value As String)
                m_ComuneResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property ProvinciaResidenza() As String
            Get
                Return m_ProvinciaResidenza
            End Get
            Set(ByVal Value As String)
                m_ProvinciaResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property CapResidenza() As String
            Get
                Return m_CapResidenza
            End Get
            Set(ByVal Value As String)
                m_CapResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property CodViaResidenza() As String
            Get
                Return m_CodViaResidenza
            End Get
            Set(ByVal Value As String)
                m_CodViaResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property ViaResidenza() As String
            Get
                Return m_ViaResidenza
            End Get
            Set(ByVal Value As String)
                m_ViaResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property PosizioneCivicoResidenza() As String
            Get
                Return m_PosizioneCivicoResidenza
            End Get
            Set(ByVal Value As String)
                m_PosizioneCivicoResidenza = Value
            End Set
        End Property

        '****************************************************************************
        '****************************************************************************
        Public Property CivicoResidenza() As String
            Get
                Return m_CivicoResidenza
            End Get
            Set(ByVal Value As String)
                m_CivicoResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property EsponenteCivicoResidenza() As String
            Get
                Return m_EsponenteCivicoResidenza
            End Get
            Set(ByVal Value As String)
                m_EsponenteCivicoResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property ScalaCivicoResidenza() As String
            Get
                Return m_ScalaCivicoResidenza
            End Get
            Set(ByVal Value As String)
                m_ScalaCivicoResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property InternoCivicoResidenza() As String
            Get
                Return m_InternoCivicoResidenza
            End Get
            Set(ByVal Value As String)
                m_InternoCivicoResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property FrazioneResidenza() As String
            Get
                Return m_FrazioneResidenza
            End Get
            Set(ByVal Value As String)
                m_FrazioneResidenza = Value
            End Set
        End Property
        '****************************************************************************
        '****************************************************************************
        Public Property NazionalitaResidenza() As String
            Get
                Return m_NazionalitaResidenza
            End Get
            Set(ByVal Value As String)
                m_NazionalitaResidenza = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        '***************************************************************************
        '*****************************FINE DATI RESIDENZA***************************
        '***************************************************************************
        '***************************************************************************
        '*****************************DATI GENERICI*********************************
        '***************************************************************************
        '***************************************************************************
        Public Property NucleoFamiliare() As String
            Get
                Return m_NucleoFamiliare
            End Get
            Set(ByVal Value As String)
                m_NucleoFamiliare = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataMorte() As String
            Get
                Return m_DataMorte
            End Get
            Set(ByVal Value As String)
                m_DataMorte = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property Professione() As String
            Get
                Return m_Professione
            End Get
            Set(ByVal Value As String)
                m_Professione = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property Note() As String
            Get
                Return m_Note
            End Get
            Set(ByVal Value As String)
                m_Note = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DaRicontrollare() As Boolean
            Get
                Return m_DaRicontrollare
            End Get
            Set(ByVal Value As Boolean)
                m_DaRicontrollare = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataInizioValidita() As String
            Get
                Return m_Data_Inizio_Validita
            End Get
            Set(ByVal Value As String)
                m_Data_Inizio_Validita = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataFineValidita() As String
            Get
                Return m_Data_Fine_Validita
            End Get
            Set(ByVal Value As String)
                m_Data_Fine_Validita = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataUltimaModifica() As String
            Get
                Return m_Data_Ultima_Modifica
            End Get
            Set(ByVal Value As String)
                m_Data_Ultima_Modifica = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property Operatore() As String
            Get
                Return m_Operatore
            End Get
            Set(ByVal Value As String)
                m_Operatore = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodContribuenteRappLegale() As String
            Get
                Return m_Cod_Contribuente_Rapp_Legale
            End Get
            Set(ByVal Value As String)
                m_Cod_Contribuente_Rapp_Legale = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodEnte() As String
            Get
                Return m_CodEnte
            End Get
            Set(ByVal Value As String)
                m_CodEnte = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodIndividuale() As String
            Get
                Return m_CodIndividuale
            End Get
            Set(ByVal Value As String)
                m_CodIndividuale = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodFamiglia() As String
            Get
                Return m_CodFamiglia
            End Get
            Set(ByVal Value As String)
                m_CodFamiglia = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property NCTributari() As String
            Get
                Return m_NC_Tributari
            End Get
            Set(ByVal Value As String)
                m_NC_Tributari = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataUltimoAggTributi() As String
            Get
                Return m_Data_Ultimo_Agg_Tributi
            End Get
            Set(ByVal Value As String)
                m_Data_Ultimo_Agg_Tributi = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property NCAnagraficaRes() As String
            Get
                Return m_NC_Anagrafica_Res
            End Get
            Set(ByVal Value As String)
                m_NC_Anagrafica_Res = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataUltimoAggAnagrafe() As String
            Get
                Return m_Data_Ultimo_Agg_Anagrafe
            End Get
            Set(ByVal Value As String)
                m_Data_Ultimo_Agg_Anagrafe = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodTributo() As String
            Get
                Return m_COD_TRIBUTO
            End Get
            Set(ByVal Value As String)
                m_COD_TRIBUTO = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        '*****************************FINE DATI GENERICI*****************************
        '***************************************************************************
        '***************************************************************************
        '*****************************DATI SPEDIZIONE*******************************
        '***************************************************************************
        '***************************************************************************
        Public Property CognomeInvio() As String
            Get
                Return m_COGNOME_INVIO
            End Get
            Set(ByVal Value As String)
                m_COGNOME_INVIO = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property NomeInvio() As String
            Get
                Return m_NOME_INVIO
            End Get
            Set(ByVal Value As String)
                m_NOME_INVIO = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodComuneRCP() As String
            Get
                Return m_COD_COMUNE_RCP
            End Get
            Set(ByVal Value As String)
                m_COD_COMUNE_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property ComuneRCP() As String
            Get
                Return m_COMUNE_RCP
            End Get
            Set(ByVal Value As String)
                m_COMUNE_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property LocRCP() As String
            Get
                Return m_LOC_RCP
            End Get
            Set(ByVal Value As String)
                m_LOC_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property ProvinciaRCP() As String
            Get
                Return m_PROVINCIA_RCP
            End Get
            Set(ByVal Value As String)
                m_PROVINCIA_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CapRCP() As String
            Get
                Return m_CAP_RCP
            End Get
            Set(ByVal Value As String)
                m_CAP_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CodViaRCP() As String
            Get
                Return m_COD_VIA_RCP
            End Get
            Set(ByVal Value As String)
                m_COD_VIA_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property ViaRCP() As String
            Get
                Return m_VIA_RCP
            End Get
            Set(ByVal Value As String)
                m_VIA_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property PosizioneCivicoRCP() As String
            Get
                Return m_POSIZIONE_CIV_RCP
            End Get
            Set(ByVal Value As String)
                m_POSIZIONE_CIV_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property CivicoRCP() As String
            Get
                Return m_CIVICO_RCP
            End Get
            Set(ByVal Value As String)
                m_CIVICO_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property EsponenteCivicoRCP() As String
            Get
                Return m_ESPONENTE_CIVICO_RCP
            End Get
            Set(ByVal Value As String)
                m_ESPONENTE_CIVICO_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property ScalaCivicoRCP() As String
            Get
                Return m_SCALA_CIVICO_RCP
            End Get
            Set(ByVal Value As String)
                m_SCALA_CIVICO_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property InternoCivicoRCP() As String
            Get
                Return m_INTERNO_CIVICO_RCP
            End Get
            Set(ByVal Value As String)
                m_INTERNO_CIVICO_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property FrazioneRCP() As String
            Get
                Return m_FRAZIONE_RCP
            End Get
            Set(ByVal Value As String)
                m_FRAZIONE_RCP = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        '***************************************************************************
        '***************************************************************************
        Public Property DataInizioValiditaSpedizione() As String
            Get
                Return m_DATA_INIZIO_VALIDITA_SPED
            End Get
            Set(ByVal Value As String)
                m_DATA_INIZIO_VALIDITA_SPED = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataFineValiditaSpedizione() As String
            Get
                Return m_DATA_FINE_VALIDITA_SPED
            End Get
            Set(ByVal Value As String)
                m_DATA_FINE_VALIDITA_SPED = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DataUltimaModificaSpedizione() As String
            Get
                Return m_DATA_ULTIMA_MODIFICA_SPED
            End Get
            Set(ByVal Value As String)
                m_DATA_ULTIMA_MODIFICA_SPED = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property OperatoreSpedizione() As String
            Get
                Return m_OPERATORE_SPEDIZIONE
            End Get
            Set(ByVal Value As String)
                m_OPERATORE_SPEDIZIONE = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        '*****************************FINE DATI SPEDIZIONE***************************
        '***************************************************************************
        '*****************************DATI CONTATTI***************************
        '***************************************************************************
        Public Property TipoRiferimento() As String
            Get
                Return m_TipoRiferimento
            End Get
            Set(ByVal Value As String)
                m_TipoRiferimento = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property DatiRiferimento() As String
            Get
                Return m_DatiRiferimento
            End Get
            Set(ByVal Value As String)
                m_DatiRiferimento = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property dsContatti() As dsContatti
            Get
                Return m_dsContatti
            End Get
            Set(ByVal Value As dsContatti)
                m_dsContatti = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************
        Public Property dsTipiContatti() As DataSet
            Get
                Return m_dsTipiContatti
            End Get
            Set(ByVal Value As DataSet)
                m_dsTipiContatti = Value
            End Set
        End Property
        '***************************************************************************
        '***************************************************************************

        '*****************************FINE DATI CONTATTI***************************
        '***************************************************************************
        '======================FINE PROPRIETA'==========================
    End Class


    <Serializable()> Public Class AnagraficaResidente
        Public _CodContribuente As String = ""
        Public _Cognome As String = ""
        Public _Nome As String = ""
        Public _CodiceFiscale As String = ""
        Public _ComuneNascita As String = ""
        Public _DataNascita As String = ""
        Public _DataMorte As String = ""
        Public _Sesso As String = ""
        Public _CodViaResidenza As String = ""
        Public _ViaResidenza As String = ""
        Public _CivicoResidenza As String = ""
        Public _EsponenteCivicoResidenza As String = ""
        Public _InternoCivicoResidenza As String = ""
        Public _Azione As String = ""
        Public _DescrParentela As String = ""
        Public _CodFamiglia As String = ""

        Public Property CodContribuente() As String
            Get
                Return _CodContribuente
            End Get
            Set(ByVal Value As String)
                _CodContribuente = Value
            End Set
        End Property
        Public Property Cognome() As String
            Get
                Return _Cognome
            End Get
            Set(ByVal Value As String)
                _Cognome = Value
            End Set
        End Property
        Public Property Nome() As String
            Get
                Return _Nome
            End Get
            Set(ByVal Value As String)
                _Nome = Value
            End Set
        End Property
        Public Property CodiceFiscale() As String
            Get
                Return _CodiceFiscale
            End Get
            Set(ByVal Value As String)
                _CodiceFiscale = Value
            End Set
        End Property
        Public Property ComuneNascita() As String
            Get
                Return _ComuneNascita
            End Get
            Set(ByVal Value As String)
                _ComuneNascita = Value
            End Set
        End Property
        Public Property DataNascita() As String
            Get
                Return _DataNascita
            End Get
            Set(ByVal Value As String)
                _DataNascita = Value
            End Set
        End Property
        Public Property DataMorte() As String
            Get
                Return _DataMorte
            End Get
            Set(ByVal Value As String)
                _DataMorte = Value
            End Set
        End Property
        Public Property Sesso() As String
            Get
                Return _Sesso
            End Get
            Set(ByVal Value As String)
                _Sesso = Value
            End Set
        End Property
        Public Property CodViaResidenza() As String
            Get
                Return _CodViaResidenza
            End Get
            Set(ByVal Value As String)
                _CodViaResidenza = Value
            End Set
        End Property
        Public Property ViaResidenza() As String
            Get
                Return _ViaResidenza
            End Get
            Set(ByVal Value As String)
                _ViaResidenza = Value
            End Set
        End Property
        Public Property CivicoResidenza() As String
            Get
                Return _CivicoResidenza
            End Get
            Set(ByVal Value As String)
                _CivicoResidenza = Value
            End Set
        End Property
        Public Property EsponenteCivicoResidenza() As String
            Get
                Return _EsponenteCivicoResidenza
            End Get
            Set(ByVal Value As String)
                _EsponenteCivicoResidenza = Value
            End Set
        End Property
        Public Property InternoCivicoResidenza() As String
            Get
                Return _InternoCivicoResidenza
            End Get
            Set(ByVal Value As String)
                _InternoCivicoResidenza = Value
            End Set
        End Property
        Public Property Azione() As String
            Get
                Return _Azione
            End Get
            Set(ByVal Value As String)
                _Azione = Value
            End Set
        End Property
        Public Property DescrParentela() As String
            Get
                Return _DescrParentela
            End Get
            Set(ByVal Value As String)
                _DescrParentela = Value
            End Set
        End Property
        Public Property CodFamiglia() As String
            Get
                Return _CodFamiglia
            End Get
            Set(ByVal Value As String)
                _CodFamiglia = Value
            End Set
        End Property
    End Class

    '//Business Interface

    Public Class GestioneAnagrafica

        '//Oggetto Data Interface accesso al Data Base

        Dim DBAccess As New getDBobject
        '===============================================================================
        'WorkFlow
        '===============================================================================
        Dim m_oSession As New RIBESFrameWork.Session
        Dim m_IDSottoAttivita As String
        '===============================================================================
        'WorkFlow
        '===============================================================================

        '//Costanti

        Dim Costant As New Costanti
        '//Utility 
        Dim Utility As New Utility
        Dim ModDate As New GestDate
        Dim DataSetContatti As New DataSetContatti

        Enum DBOperation
            DB_INSERT = 1
            DB_UPDATE = 0
        End Enum
        '===============================================================================
        'WorkFlow
        '===============================================================================
        Public Sub New(ByVal oSession As RIBESFrameWork.Session, ByVal IDSottoAttivita As String)
            m_oSession = oSession
            m_IDSottoAttivita = IDSottoAttivita
        End Sub
        '===============================================================================
        'WorkFlow
        '===============================================================================

        Public Function GetListComuni(ByVal oRicercaComune As RicercaComune) As DataSet
            Try

                Dim strSQL As String
                Dim dsComuni As New DataSet


                strSQL = "SELECT IDENTIFICATIVO, COMUNE, PV, CAP" & vbCrLf
                strSQL = strSQL & " FROM COMUNI" & vbCrLf
                strSQL = strSQL & " WHERE 1=1 " & vbCrLf

                If oRicercaComune.COMUNE.CompareTo("") <> "0" Then

                    strSQL = strSQL & " AND (COMUNE LIKE '" & Replace(Replace(Trim(oRicercaComune.COMUNE), "'", "''"), "*", "%") & "%')" & vbCrLf

                End If

                If oRicercaComune.CAP.CompareTo("") <> "0" Then

                    strSQL = strSQL & " AND (CAP= '" + oRicercaComune.CAP + "')" & vbCrLf

                End If

                If oRicercaComune.PROVINCIA.CompareTo("") <> "0" Then

                    strSQL = strSQL & " AND (PV= '" + oRicercaComune.PROVINCIA + "')" & vbCrLf

                End If

                strSQL = strSQL & "ORDER BY COMUNE,CAP"


                Dim objDBAccess As New RIBESFrameWork.DBManager
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)

                dsComuni = objDBAccess.GetPrivateDataSet(strSQL)

                Return dsComuni

            Catch Err As Exception
                Throw New Exception("GestioneAnagrafica::GetListComuni" & Err.Message)
            End Try

        End Function

        Public Function UpdateCodContribNelSistema(ByVal CODICE_CONTRIBUENTE_PRINCIPALE As Long, ByVal CODICE_CONTRIBUENTE_SECONDARIO As Long, ByVal IDDATAANAGRAFICA_SECONDARIA As Long, ByVal cod_ente As String) As Boolean

            Dim GetList As New GetList
            Dim objConn As New SqlConnection
            Dim objCmd As New SqlCommand

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager

            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objCmd.Parameters.Clear()

                GetList.lngRecordCount = objDBAccess.GetPrivateRunSPForRibesDataGrid("sp_UpdateContrib", objConn, objCmd, _
                New SqlParameter("@CodContribuentePrinc", CODICE_CONTRIBUENTE_PRINCIPALE), _
                New SqlParameter("@CodContribuenteSecond", CODICE_CONTRIBUENTE_SECONDARIO), _
                New SqlParameter("@IdDataAnagraficaSecond", IDDATAANAGRAFICA_SECONDARIA), _
                New SqlParameter("@codente", cod_ente))


                GetList.oConn = objConn
                GetList.oComm = objCmd
                'objCmd.Parameters.Clear() 'commentata per OpenTerritorio : errore di index nella datagrid Ale-Fabio 12/01/2005


                Return True

            Catch ex As Exception

                Throw New Exception("Anagrafica::UpdateCodContribNelSistema::" & ex.Message)
            Finally
                '********************Gestione Aggiornamento Codice Contribuente all'interno del sistema****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Aggiornamento Codice Contribuente all'interno del sistema****************************

            End Try

        End Function



#Region "RESTITUISCE IL DETTAGLIO ANAGRAFICA"
        '=================================================================================
        'Consente di caricare i dati relativi ad una Anagrafica tramite l'utilizzo della Dll che consente di accedere al DataBase tramite
        'WorkFlow
        'PARAMETRI:COD_CONTRIBUENTE,COD_TRIBUTO
        'COD_CONTRIBUENTE può essere uguale al COD_CONTRIBUENTE selezionato da una pagina di ricerca
        'COD_TRIBUTO si riferisce al tributo a cui l'Anagrafica è associata Es:H20 -->9000
        '==================================================================================

        Public Function GetAnagrafica(ByVal COD_CONTRIBUENTE As Long, ByVal COD_TRIBUTO As String) As DettaglioAnagrafica
            Dim sMyLog As String = "::si inizia"
            Dim strSql As String
            Dim lgnTipoOperazione As Long = DBOperation.DB_UPDATE
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            '===============================================================================
            'WorkFlow
            '===============================================================================

            Dim DettaglioAnagrafica As New DettaglioAnagrafica
            If COD_CONTRIBUENTE = Costant.INIT_VALUE_NUMBER Then lgnTipoOperazione = DBOperation.DB_INSERT
            '===============================================================================
            'MODIFICA
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Try
                sMyLog += "::devo inizializzare il dbmanager::" & m_IDSottoAttivita
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                sMyLog += "::valorizzo dataset contatti::SELECT * FROM TIPI_CONTATTI ORDER BY IDTipoRiferimento::connessione::" & objDBAccess.GetConnection.ConnectionString
                DettaglioAnagrafica.dsTipiContatti = AddItemToDataSet(objDBAccess.GetPrivateDataSet("SELECT * FROM TIPI_CONTATTI ORDER BY IDTipoRiferimento"))
                '===============================================================================
                'WorkFlow
                '===============================================================================
                'm_oSession.oAppDB.DisposeConnection()
                sMyLog += "::chiudo dbmanager"
                objDBAccess.DisposeConnection()

                If lgnTipoOperazione = DBOperation.DB_UPDATE Then
                    sMyLog += "::gestione update"
                    'Gestione Contatti
                    '===============================================================================
                    strSql = ""
                    strSql = "SELECT * FROM CONTATTI" & vbCrLf
                    strSql = strSql & "INNER JOIN ANAGRAFICA ON CONTATTI.COD_CONTRIBUENTE = ANAGRAFICA.COD_CONTRIBUENTE AND CONTATTI.IDDATAANAGRAFICA = ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                    strSql = strSql & "WHERE ANAGRAFICA.COD_CONTRIBUENTE =" & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "(ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')" & vbCrLf
                    strSql = strSql & "ORDER BY CONTATTI.TIPORIFERIMENTO"
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    DataSetContatti.daContattiPersona = objDBAccess.GetPrivateDataAdapter(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    ' m_oSession.oAppDB.DisposeConnection()


                    Dim ds As dsContatti = DataSetContatti.DataSetCompleto
                    'Caricamento dei contatti associati alla persona se esistenti
                    DettaglioAnagrafica.dsContatti = ds
                    objDBAccess.DisposeConnection()
                    'Caricamento dei tipi contatti 
                    '===============================================================================
                    '===============================================================================
                    'Gestione Anagrafica
                    '===============================================================================
                    '===============================================================================
                    strSql = ""
                    strSql = "SELECT * FROM ANAGRAFICA" & vbCrLf
                    strSql = strSql & "INNER JOIN" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE" & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "(ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')" & vbCrLf

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    Dim drDetailsAnagrafica As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    If drDetailsAnagrafica.Read Then
                        'Dati Nascita
                        '********************************************************************************************************************
                        '********************************************************************************************************************
                        DettaglioAnagrafica.COD_CONTRIBUENTE = Utility.GetParametro(drDetailsAnagrafica("COD_CONTRIBUENTE"))
                        DettaglioAnagrafica.ID_DATA_ANAGRAFICA = Utility.CIdFromDB(drDetailsAnagrafica("IDDATAANAGRAFICA"))
                        DettaglioAnagrafica.Cognome = Utility.GetParametro(drDetailsAnagrafica("COGNOME_DENOMINAZIONE"))
                        DettaglioAnagrafica.Nome = Utility.GetParametro(drDetailsAnagrafica("NOME"))
                        DettaglioAnagrafica.CodiceFiscale = Utility.GetParametro(drDetailsAnagrafica("COD_FISCALE"))
                        DettaglioAnagrafica.PartitaIva = Utility.GetParametro(drDetailsAnagrafica("PARTITA_IVA"))
                        DettaglioAnagrafica.CodiceComuneNascita = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_NASCITA"))
                        DettaglioAnagrafica.ComuneNascita = Utility.GetParametro(drDetailsAnagrafica("COMUNE_NASCITA"))
                        DettaglioAnagrafica.ProvinciaNascita = Utility.GetParametro(drDetailsAnagrafica("PROV_NASCITA"))
                        DettaglioAnagrafica.DataNascita = ModDate.GiraDataFromDB(Utility.GetParametro(drDetailsAnagrafica("DATA_NASCITA")))
                        DettaglioAnagrafica.DataMorte = ModDate.GiraDataFromDB(Utility.GetParametro(drDetailsAnagrafica("DATA_MORTE")))
                        DettaglioAnagrafica.NazionalitaNascita = Utility.GetParametro(drDetailsAnagrafica("NAZIONALITA_NASCITA"))
                        DettaglioAnagrafica.Sesso = Utility.GetParametro(drDetailsAnagrafica("SESSO"))

                        '===============================================================================
                        '===============================================================================
                        'Dati Residenza
                        '===============================================================================
                        '===============================================================================
                        DettaglioAnagrafica.CodiceComuneResidenza = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_RES"))
                        DettaglioAnagrafica.ComuneResidenza = Utility.GetParametro(drDetailsAnagrafica("COMUNE_RES"))
                        DettaglioAnagrafica.ProvinciaResidenza = Utility.GetParametro(drDetailsAnagrafica("PROVINCIA_RES"))
                        DettaglioAnagrafica.CapResidenza = Utility.GetParametro(drDetailsAnagrafica("CAP_RES"))
                        DettaglioAnagrafica.CodViaResidenza = Utility.GetParametro(drDetailsAnagrafica("COD_VIA_RES"))
                        DettaglioAnagrafica.ViaResidenza = Utility.GetParametro(drDetailsAnagrafica("VIA_RES"))
                        DettaglioAnagrafica.PosizioneCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("POSIZIONE_CIVICO_RES"))
                        DettaglioAnagrafica.CivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("CIVICO_RES"))
                        DettaglioAnagrafica.EsponenteCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("ESPONENTE_CIVICO_RES"))
                        DettaglioAnagrafica.ScalaCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("SCALA_CIVICO_RES"))
                        DettaglioAnagrafica.InternoCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("INTERNO_CIVICO_RES"))
                        DettaglioAnagrafica.FrazioneResidenza = Utility.GetParametro(drDetailsAnagrafica("FRAZIONE_RES"))
                        DettaglioAnagrafica.NazionalitaResidenza = Utility.GetParametro(drDetailsAnagrafica("NAZIONALITA_RES"))
                        '===============================================================================
                        '===============================================================================
                        'Dati generici
                        '===============================================================================
                        '===============================================================================
                        DettaglioAnagrafica.Professione = Utility.GetParametro(drDetailsAnagrafica("PROFESSIONE"))
                        DettaglioAnagrafica.Note = Utility.GetParametro(drDetailsAnagrafica("NOTE"))
                        DettaglioAnagrafica.DaRicontrollare = Utility.cToBool(drDetailsAnagrafica("DA_RICONTROLLARE"))
                        DettaglioAnagrafica.NucleoFamiliare = Utility.GetParametro(drDetailsAnagrafica("NUCLEO_FAMILIARE"))
                        DettaglioAnagrafica.CodContribuenteRappLegale = Utility.GetParametro(drDetailsAnagrafica("COD_CONTRIBUENTE_RAPP_LEGALE"))
                        DettaglioAnagrafica.Operatore = Utility.GetParametro(drDetailsAnagrafica("OPERATORE"))
                        '===============================================================================
                        '===============================================================================
                        'GESTIONE LOCK
                        '===============================================================================

                        If IsDBNull(drDetailsAnagrafica("CURRENCY")) Then
                            DettaglioAnagrafica.Concurrency = CType(Now, Date)
                        Else
                            DettaglioAnagrafica.Concurrency = drDetailsAnagrafica("CURRENCY")
                        End If
                        DettaglioAnagrafica.CodTributo = COD_TRIBUTO
                        '===============================================================================
                        'GESTIONE LOCK
                        '===============================================================================
                    End If
                    drDetailsAnagrafica.Close()
                    ' m_oSession.oAppDB.DisposeConnection()
                    objDBAccess.DisposeConnection()
                    If Len(DettaglioAnagrafica.CodContribuenteRappLegale) > 0 Then
                        strSql = ""
                        strSql = "SELECT * FROM ANAGRAFICA" & vbCrLf
                        strSql = strSql & "INNER JOIN" & vbCrLf
                        strSql = strSql & "DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE" & vbCrLf
                        strSql = strSql & "AND" & vbCrLf
                        strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                        strSql = strSql & "WHERE DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE = " & DettaglioAnagrafica.CodContribuenteRappLegale & vbCrLf
                        strSql = strSql & "AND" & vbCrLf
                        strSql = strSql & "(ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')" & vbCrLf
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                        drDetailsAnagrafica = objDBAccess.GetPrivateDataReader(strSql)
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        If drDetailsAnagrafica.Read Then
                            DettaglioAnagrafica.RappresentanteLegale = Utility.GetParametro(drDetailsAnagrafica("COGNOME_DENOMINAZIONE")) & " " & Utility.GetParametro(drDetailsAnagrafica("NOME"))
                        End If
                        drDetailsAnagrafica.Close()
                        'm_oSession.oAppDB.DisposeConnection()
                        objDBAccess.DisposeConnection()
                    End If
                    '===============================================================================
                    'DATI SPEDIZIONE
                    '===============================================================================
                    '===============================================================================
                    strSql = ""
                    strSql = "SELECT * " & vbCrLf
                    strSql = strSql & "FROM INDIRIZZI_SPEDIZIONE INNER JOIN" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_SPEDIZIONE ON INDIRIZZI_SPEDIZIONE.COD_TRIBUTO = DATA_VALIDITA_SPEDIZIONE.COD_TRIBUTO AND " & vbCrLf
                    strSql = strSql & "INDIRIZZI_SPEDIZIONE.COD_CONTRIBUENTE = DATA_VALIDITA_SPEDIZIONE.COD_CONTRIBUENTE AND" & vbCrLf
                    strSql = strSql & "INDIRIZZI_SPEDIZIONE.IDDATA = DATA_VALIDITA_SPEDIZIONE.IDDATA" & vbCrLf
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_SPEDIZIONE.COD_TRIBUTO = " & Utility.CStrToDB(COD_TRIBUTO) & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_SPEDIZIONE.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "(INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA IS NULL OR INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA='')"
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    Dim drDetailsSpedizione As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    If drDetailsSpedizione.Read Then

                        DettaglioAnagrafica.ID_DATA_SPEDIZIONE = Utility.GetParametro(drDetailsSpedizione("IDDATA"))
                        DettaglioAnagrafica.CognomeInvio = Utility.GetParametro(drDetailsSpedizione("COGNOME_INVIO"))
                        DettaglioAnagrafica.NomeInvio = Utility.GetParametro(drDetailsSpedizione("NOME_INVIO"))
                        DettaglioAnagrafica.CodComuneRCP = Utility.GetParametro(drDetailsSpedizione("COD_COMUNE_RCP"))
                        DettaglioAnagrafica.ComuneRCP = Utility.GetParametro(drDetailsSpedizione("COMUNE_RCP"))
                        DettaglioAnagrafica.LocRCP = Utility.GetParametro(drDetailsSpedizione("LOC_RCP"))
                        DettaglioAnagrafica.ProvinciaRCP = Utility.GetParametro(drDetailsSpedizione("PROVINCIA_RCP"))
                        DettaglioAnagrafica.CapRCP = Utility.GetParametro(drDetailsSpedizione("CAP_RCP"))
                        DettaglioAnagrafica.CodViaRCP = Utility.GetParametro(drDetailsSpedizione("COD_VIA_RCP"))
                        DettaglioAnagrafica.ViaRCP = Utility.GetParametro(drDetailsSpedizione("VIA_RCP"))
                        DettaglioAnagrafica.PosizioneCivicoRCP = Utility.GetParametro(drDetailsSpedizione("POSIZIONE_CIV_RCP"))
                        DettaglioAnagrafica.CivicoRCP = Utility.GetParametro(drDetailsSpedizione("CIVICO_RCP"))
                        DettaglioAnagrafica.EsponenteCivicoRCP = Utility.GetParametro(drDetailsSpedizione("ESPONENTE_CIVICO_RCP"))
                        DettaglioAnagrafica.ScalaCivicoRCP = Utility.GetParametro(drDetailsSpedizione("SCALA_CIVICO_RCP"))
                        DettaglioAnagrafica.InternoCivicoRCP = Utility.GetParametro(drDetailsSpedizione("INTERNO_CIVICO_RCP"))
                        DettaglioAnagrafica.FrazioneRCP = Utility.GetParametro(drDetailsSpedizione("FRAZIONE_RCP"))

                    Else

                        DettaglioAnagrafica.ID_DATA_SPEDIZIONE = Costant.INIT_VALUE_NUMBER

                    End If
                    drDetailsSpedizione.Close()
                    'm_oSession.oAppDB.DisposeConnection()
                    objDBAccess.DisposeConnection()
                End If
                '===============================================================================
                'FINE MODIFICA
                '===============================================================================

                If lgnTipoOperazione = DBOperation.DB_INSERT Then
                    sMyLog += "::gestione insert"

                    strSql = ""
                    strSql = "SELECT * FROM CONTATTI" & vbCrLf
                    strSql = strSql & "INNER JOIN ANAGRAFICA ON CONTATTI.COD_CONTRIBUENTE = ANAGRAFICA.COD_CONTRIBUENTE AND CONTATTI.IDDATAANAGRAFICA = ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                    strSql = strSql & "WHERE ANAGRAFICA.COD_CONTRIBUENTE =" & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "(ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')" & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "1=0" & vbCrLf
                    strSql = strSql & "ORDER BY CONTATTI.TIPORIFERIMENTO"
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    DataSetContatti.daContattiPersona = objDBAccess.GetPrivateDataAdapter(strSql)
                    Dim ds As dsContatti = DataSetContatti.DataSetCompleto
                    'Caricamento dei contatti associati alla persona se esistenti
                    DettaglioAnagrafica.dsContatti = ds

                    DettaglioAnagrafica.COD_CONTRIBUENTE = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.ID_DATA_ANAGRAFICA = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.Cognome = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.Nome = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CodiceFiscale = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.PartitaIva = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CodiceComuneNascita = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.ComuneNascita = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.ProvinciaNascita = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.DataNascita = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.DataMorte = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.NazionalitaNascita = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.Sesso = Costant.INIT_VALUE_NUMBER
                    '===============================================================================
                    '===============================================================================
                    'Dati Residenza
                    '===============================================================================
                    '===============================================================================
                    DettaglioAnagrafica.ID_DATA_SPEDIZIONE = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.CodiceComuneResidenza = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.ComuneResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.ProvinciaResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CapResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CodViaResidenza = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.ViaResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.PosizioneCivicoResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CivicoResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.EsponenteCivicoResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.ScalaCivicoResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.InternoCivicoResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.FrazioneResidenza = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.NazionalitaResidenza = Costant.INIT_VALUE_STRING
                    '===============================================================================
                    '===============================================================================
                    'Dati generici
                    '===============================================================================
                    '===============================================================================
                    DettaglioAnagrafica.Professione = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.Note = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.DaRicontrollare = False
                    DettaglioAnagrafica.NucleoFamiliare = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CodContribuenteRappLegale = Costant.INIT_VALUE_NUMBER

                    '===============================================================================
                    '===============================================================================
                    'DATI SPEDIZIONE
                    '===============================================================================
                    '===============================================================================
                    DettaglioAnagrafica.CognomeInvio = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.NomeInvio = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CodComuneRCP = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.ComuneRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.LocRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.ProvinciaRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CapRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CodViaRCP = Costant.INIT_VALUE_NUMBER
                    DettaglioAnagrafica.ViaRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.PosizioneCivicoRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.CivicoRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.EsponenteCivicoRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.ScalaCivicoRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.InternoCivicoRCP = Costant.INIT_VALUE_STRING
                    DettaglioAnagrafica.FrazioneRCP = Costant.INIT_VALUE_STRING

                End If


                Return DettaglioAnagrafica
            Catch ex As Exception
                Throw New Exception("Anagrafica::GetAnagrafica::" & ex.Message & "::SQL::" & strSql & "::MYLOG::" & sMyLog)
            Finally
                '********************Gestione Anagrafiche massive****************************
                If Not m_oSession.oAppDB Is Nothing Then
                    m_oSession.oAppDB.DisposeConnection()
                End If
                If Not m_oSession.oAppDB Is Nothing Then
                    m_oSession.oAppDB.Dispose()
                End If
                If Not objDBAccess Is Nothing Then
                    objDBAccess.DisposeConnection()
                End If
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Function

        ' MODIFICA MARCOG 2011-06-01 - AGGIUNTA GetAnagrafica PER PERMETTERE LOAD ANAGRAFICA
        ' CON COD_CONTRIBUENTE E IDDATAANAGRAFICA
        Public Function GetAnagrafica(ByVal COD_CONTRIBUENTE As Long, ByVal IDDATAANAGRAFICA As Long, _
            ByVal COD_TRIBUTO As String) As DettaglioAnagrafica

            Dim sMyLog As String = "::si inizia"
            Dim strSql As String
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            '===============================================================================
            'WorkFlow
            '===============================================================================

            Dim DettaglioAnagrafica As New DettaglioAnagrafica

            If IDDATAANAGRAFICA = -1 Then
                Return Me.GetAnagrafica(COD_CONTRIBUENTE, COD_TRIBUTO)
            Else
                '===============================================================================
                'MODIFICA
                '===============================================================================
                'WorkFlow
                '===============================================================================
                Try
                    sMyLog += "::devo inizializzare il dbmanager::" & m_IDSottoAttivita
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    sMyLog += "::valorizzo dataset contatti::SELECT * FROM TIPI_CONTATTI ORDER BY IDTipoRiferimento::connessione::" & objDBAccess.GetConnection.ConnectionString
                    DettaglioAnagrafica.dsTipiContatti = AddItemToDataSet(objDBAccess.GetPrivateDataSet("SELECT * FROM TIPI_CONTATTI ORDER BY IDTipoRiferimento"))
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    'm_oSession.oAppDB.DisposeConnection()
                    sMyLog += "::chiudo dbmanager"
                    objDBAccess.DisposeConnection()

                    sMyLog += "::gestione update"
                    'Gestione Contatti
                    '===============================================================================
                    strSql = ""
                    strSql = "SELECT * FROM CONTATTI" & vbCrLf
                    strSql = strSql & "INNER JOIN ANAGRAFICA ON CONTATTI.COD_CONTRIBUENTE = ANAGRAFICA.COD_CONTRIBUENTE AND CONTATTI.IDDATAANAGRAFICA = ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                    strSql = strSql & "WHERE ANAGRAFICA.COD_CONTRIBUENTE =" & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA =" & IDDATAANAGRAFICA & vbCrLf
                    strSql = strSql & "ORDER BY CONTATTI.TIPORIFERIMENTO"
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    DataSetContatti.daContattiPersona = objDBAccess.GetPrivateDataAdapter(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    ' m_oSession.oAppDB.DisposeConnection()


                    Dim ds As dsContatti = DataSetContatti.DataSetCompleto
                    'Caricamento dei contatti associati alla persona se esistenti
                    DettaglioAnagrafica.dsContatti = ds
                    objDBAccess.DisposeConnection()
                    'Caricamento dei tipi contatti 
                    '===============================================================================
                    '===============================================================================
                    'Gestione Anagrafica
                    '===============================================================================
                    '===============================================================================
                    strSql = ""
                    strSql = "SELECT * FROM ANAGRAFICA" & vbCrLf
                    strSql = strSql & "INNER JOIN" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE" & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA =" & IDDATAANAGRAFICA & vbCrLf

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    Dim drDetailsAnagrafica As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    If drDetailsAnagrafica.Read Then
                        'Dati Nascita
                        '********************************************************************************************************************
                        '********************************************************************************************************************
                        DettaglioAnagrafica.COD_CONTRIBUENTE = Utility.GetParametro(drDetailsAnagrafica("COD_CONTRIBUENTE"))
                        DettaglioAnagrafica.ID_DATA_ANAGRAFICA = Utility.CIdFromDB(drDetailsAnagrafica("IDDATAANAGRAFICA"))
                        DettaglioAnagrafica.Cognome = Utility.GetParametro(drDetailsAnagrafica("COGNOME_DENOMINAZIONE"))
                        DettaglioAnagrafica.Nome = Utility.GetParametro(drDetailsAnagrafica("NOME"))
                        DettaglioAnagrafica.CodiceFiscale = Utility.GetParametro(drDetailsAnagrafica("COD_FISCALE"))
                        DettaglioAnagrafica.PartitaIva = Utility.GetParametro(drDetailsAnagrafica("PARTITA_IVA"))
                        DettaglioAnagrafica.CodiceComuneNascita = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_NASCITA"))
                        DettaglioAnagrafica.ComuneNascita = Utility.GetParametro(drDetailsAnagrafica("COMUNE_NASCITA"))
                        DettaglioAnagrafica.ProvinciaNascita = Utility.GetParametro(drDetailsAnagrafica("PROV_NASCITA"))
                        DettaglioAnagrafica.DataNascita = ModDate.GiraDataFromDB(Utility.GetParametro(drDetailsAnagrafica("DATA_NASCITA")))
                        DettaglioAnagrafica.DataMorte = ModDate.GiraDataFromDB(Utility.GetParametro(drDetailsAnagrafica("DATA_MORTE")))
                        DettaglioAnagrafica.NazionalitaNascita = Utility.GetParametro(drDetailsAnagrafica("NAZIONALITA_NASCITA"))
                        DettaglioAnagrafica.Sesso = Utility.GetParametro(drDetailsAnagrafica("SESSO"))

                        '===============================================================================
                        '===============================================================================
                        'Dati Residenza
                        '===============================================================================
                        '===============================================================================
                        DettaglioAnagrafica.CodiceComuneResidenza = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_RES"))
                        DettaglioAnagrafica.ComuneResidenza = Utility.GetParametro(drDetailsAnagrafica("COMUNE_RES"))
                        DettaglioAnagrafica.ProvinciaResidenza = Utility.GetParametro(drDetailsAnagrafica("PROVINCIA_RES"))
                        DettaglioAnagrafica.CapResidenza = Utility.GetParametro(drDetailsAnagrafica("CAP_RES"))
                        DettaglioAnagrafica.CodViaResidenza = Utility.GetParametro(drDetailsAnagrafica("COD_VIA_RES"))
                        DettaglioAnagrafica.ViaResidenza = Utility.GetParametro(drDetailsAnagrafica("VIA_RES"))
                        DettaglioAnagrafica.PosizioneCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("POSIZIONE_CIVICO_RES"))
                        DettaglioAnagrafica.CivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("CIVICO_RES"))
                        DettaglioAnagrafica.EsponenteCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("ESPONENTE_CIVICO_RES"))
                        DettaglioAnagrafica.ScalaCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("SCALA_CIVICO_RES"))
                        DettaglioAnagrafica.InternoCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("INTERNO_CIVICO_RES"))
                        DettaglioAnagrafica.FrazioneResidenza = Utility.GetParametro(drDetailsAnagrafica("FRAZIONE_RES"))
                        DettaglioAnagrafica.NazionalitaResidenza = Utility.GetParametro(drDetailsAnagrafica("NAZIONALITA_RES"))
                        '===============================================================================
                        '===============================================================================
                        'Dati generici
                        '===============================================================================
                        '===============================================================================
                        DettaglioAnagrafica.Professione = Utility.GetParametro(drDetailsAnagrafica("PROFESSIONE"))
                        DettaglioAnagrafica.Note = Utility.GetParametro(drDetailsAnagrafica("NOTE"))
                        DettaglioAnagrafica.DaRicontrollare = Utility.cToBool(drDetailsAnagrafica("DA_RICONTROLLARE"))
                        DettaglioAnagrafica.NucleoFamiliare = Utility.GetParametro(drDetailsAnagrafica("NUCLEO_FAMILIARE"))
                        DettaglioAnagrafica.CodContribuenteRappLegale = Utility.GetParametro(drDetailsAnagrafica("COD_CONTRIBUENTE_RAPP_LEGALE"))
                        DettaglioAnagrafica.Operatore = Utility.GetParametro(drDetailsAnagrafica("OPERATORE"))
                        '===============================================================================
                        '===============================================================================
                        'GESTIONE LOCK
                        '===============================================================================

                        If IsDBNull(drDetailsAnagrafica("CURRENCY")) Then
                            DettaglioAnagrafica.Concurrency = CType(Now, Date)
                        Else
                            DettaglioAnagrafica.Concurrency = drDetailsAnagrafica("CURRENCY")
                        End If
                        DettaglioAnagrafica.CodTributo = COD_TRIBUTO
                        '===============================================================================
                        'GESTIONE LOCK
                        '===============================================================================
                    End If
                    drDetailsAnagrafica.Close()
                    ' m_oSession.oAppDB.DisposeConnection()
                    objDBAccess.DisposeConnection()
                    If Len(DettaglioAnagrafica.CodContribuenteRappLegale) > 0 Then
                        strSql = ""
                        strSql = "SELECT * FROM ANAGRAFICA" & vbCrLf
                        strSql = strSql & "INNER JOIN" & vbCrLf
                        strSql = strSql & "DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE" & vbCrLf
                        strSql = strSql & "AND" & vbCrLf
                        strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
                        strSql = strSql & "WHERE DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE = " & DettaglioAnagrafica.CodContribuenteRappLegale & vbCrLf
                        strSql = strSql & "AND" & vbCrLf
                        strSql = strSql & "(ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')" & vbCrLf
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                        drDetailsAnagrafica = objDBAccess.GetPrivateDataReader(strSql)
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        If drDetailsAnagrafica.Read Then
                            DettaglioAnagrafica.RappresentanteLegale = Utility.GetParametro(drDetailsAnagrafica("COGNOME_DENOMINAZIONE")) & " " & Utility.GetParametro(drDetailsAnagrafica("NOME"))
                        End If
                        drDetailsAnagrafica.Close()
                        'm_oSession.oAppDB.DisposeConnection()
                        objDBAccess.DisposeConnection()
                    End If
                    '===============================================================================
                    'DATI SPEDIZIONE
                    '===============================================================================
                    '===============================================================================
                    strSql = ""
                    strSql = "SELECT * " & vbCrLf
                    strSql = strSql & "FROM INDIRIZZI_SPEDIZIONE INNER JOIN" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_SPEDIZIONE ON INDIRIZZI_SPEDIZIONE.COD_TRIBUTO = DATA_VALIDITA_SPEDIZIONE.COD_TRIBUTO AND " & vbCrLf
                    strSql = strSql & "INDIRIZZI_SPEDIZIONE.COD_CONTRIBUENTE = DATA_VALIDITA_SPEDIZIONE.COD_CONTRIBUENTE AND" & vbCrLf
                    strSql = strSql & "INDIRIZZI_SPEDIZIONE.IDDATA = DATA_VALIDITA_SPEDIZIONE.IDDATA" & vbCrLf
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_SPEDIZIONE.COD_TRIBUTO = " & Utility.CStrToDB(COD_TRIBUTO) & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "DATA_VALIDITA_SPEDIZIONE.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "INDIRIZZI_SPEDIZIONE.IDDATA = " & IDDATAANAGRAFICA & vbCrLf
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    Dim drDetailsSpedizione As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    If drDetailsSpedizione.Read Then

                        DettaglioAnagrafica.ID_DATA_SPEDIZIONE = Utility.GetParametro(drDetailsSpedizione("IDDATA"))
                        DettaglioAnagrafica.CognomeInvio = Utility.GetParametro(drDetailsSpedizione("COGNOME_INVIO"))
                        DettaglioAnagrafica.NomeInvio = Utility.GetParametro(drDetailsSpedizione("NOME_INVIO"))
                        DettaglioAnagrafica.CodComuneRCP = Utility.GetParametro(drDetailsSpedizione("COD_COMUNE_RCP"))
                        DettaglioAnagrafica.ComuneRCP = Utility.GetParametro(drDetailsSpedizione("COMUNE_RCP"))
                        DettaglioAnagrafica.LocRCP = Utility.GetParametro(drDetailsSpedizione("LOC_RCP"))
                        DettaglioAnagrafica.ProvinciaRCP = Utility.GetParametro(drDetailsSpedizione("PROVINCIA_RCP"))
                        DettaglioAnagrafica.CapRCP = Utility.GetParametro(drDetailsSpedizione("CAP_RCP"))
                        DettaglioAnagrafica.CodViaRCP = Utility.GetParametro(drDetailsSpedizione("COD_VIA_RCP"))
                        DettaglioAnagrafica.ViaRCP = Utility.GetParametro(drDetailsSpedizione("VIA_RCP"))
                        DettaglioAnagrafica.PosizioneCivicoRCP = Utility.GetParametro(drDetailsSpedizione("POSIZIONE_CIV_RCP"))
                        DettaglioAnagrafica.CivicoRCP = Utility.GetParametro(drDetailsSpedizione("CIVICO_RCP"))
                        DettaglioAnagrafica.EsponenteCivicoRCP = Utility.GetParametro(drDetailsSpedizione("ESPONENTE_CIVICO_RCP"))
                        DettaglioAnagrafica.ScalaCivicoRCP = Utility.GetParametro(drDetailsSpedizione("SCALA_CIVICO_RCP"))
                        DettaglioAnagrafica.InternoCivicoRCP = Utility.GetParametro(drDetailsSpedizione("INTERNO_CIVICO_RCP"))
                        DettaglioAnagrafica.FrazioneRCP = Utility.GetParametro(drDetailsSpedizione("FRAZIONE_RCP"))

                    Else

                        DettaglioAnagrafica.ID_DATA_SPEDIZIONE = Costant.INIT_VALUE_NUMBER

                    End If
                    drDetailsSpedizione.Close()
                    objDBAccess.DisposeConnection()

                    Return DettaglioAnagrafica
                Catch ex As Exception
                    Throw New Exception("Anagrafica::GetAnagrafica::" & ex.Message & "::SQL::" & strSql & "::MYLOG::" & sMyLog)
                Finally
                    '********************Gestione Anagrafiche massive****************************
                    If Not m_oSession.oAppDB Is Nothing Then
                        m_oSession.oAppDB.DisposeConnection()
                    End If
                    If Not m_oSession.oAppDB Is Nothing Then
                        m_oSession.oAppDB.Dispose()
                    End If
                    If Not objDBAccess Is Nothing Then
                        objDBAccess.DisposeConnection()
                    End If
                    '********************Gestione Anagrafiche massive****************************
                End Try
            End If
        End Function


        '===============================================================================
        'FINE GetAnagrafica
        '===============================================================================
#End Region


#Region "RESTITUISCE  DETTAGLI ANAGRAFICA PER ANAGRAFICHE DOPPIE"
        '=================================================================================
        'Consente di caricare i dati relativi ad una Anagrafica tramite l'utilizzo della Dll che consente di accedere al DataBase tramite
        'WorkFlow
        'PARAMETRI:nessuno
        '==================================================================================

        Public Function GetAnagraficaAnagraficheDoppieOLD(ByVal arrayCodiciContrib() As String, ByVal intTipoRicerca As Integer, ByVal dblPercentuale As Double, ByVal CodEnte As String) As DataSet
            'intTipoRicerca =0 RICERCA PER CODICE FISCALE - PARTITA IVA
            'intTipoRicerca =1 RICERCA PER NOMINATIVO

            Dim strSql As String
            Dim dsAnagraficheDoppie As DataSet
            Dim strCodContribuentiList, sTmp As String
            Dim intCount As Long

            If Not IsNothing(arrayCodiciContrib) Then
                For intCount = 0 To arrayCodiciContrib.Length - 1
                    sTmp = sTmp & arrayCodiciContrib(intCount).ToString & ","
                Next
                If Len(sTmp) > 0 Then
                    '*********************************************************
                    'Pulizia dell'Ultima "Virgola" inserita
                    '*********************************************************
                    strCodContribuentiList = Mid(sTmp, 1, Len(sTmp) - 1)
                End If
            Else
                strCodContribuentiList = ""
            End If

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================
            '===============================================================================
            'WorkFlow
            '===============================================================================

            Try

                '===============================================================================
                '===============================================================================
                'Gestione Anagrafica Anagrafiche Doppie
                '===============================================================================
                '===============================================================================

                Dim dsAnagDoppieDettaglio As New DataSet
                dsAnagDoppieDettaglio.Tables.Add("Anagrafiche")

                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("COD_CONTRIBUENTE")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("COGNOME_DENOMINAZIONE")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("NOME")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("COD_FISCALE")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("PARTITA_IVA")

                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("VIA_RES")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("CIVICO_RES")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("COMUNE_RES")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("CAP_RES")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("PROVINCIA_RES")
                dsAnagDoppieDettaglio.Tables("Anagrafiche").Columns.Add("IDDATAANAGRAFICA")


                strSql = ""

                If intTipoRicerca = 0 Then          'RICERCA PER CODICE FISCALE / PARTITA IVA

                    '===============================================================================
                    '===============================================================================
                    'Prelevo tutte le anagrafiche che hanno il count del CF/PI maggiore di 1

                    strSql = "SELECT SUBSTRING(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END, 1, " & vbCrLf
                    strSql = strSql & " (LEN(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END)*" + CStr(dblPercentuale) + ")/100) AS VALORE,COUNT(*)" & vbCrLf
                    strSql = strSql & " FROM ANAGRAFICA" & vbCrLf
                    strSql = strSql & " WHERE (DATA_FINE_VALIDITA IS NULL OR DATA_FINE_VALIDITA='') " & vbCrLf
                    strSql = strSql & " AND COD_ENTE='" & CodEnte & "'"

                    If strCodContribuentiList.Length > 0 Then

                        strSql = strSql & " AND (COD_CONTRIBUENTE IN (" + strCodContribuentiList + "))" & vbCrLf

                    End If

                    strSql = strSql & " GROUP BY SUBSTRING(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END, 1, (LEN(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END)*" + CStr(dblPercentuale) + ")/100)" & vbCrLf
                    strSql = strSql & " HAVING (COUNT(*) > 1)"
                    strSql = strSql & " ORDER BY VALORE"

                    '===============================================================================
                    '===============================================================================

                ElseIf intTipoRicerca = 1 Then          'RICERCA PER NOMINATIVO

                    '===============================================================================
                    '===============================================================================
                    'Prelevo tutte le anagrafiche che hanno il count del NOMINATIVO maggiore di 1

                    strSql = " select  substring(cognome_denominazione,1,((len(cognome_denominazione )*" + CStr(dblPercentuale) + ")/100)) as VALORE_COGNOME," & vbCrLf
                    strSql = strSql & " substring(nome,1,((len(nome )*" + CStr(dblPercentuale) + ")/100)) as VALORE_NOME, count(*) " & vbCrLf
                    strSql = strSql & " from anagrafica" & vbCrLf
                    strSql = strSql & " where (data_fine_validita is null OR ANAGRAFICA.DATA_FINE_VALIDITA='')" & vbCrLf
                    strSql = strSql & " and cod_ente='" & CodEnte & "'"

                    If strCodContribuentiList.Length > 0 Then

                        strSql = strSql & " AND (COD_CONTRIBUENTE IN (" + strCodContribuentiList + "))" & vbCrLf

                    End If

                    strSql = strSql & " group by substring(cognome_denominazione,1,((len(cognome_denominazione )*" + CStr(dblPercentuale) + ")/100)),substring(nome,1,((len(nome)*" + CStr(dblPercentuale) + ")/100))" & vbCrLf
                    strSql = strSql & " having count(*)>1"
                    strSql = strSql & " order by VALORE_COGNOME, VALORE_NOME"

                    '===============================================================================
                    '===============================================================================

                End If

                '===============================================================================
                'WorkFlow
                '===============================================================================

                Dim drAnagDoppie As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)
                '===============================================================================
                'WorkFlow
                '===============================================================================


                If drAnagDoppie.HasRows = True Then


                    Dim strValore, strValoreCognome, strValoreNome As String
                    Dim intLenght, intLenghtCognome, intLenghtNome As Integer

                    Dim AnagDoppieDettagliorow As DataRow
                    Dim AnagDoppieDettaglioTable As DataTable
                    AnagDoppieDettaglioTable = dsAnagDoppieDettaglio.Tables("Anagrafiche")

                    Do While drAnagDoppie.Read

                        If intTipoRicerca = 0 Then          'RICERCA PER CODICE FISCALE / PARTITA IVA

                            strValore = drAnagDoppie("VALORE")
                            'intLenght = strValore.Length

                            If strValore.Length = 0 Then
                                intLenght = 1
                            Else
                                intLenght = strValore.Length
                            End If

                            strSql = ""
                            strSql = strSql & "select *" & vbCrLf
                            strSql = strSql & " from anagrafica " & vbCrLf
                            strSql = strSql & " where substring(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END,1," + CStr(intLenght) + ")='" + strValore + "' and (data_fine_validita is null OR DATA_FINE_VALIDITA='')"
                            strSql = strSql & " and cod_ente='" & CodEnte & "'"

                        ElseIf intTipoRicerca = 1 Then          'RICERCA PER NOMINATIVO

                            'strValoreCognome = drAnagDoppie("VALORE_COGNOME")
                            'strValoreNome = drAnagDoppie("VALORE_NOME")
                            'intLenghtCognome = strValoreCognome.Length
                            'intLenghtNome = strValoreNome.Length

                            strValoreCognome = drAnagDoppie("VALORE_COGNOME")
                            strValoreCognome = strValoreCognome.Replace("'", "''")
                            strValoreNome = drAnagDoppie("VALORE_NOME")
                            strValoreNome = strValoreNome.Replace("'", "''")
                            If strValoreCognome.Length = 0 Then
                                intLenghtCognome = 1
                            Else
                                intLenghtCognome = strValoreCognome.Length
                            End If

                            If strValoreNome.Length = 0 Then
                                intLenghtNome = 1
                            Else

                                intLenghtNome = strValoreNome.Length
                            End If

                            strSql = ""
                            strSql = strSql & "select *" & vbCrLf
                            strSql = strSql & " from anagrafica" & vbCrLf
                            strSql = strSql & " where substring(cognome_denominazione,1," + CStr(intLenghtCognome) + ")='" + strValoreCognome + "'" & vbCrLf
                            strSql = strSql & " and substring(nome,1," + CStr(intLenghtNome) + ")='" + strValoreNome + "' and (data_fine_validita is null OR ANAGRAFICA.DATA_FINE_VALIDITA='')"
                            strSql = strSql & " and cod_ente='" & CodEnte & "'"

                        End If


                        Dim drAnagDoppieDettaglioTMP As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)

                        Do While drAnagDoppieDettaglioTMP.Read

                            AnagDoppieDettagliorow = AnagDoppieDettaglioTable.NewRow()

                            AnagDoppieDettagliorow.Item("COD_CONTRIBUENTE") = drAnagDoppieDettaglioTMP("COD_CONTRIBUENTE")
                            AnagDoppieDettagliorow.Item("COGNOME_DENOMINAZIONE") = drAnagDoppieDettaglioTMP("COGNOME_DENOMINAZIONE")
                            AnagDoppieDettagliorow.Item("NOME") = drAnagDoppieDettaglioTMP("NOME")
                            AnagDoppieDettagliorow.Item("COD_FISCALE") = drAnagDoppieDettaglioTMP("COD_FISCALE")
                            AnagDoppieDettagliorow.Item("PARTITA_IVA") = drAnagDoppieDettaglioTMP("PARTITA_IVA")
                            AnagDoppieDettagliorow.Item("VIA_RES") = drAnagDoppieDettaglioTMP("VIA_RES")
                            AnagDoppieDettagliorow.Item("CIVICO_RES") = drAnagDoppieDettaglioTMP("CIVICO_RES")
                            AnagDoppieDettagliorow.Item("COMUNE_RES") = drAnagDoppieDettaglioTMP("COMUNE_RES")
                            AnagDoppieDettagliorow.Item("CAP_RES") = drAnagDoppieDettaglioTMP("CAP_RES")
                            AnagDoppieDettagliorow.Item("PROVINCIA_RES") = drAnagDoppieDettaglioTMP("PROVINCIA_RES")
                            AnagDoppieDettagliorow.Item("IDDATAANAGRAFICA") = drAnagDoppieDettaglioTMP("IDDATAANAGRAFICA")

                            AnagDoppieDettaglioTable.Rows.Add(AnagDoppieDettagliorow)

                        Loop

                    Loop

                End If



                '===============================================================================
                'FINE MODIFICA
                '===============================================================================


                Return dsAnagDoppieDettaglio

            Catch ex As Exception
                Throw New Exception("Anagrafica::GetAnagraficaControlloDatiMancanti::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try

        End Function

        Public Function GetAnagraficaAnagraficheDoppie(ByVal intTipoRicerca As Integer, ByVal dblPercentuale As Double, ByVal sNominativo As String, ByVal CodEnte As String) As DataSet
            Dim dsAnagDoppieDettaglio As New DataSet
            Dim sSQL As String
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                If intTipoRicerca = 0 Then
                    sSQL = "SELECT *"
                    sSQL += " FROM ANAGRAFICA"
                    sSQL += " WHERE SUBSTRING(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END,1,"
                    sSQL += " (LEN(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END)*" + dblPercentuale.ToString + ")/100) IN"
                    sSQL += " (SELECT SUBSTRING(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END, 1,"
                    sSQL += " (LEN(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END)*" + dblPercentuale.ToString + ")/100)"
                    sSQL += " FROM ANAGRAFICA"
                    sSQL += " WHERE (DATA_FINE_VALIDITA IS NULL OR DATA_FINE_VALIDITA='')"
                    sSQL += " AND (COD_ENTE='" + CodEnte + "')"
                    sSQL += " GROUP BY SUBSTRING(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END, 1, (LEN(CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END)*" + dblPercentuale.ToString + ")/100)"
                    sSQL += " HAVING (COUNT(*) > 1))"
                    sSQL += " AND (DATA_FINE_VALIDITA IS NULL OR DATA_FINE_VALIDITA='')"
                    sSQL += " AND (COD_ENTE='" + CodEnte + "')"
                    If sNominativo <> "" Then
                        sSQL += " AND (COGNOME_DENOMINAZIONE+' '+NOME LIKE'" & sNominativo & "%')"
                    End If
                    sSQL += " ORDER BY CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END, COGNOME_DENOMINAZIONE, NOME"
                Else
                    sSQL = "SELECT *"
                    sSQL += " FROM ANAGRAFICA"
                    sSQL += " WHERE SUBSTRING(COGNOME_DENOMINAZIONE+' '+NOME,1,((LEN(COGNOME_DENOMINAZIONE+' '+NOME)*" + dblPercentuale.ToString + ")/100)) IN"
                    sSQL += " (SELECT SUBSTRING(COGNOME_DENOMINAZIONE+' '+NOME,1,((LEN(COGNOME_DENOMINAZIONE+' '+NOME)*" + dblPercentuale.ToString + ")/100))"
                    sSQL += " FROM ANAGRAFICA"
                    sSQL += " WHERE (DATA_FINE_VALIDITA IS NULL OR DATA_FINE_VALIDITA='')"
                    sSQL += " AND (COD_ENTE='" + CodEnte + "')"
                    sSQL += " GROUP BY SUBSTRING(COGNOME_DENOMINAZIONE+' '+NOME,1,((LEN(COGNOME_DENOMINAZIONE+' '+NOME)*" + dblPercentuale.ToString + ")/100))"
                    sSQL += " HAVING (COUNT(*) > 1))"
                    sSQL += " AND (DATA_FINE_VALIDITA IS NULL OR DATA_FINE_VALIDITA='')"
                    sSQL += " AND (COD_ENTE='" + CodEnte + "')"
                    If sNominativo <> "" Then
                        sSQL += " AND (COGNOME_DENOMINAZIONE+' '+NOME LIKE'" & sNominativo & "%')"
                    End If
                    sSQL += " ORDER BY COGNOME_DENOMINAZIONE, NOME, CASE WHEN PARTITA_IVA<>'' THEN PARTITA_IVA ELSE COD_FISCALE END"
                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================
                dsAnagDoppieDettaglio = objDBAccess.GetPrivateDataSet(sSQL)

                Return dsAnagDoppieDettaglio

            Catch ex As Exception
                Throw New Exception("Anagrafica::GetAnagraficaControlloDatiMancanti::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Function

        '===============================================================================
        'FINE GetAnagraficaControlloDatiMancanti
        '===============================================================================
#End Region



#Region "RESTITUISCE UN DATASET (DETTAGLIO ANAGRAFICA) PER CONTROLLO DATI MANCANTI"
        '=================================================================================
        'Consente di caricare i dati relativi ad una Anagrafica tramite l'utilizzo della Dll che consente di accedere al DataBase tramite
        'WorkFlow
        'PARAMETRI:nessuno
        '==================================================================================

        Public Function GetAnagraficaControlloDatiMancanti(ByVal CodEnte As String, Optional ByVal SqlXtributo As String = "") As DataSet
            'Public Function GetAnagraficaControlloDatiMancanti() As DettaglioAnagrafica()

            Dim strSql As String

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            '===============================================================================
            'WorkFlow
            '===============================================================================
            'Dim ArrayDettaglioAnagrafica() As DettaglioAnagrafica
            '===============================================================================
            'WorkFlow
            '===============================================================================

            Try

                '===============================================================================
                '===============================================================================
                'Gestione Anagrafica Anagrafiche Dati Mancanti
                '===============================================================================
                '===============================================================================
                '''strSql = ""
                '''strSql = "SELECT DISTINCT *" & vbCrLf
                '''strSql = strSql & " from anagrafica" & vbCrLf
                '''strSql = strSql & " where data_fine_validita is null and (cod_contribuente in (" & vbCrLf
                '''strSql = strSql & " select cod_contribuente from anagrafica" & vbCrLf
                '''strSql = strSql & " where data_fine_validita is null and cognome_denominazione+nome='' " & vbCrLf

                '''strSql = strSql & " union" & vbCrLf

                '''strSql = strSql & " select anagrafica.cod_contribuente" & vbCrLf
                '''strSql = strSql & " from anagrafica left join indirizzi_spedizione on anagrafica.cod_contribuente=indirizzi_spedizione.cod_contribuente " & vbCrLf
                '''strSql = strSql & " where anagrafica.data_fine_validita is null and (via_res='' or cap_res='' or comune_res='' or provincia_res='') and " & vbCrLf
                '''strSql = strSql & " (indirizzi_spedizione.cod_contribuente is null or not indirizzi_spedizione.data_fine_validita is null) " & vbCrLf

                '''strSql = strSql & " union" & vbCrLf

                '''strSql = strSql & " select anagrafica.cod_contribuente" & vbCrLf
                '''strSql = strSql & " from anagrafica INNER JOIN indirizzi_spedizione ON anagrafica.cod_contribuente = indirizzi_spedizione.cod_contribuente" & vbCrLf
                '''strSql = strSql & " WHERE indirizzi_spedizione.data_fine_validita IS NULL AND (via_rcp = '' OR cap_rcp = '' OR comune_rcp = '' OR provincia_rcp = '') " & vbCrLf
                '''strSql = strSql & " AND (via_res = '' OR cap_res = '' OR comune_res = '' OR provincia_res = '')))"

                If SqlXtributo = "" Then
                    strSql = ""
                    strSql += "SELECT * " & _
                              " FROM ANAGRAFICA LEFT OUTER JOIN" & _
                              " INDIRIZZI_SPEDIZIONE ON ANAGRAFICA.COD_CONTRIBUENTE = INDIRIZZI_SPEDIZIONE.COD_CONTRIBUENTE" & _
                              " where cod_ente= '" & CodEnte & "' and anagrafica.data_fine_validita is null and anagrafica.cod_contribuente in (" & _
                              "" + _
                              " select anagrafica.cod_contribuente from anagrafica" & _
                              " where cod_ente= '" & CodEnte & "' and data_fine_validita is null and cognome_denominazione+nome='' " & _
                              " union " & _
                              " select anagrafica.cod_contribuente" & _
                              " from anagrafica left join indirizzi_spedizione on anagrafica.cod_contribuente=indirizzi_spedizione.cod_contribuente " & _
                              " where cod_ente= '" & CodEnte & "' and anagrafica.data_fine_validita is null and (via_res='' or cap_res='' or comune_res='' or provincia_res='') and " & _
                              " (indirizzi_spedizione.cod_contribuente is null or not indirizzi_spedizione.data_fine_validita is null)" & _
                              " union" & _
                              " select anagrafica.cod_contribuente" & _
                              " from anagrafica INNER JOIN indirizzi_spedizione ON anagrafica.cod_contribuente = indirizzi_spedizione.cod_contribuente" & _
                              " WHERE cod_ente= '" & CodEnte & "' and indirizzi_spedizione.data_fine_validita IS NULL AND (via_rcp = '' OR cap_rcp = '' OR comune_rcp = '' OR provincia_rcp = '')" & _
                              " AND (via_res = '' OR cap_res = '' OR comune_res = '' OR provincia_res = '')" & _
                              ") ORDER BY cognome_denominazione, nome"
                Else
                    strSql = ""
                    strSql = SqlXtributo
                End If

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                'Dim drDetailsAnagrafica As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)
                Dim dsDetailsAnagrafica As DataSet = objDBAccess.GetPrivateDataSet(strSql)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Return dsDetailsAnagrafica
                '''''Dim lngCountAnag As Long = 0
                '''''Do While drDetailsAnagrafica.Read

                '''''  'Dati Nascita
                '''''  '********************************************************************************************************************
                '''''  '********************************************************************************************************************

                '''''  Dim oDettaglioAnagrafica As New DettaglioAnagrafica

                '''''  oDettaglioAnagrafica.COD_CONTRIBUENTE = Utility.GetParametro(drDetailsAnagrafica("COD_CONTRIBUENTE"))
                '''''  oDettaglioAnagrafica.ID_DATA_ANAGRAFICA = Utility.CIdFromDB(drDetailsAnagrafica("IDDATAANAGRAFICA"))
                '''''  oDettaglioAnagrafica.Cognome = Utility.GetParametro(drDetailsAnagrafica("COGNOME_DENOMINAZIONE"))
                '''''  oDettaglioAnagrafica.Nome = Utility.GetParametro(drDetailsAnagrafica("NOME"))
                '''''  oDettaglioAnagrafica.CodiceFiscale = Utility.GetParametro(drDetailsAnagrafica("COD_FISCALE"))
                '''''  oDettaglioAnagrafica.PartitaIva = Utility.GetParametro(drDetailsAnagrafica("PARTITA_IVA"))
                '''''  oDettaglioAnagrafica.CodiceComuneNascita = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_NASCITA"))
                '''''  oDettaglioAnagrafica.ComuneNascita = Utility.GetParametro(drDetailsAnagrafica("COMUNE_NASCITA"))
                '''''  oDettaglioAnagrafica.ProvinciaNascita = Utility.GetParametro(drDetailsAnagrafica("PROV_NASCITA"))
                '''''  oDettaglioAnagrafica.DataNascita = ModDate.GiraDataFromDB(Utility.GetParametro(drDetailsAnagrafica("DATA_NASCITA")))
                '''''  oDettaglioAnagrafica.NazionalitaNascita = Utility.GetParametro(drDetailsAnagrafica("NAZIONALITA_NASCITA"))
                '''''  oDettaglioAnagrafica.Sesso = Utility.GetParametro(drDetailsAnagrafica("SESSO"))

                '''''  '===============================================================================
                '''''  '===============================================================================
                '''''  'Dati Residenza
                '''''  '===============================================================================
                '''''  '===============================================================================
                '''''  oDettaglioAnagrafica.CodiceComuneResidenza = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_RES"))
                '''''  oDettaglioAnagrafica.ComuneResidenza = Utility.GetParametro(drDetailsAnagrafica("COMUNE_RES"))
                '''''  oDettaglioAnagrafica.ProvinciaResidenza = Utility.GetParametro(drDetailsAnagrafica("PROVINCIA_RES"))
                '''''  oDettaglioAnagrafica.CapResidenza = Utility.GetParametro(drDetailsAnagrafica("CAP_RES"))
                '''''  oDettaglioAnagrafica.CodViaResidenza = Utility.GetParametro(drDetailsAnagrafica("COD_VIA_RES"))
                '''''  oDettaglioAnagrafica.ViaResidenza = Utility.GetParametro(drDetailsAnagrafica("VIA_RES"))
                '''''  oDettaglioAnagrafica.PosizioneCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("POSIZIONE_CIVICO_RES"))
                '''''  oDettaglioAnagrafica.CivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("CIVICO_RES"))
                '''''  oDettaglioAnagrafica.EsponenteCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("ESPONENTE_CIVICO_RES"))
                '''''  oDettaglioAnagrafica.ScalaCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("SCALA_CIVICO_RES"))
                '''''  oDettaglioAnagrafica.InternoCivicoResidenza = Utility.GetParametro(drDetailsAnagrafica("INTERNO_CIVICO_RES"))
                '''''  oDettaglioAnagrafica.FrazioneResidenza = Utility.GetParametro(drDetailsAnagrafica("FRAZIONE_RES"))
                '''''  oDettaglioAnagrafica.NazionalitaResidenza = Utility.GetParametro(drDetailsAnagrafica("NAZIONALITA_RES"))
                '''''  '===============================================================================
                '''''  '===============================================================================
                '''''  'Dati generici
                '''''  '===============================================================================
                '''''  '===============================================================================
                '''''  oDettaglioAnagrafica.Professione = Utility.GetParametro(drDetailsAnagrafica("PROFESSIONE"))
                '''''  oDettaglioAnagrafica.Note = Utility.GetParametro(drDetailsAnagrafica("NOTE"))
                '''''  oDettaglioAnagrafica.DaRicontrollare = Utility.cToBool(drDetailsAnagrafica("DA_RICONTROLLARE"))
                '''''  oDettaglioAnagrafica.NucleoFamiliare = Utility.GetParametro(drDetailsAnagrafica("NUCLEO_FAMILIARE"))
                '''''  oDettaglioAnagrafica.CodContribuenteRappLegale = Utility.GetParametro(drDetailsAnagrafica("COD_CONTRIBUENTE_RAPP_LEGALE"))
                '''''  oDettaglioAnagrafica.Operatore = Utility.GetParametro(drDetailsAnagrafica("OPERATORE"))

                '''''  ''''===============================================================================
                '''''  ''''DATI SPEDIZIONE
                '''''  ''''===============================================================================
                '''''  ''''===============================================================================

                '''''  '''oDettaglioAnagrafica.ID_DATA_SPEDIZIONE = Utility.GetParametro(drDetailsAnagrafica("IDDATA"))
                '''''  '''oDettaglioAnagrafica.CognomeInvio = Utility.GetParametro(drDetailsAnagrafica("COGNOME_INVIO"))
                '''''  '''oDettaglioAnagrafica.NomeInvio = Utility.GetParametro(drDetailsAnagrafica("NOME_INVIO"))
                '''''  '''oDettaglioAnagrafica.CodComuneRCP = Utility.GetParametro(drDetailsAnagrafica("COD_COMUNE_RCP"))
                '''''  '''oDettaglioAnagrafica.ComuneRCP = Utility.GetParametro(drDetailsAnagrafica("COMUNE_RCP"))
                '''''  '''oDettaglioAnagrafica.LocRCP = Utility.GetParametro(drDetailsAnagrafica("LOC_RCP"))
                '''''  '''oDettaglioAnagrafica.ProvinciaRCP = Utility.GetParametro(drDetailsAnagrafica("PROVINCIA_RCP"))
                '''''  '''oDettaglioAnagrafica.CapRCP = Utility.GetParametro(drDetailsAnagrafica("CAP_RCP"))
                '''''  '''oDettaglioAnagrafica.CodViaRCP = Utility.GetParametro(drDetailsAnagrafica("COD_VIA_RCP"))
                '''''  '''oDettaglioAnagrafica.ViaRCP = Utility.GetParametro(drDetailsAnagrafica("VIA_RCP"))
                '''''  '''oDettaglioAnagrafica.PosizioneCivicoRCP = Utility.GetParametro(drDetailsAnagrafica("POSIZIONE_CIV_RCP"))
                '''''  '''oDettaglioAnagrafica.CivicoRCP = Utility.GetParametro(drDetailsAnagrafica("CIVICO_RCP"))
                '''''  '''oDettaglioAnagrafica.EsponenteCivicoRCP = Utility.GetParametro(drDetailsAnagrafica("ESPONENTE_CIVICO_RCP"))
                '''''  '''oDettaglioAnagrafica.ScalaCivicoRCP = Utility.GetParametro(drDetailsAnagrafica("SCALA_CIVICO_RCP"))
                '''''  '''oDettaglioAnagrafica.InternoCivicoRCP = Utility.GetParametro(drDetailsAnagrafica("INTERNO_CIVICO_RCP"))
                '''''  '''oDettaglioAnagrafica.FrazioneRCP = Utility.GetParametro(drDetailsAnagrafica("FRAZIONE_RCP"))

                '''''  ReDim Preserve ArrayDettaglioAnagrafica(lngCountAnag)
                '''''  ArrayDettaglioAnagrafica(lngCountAnag) = oDettaglioAnagrafica
                '''''  lngCountAnag = lngCountAnag + 1

                '''''Loop
                ''''''===============================================================================
                ''''''FINE MODIFICA
                ''''''===============================================================================

                '''''Return ArrayDettaglioAnagrafica

            Catch ex As Exception
                Throw New Exception("Anagrafica::GetAnagraficaControlloDatiMancanti::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try

        End Function
        '===============================================================================
        'FINE GetAnagraficaControlloDatiMancanti
        '===============================================================================
#End Region




#Region "INSERIMENTO MODIFICA ANAGRAFICA"


        '===============================================================================
        'Gestione Inserimento e Modifica Dati Anagrafici nel DataBase Anagrafica
        '===============================================================================
        Public Function GestisciAnagrafica(ByVal oDettaglioAnagrafica As DettaglioAnagrafica) As DettaglioAnagraficaReturn

            Try

                Dim objDettaglioAnagReturn As New DettaglioAnagraficaReturn
                Dim lngCOD_CONTRIBUENTE As Long
                Dim intRetValDatiVariati As Integer

                '===============================================================================
                'verifico se il codice fiscale/partita iva è presente. se non è presente, creo un CODICE FISCALE VIRTUALE
                'se è presente, verifico che il contribuente relativo sia presente
                '===============================================================================
                If CStr(oDettaglioAnagrafica.CodiceFiscale).CompareTo("") = 0 And CStr(oDettaglioAnagrafica.PartitaIva).CompareTo("") = 0 Then


                    objDettaglioAnagReturn = SetVirtualCF(oDettaglioAnagrafica)
                    oDettaglioAnagrafica.CodiceFiscale = objDettaglioAnagReturn.CODICEFISCALE

                    lngCOD_CONTRIBUENTE = SetAnagraficaIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE)

                    objDettaglioAnagReturn.COD_CONTRIBUENTE = lngCOD_CONTRIBUENTE

                Else

                    Dim dsListaPersone As DataSet
                    If CStr(oDettaglioAnagrafica.CodiceFiscale).CompareTo("") = 0 Then
                        '===============================================================================
                        'codice fiscale vuoto, cerco per partita iva
                        '===============================================================================
                        dsListaPersone = GetListaPersone("", "", "", oDettaglioAnagrafica.PartitaIva, oDettaglioAnagrafica.CodEnte)
                    Else
                        '===============================================================================
                        'codice fiscale pieno, cerco per codice fiscale
                        '===============================================================================
                        dsListaPersone = GetListaPersone("", "", oDettaglioAnagrafica.CodiceFiscale, "", oDettaglioAnagrafica.CodEnte)
                    End If

                    '===============================================================================
                    'se contribuente GIA' presente
                    '===============================================================================
                    If dsListaPersone.Tables(0).Rows.Count > 0 Then
                        '===============================================================================
                        'controllo se i dati sono variati
                        'ControlloDatiVariati ritorna:
                        ' 0 se i dati non sono variati
                        ' 1 se sono variati i dati anagrafici
                        ' 2 se sono variati i dati di spedizione
                        ' 3 se sono variati i dati di anagrafici e spedizione
                        If oDettaglioAnagrafica.COD_CONTRIBUENTE = "-1" Then
                            oDettaglioAnagrafica.COD_CONTRIBUENTE = dsListaPersone.Tables(0).Rows(0)(0)
                            oDettaglioAnagrafica.ID_DATA_ANAGRAFICA = dsListaPersone.Tables(0).Rows(0)(1)
                        End If
                        If oDettaglioAnagrafica.ID_DATA_ANAGRAFICA = "-1" Then
                            oDettaglioAnagrafica.ID_DATA_ANAGRAFICA = dsListaPersone.Tables(0).Rows(0)(1)
                        End If
                        intRetValDatiVariati = ControlloDatiVariati(oDettaglioAnagrafica)
                        '===============================================================================
                        If intRetValDatiVariati = 0 Then
                            '===============================================================================
                            'se non sono variati 
                            '===============================================================================

                            '///Gestione dei Contatti l'Anagrafica non viene storicizzata
                            GestContattiAnagrafica(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA)

                            objDettaglioAnagReturn.COD_CONTRIBUENTE = oDettaglioAnagrafica.COD_CONTRIBUENTE
                            objDettaglioAnagReturn.CODICEFISCALE = oDettaglioAnagrafica.CodiceFiscale

                        ElseIf intRetValDatiVariati = 1 Then
                            '===============================================================================
                            ' sono variati i dati anagrafici --> aggiorno
                            '===============================================================================

                            lngCOD_CONTRIBUENTE = SetAnagraficaIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE)

                            objDettaglioAnagReturn.COD_CONTRIBUENTE = lngCOD_CONTRIBUENTE
                            objDettaglioAnagReturn.CODICEFISCALE = oDettaglioAnagrafica.CodiceFiscale

                        ElseIf intRetValDatiVariati = 2 Then
                            '===============================================================================
                            ' sono variati  i dati di spedizione--> aggiorno
                            '===============================================================================

                            SetIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE)

                            objDettaglioAnagReturn.COD_CONTRIBUENTE = oDettaglioAnagrafica.COD_CONTRIBUENTE
                            objDettaglioAnagReturn.CODICEFISCALE = oDettaglioAnagrafica.CodiceFiscale

                        ElseIf intRetValDatiVariati = 3 Then
                            '===============================================================================
                            ' sono variati  i dati anagrafici e di spedizione--> aggiorno
                            '===============================================================================

                            lngCOD_CONTRIBUENTE = SetAnagraficaIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE, False)

                            objDettaglioAnagReturn.COD_CONTRIBUENTE = lngCOD_CONTRIBUENTE
                            objDettaglioAnagReturn.CODICEFISCALE = oDettaglioAnagrafica.CodiceFiscale

                        End If

                        '===============================================================================
                        'se contribuente NON presente, lo inserisco
                        '===============================================================================
                    ElseIf dsListaPersone.Tables(0).Rows.Count = 0 Then
                        oDettaglioAnagrafica.COD_CONTRIBUENTE = Utility.GetNewId("ANAGRAFICA", m_oSession, m_IDSottoAttivita)
                        oDettaglioAnagrafica.ID_DATA_ANAGRAFICA = Utility.GetNewId("DATA_VALIDITA_ANAGRAFICA", m_oSession, m_IDSottoAttivita)

                        lngCOD_CONTRIBUENTE = SetAnagraficaIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE)

                        objDettaglioAnagReturn.COD_CONTRIBUENTE = lngCOD_CONTRIBUENTE
                        objDettaglioAnagReturn.CODICEFISCALE = oDettaglioAnagrafica.CodiceFiscale

                    End If

                End If

                Return objDettaglioAnagReturn



            Catch ex As Exception

                Throw New Exception("ANAGRAFICA::GestisciAnagrafica::" & ex.Message)

            End Try



        End Function


        Public Function ControlloDatiVariati(ByRef oDettaglioAnagrafica As DettaglioAnagrafica) As Integer
            'ritorna 0 se i dati non sono variati
            'ritorna 1 se sono variati i dati anagrafici
            'ritorna 2 se sono variati i dati di spedizione
            ' 3 se sono variati i dati di anagrafici e spedizione

            Try

                Dim strDatiInzioAnagrafica As String
                Dim strDatiInzioSpedizione As String

                Dim strDatiInzioAnagraficaCompare As String
                Dim strDatiInzioSpedizioneCompare As String

                Dim intRetVal As Integer = 0

                strDatiInzioAnagrafica = ""
                strDatiInzioAnagrafica = oDettaglioAnagrafica.CodiceFiscale
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.PartitaIva
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.Cognome
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.Nome
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.Sesso
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.ComuneNascita
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.ProvinciaNascita
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.DataNascita
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.DataMorte
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.NazionalitaNascita
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.Professione
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.NucleoFamiliare

                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.RappresentanteLegale
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.ComuneResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.CapResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.ProvinciaResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.ViaResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.PosizioneCivicoResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.CivicoResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.EsponenteCivicoResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.ScalaCivicoResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.InternoCivicoResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.FrazioneResidenza
                strDatiInzioAnagrafica = strDatiInzioAnagrafica & oDettaglioAnagrafica.NazionalitaResidenza

                strDatiInzioSpedizione = ""
                strDatiInzioSpedizione = oDettaglioAnagrafica.CognomeInvio
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.NomeInvio
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.ComuneRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.CapRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.ProvinciaRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.ViaRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.PosizioneCivicoRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.CivicoRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.EsponenteCivicoRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.ScalaCivicoRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.InternoCivicoRCP
                strDatiInzioSpedizione = strDatiInzioSpedizione & oDettaglioAnagrafica.FrazioneRCP



                Dim oDettaglioAnagraficaCompare As New DLL.DettaglioAnagrafica

                oDettaglioAnagraficaCompare = GetAnagrafica(oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo)
                If oDettaglioAnagrafica.ID_DATA_SPEDIZIONE = "-1" Or oDettaglioAnagrafica.ID_DATA_SPEDIZIONE = "" Then
                    oDettaglioAnagrafica.ID_DATA_SPEDIZIONE = oDettaglioAnagraficaCompare.ID_DATA_SPEDIZIONE
                End If

                strDatiInzioAnagraficaCompare = ""
                strDatiInzioAnagraficaCompare = oDettaglioAnagraficaCompare.CodiceFiscale
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.PartitaIva
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.Cognome
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.Nome
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.Sesso
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.ComuneNascita
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.ProvinciaNascita
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.DataNascita
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.DataMorte
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.NazionalitaNascita
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.Professione
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.NucleoFamiliare

                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.RappresentanteLegale
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.ComuneResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.CapResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.ProvinciaResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.ViaResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.PosizioneCivicoResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.CivicoResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.EsponenteCivicoResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.ScalaCivicoResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.InternoCivicoResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.FrazioneResidenza
                strDatiInzioAnagraficaCompare = strDatiInzioAnagraficaCompare & oDettaglioAnagraficaCompare.NazionalitaResidenza

                strDatiInzioSpedizioneCompare = ""
                strDatiInzioSpedizioneCompare = oDettaglioAnagraficaCompare.CognomeInvio
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.NomeInvio
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.ComuneRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.CapRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.ProvinciaRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.ViaRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.PosizioneCivicoRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.CivicoRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.EsponenteCivicoRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.ScalaCivicoRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.InternoCivicoRCP
                strDatiInzioSpedizioneCompare = strDatiInzioSpedizioneCompare & oDettaglioAnagraficaCompare.FrazioneRCP

                Dim intCompareString As Integer

                intCompareString = StrComp(strDatiInzioAnagrafica, strDatiInzioAnagraficaCompare, CompareMethod.Text)

                'Se i dati anagrafici caricati dal data base sono uguali a quelli dopo il salvataggio
                If intCompareString = 0 Then

                    intCompareString = StrComp(strDatiInzioSpedizione, strDatiInzioSpedizioneCompare, CompareMethod.Text)

                    'Se i dati di spedizione caricati dal data base sono uguali a quelli dopo il salvataggio
                    If intCompareString = 0 Then
                        '///Gestione dei Contatti l'Anagrafica non viene storicizzata
                        'GestContattiAnagrafica(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA)

                        intRetVal = 0 'dati non variati 

                        'Se i dati di spedizione caricati dal data base NON sono uguali a quelli dopo il salvataggio
                    Else
                        '///Vengono Salvati e storicizzati gli indirizzi
                        'SetIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE)

                        intRetVal = 2 'variati i dati di spedizione

                    End If

                    'Se i dati anagrafici caricati dal data base NON sono uguali a quelli dopo il salvataggio
                Else
                    intCompareString = StrComp(strDatiInzioSpedizione, strDatiInzioSpedizioneCompare, CompareMethod.Text)
                    Dim lngCOD_CONTRIBUENTE As Long
                    If intCompareString = 0 Then
                        '///Vengono Salvati e storicizzati gli indirizzi ed i dati anagrafici
                        'lngCOD_CONTRIBUENTE = SetAnagraficaIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE)

                        intRetVal = 1

                    Else
                        '///Vengono Salvati e storicizzati gli indirizzi ed i dati anagrafici
                        'lngCOD_CONTRIBUENTE = SetAnagraficaIndirizziSpedizione(oDettaglioAnagrafica, oDettaglioAnagrafica.COD_CONTRIBUENTE, oDettaglioAnagrafica.CodTributo, oDettaglioAnagrafica.ID_DATA_ANAGRAFICA, oDettaglioAnagrafica.ID_DATA_SPEDIZIONE, False)

                        intRetVal = 3

                    End If
                End If


                Return intRetVal


            Catch ex As Exception

                Throw New Exception("ANAGRAFICA::ControlloDatiVariati::" & ex.Message & " ::StackTrace::" & ex.StackTrace)

            End Try

        End Function


        Public Function SetVirtualCF(ByVal oDettaglioAnagrafica As DettaglioAnagrafica) As DettaglioAnagraficaReturn

            Dim objCONST As New Costanti
            Dim lngIdentificativoCF As Long
            Dim strIdentificativoCF As String
            Dim strVirtualCF As String
            '===============================================================================
            'WorkFlow
            '===============================================================================
            'Dim objDBAccess As New RIBESFrameWork.DBManager
            'objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)

            '===============================================================================
            'WorkFlow
            '===============================================================================

            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            lngIdentificativoCF = Utility.GetNewId("VIRTUALCF", m_oSession, m_IDSottoAttivita)
            strIdentificativoCF = Utility.NumberToChar(lngIdentificativoCF, "0", 13)
            strVirtualCF = objCONST.VALUE_VIRTUALCF_DEFAULT + strIdentificativoCF

            Dim objDettaglioAnagReturn As New DettaglioAnagraficaReturn

            objDettaglioAnagReturn.COD_CONTRIBUENTE = oDettaglioAnagrafica.COD_CONTRIBUENTE
            objDettaglioAnagReturn.CODICEFISCALE = strVirtualCF

            Return objDettaglioAnagReturn

        End Function




        '===============================================================================
        'Gestione Inserimento e Modifica Dati Anagrafici nel DataBase Anagrafica
        '===============================================================================
        Public Function SetAnagrafica(ByVal oDettaglioAnagrafica As DettaglioAnagrafica) As Long

            Dim strSql As String
            Dim lngTipoOperazione As Long = DBOperation.DB_UPDATE
            Dim lngCodContribuente As Long
            Dim lngIDDataAnagrafica As Long
            Dim lngIDData As Long
            Dim intRetVal As Integer
            Dim objCONST As New Costanti
            '===============================================================================
            'WorkFlow
            '===============================================================================
            'Dim objDBAccess As New RIBESFrameWork.DBManager
            'objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)

            '===============================================================================
            'WorkFlow
            '===============================================================================

            If CInt(oDettaglioAnagrafica.COD_CONTRIBUENTE) = Costant.INIT_VALUE_NUMBER Then lngTipoOperazione = DBOperation.DB_INSERT


            If lngTipoOperazione = DBOperation.DB_INSERT Then

                '===============================================================================
                'WorkFlow
                '===============================================================================
                Dim objDBAccess As New RIBESFrameWork.DBManager
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                lngCodContribuente = Utility.GetNewId("ANAGRAFICA", m_oSession, m_IDSottoAttivita)
                lngIDDataAnagrafica = Utility.GetNewId("DATA_VALIDITA_ANAGRAFICA", m_oSession, m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                strSql = "INSERT INTO ANAGRAFICA" & vbCrLf
                strSql = strSql & "(COD_CONTRIBUENTE,IDDATAANAGRAFICA,COGNOME_DENOMINAZIONE,NOME,COD_FISCALE,PARTITA_IVA," & vbCrLf
                strSql = strSql & "COD_COMUNE_NASCITA,COMUNE_NASCITA,PROV_NASCITA,DATA_NASCITA,DATA_MORTE,NAZIONALITA_NASCITA,SESSO, " & vbCrLf
                strSql = strSql & "COD_COMUNE_RES,COMUNE_RES,PROVINCIA_RES,CAP_RES,COD_VIA_RES,VIA_RES,POSIZIONE_CIVICO_RES, " & vbCrLf
                strSql = strSql & "CIVICO_RES,ESPONENTE_CIVICO_RES,SCALA_CIVICO_RES,INTERNO_CIVICO_RES,FRAZIONE_RES,NAZIONALITA_RES, " & vbCrLf
                strSql = strSql & "PROFESSIONE,NOTE,DA_RICONTROLLARE,OPERATORE,COD_CONTRIBUENTE_RAPP_LEGALE,COD_ENTE,COD_INDIVIDUALE ,NUCLEO_FAMILIARE)" & vbCrLf
                strSql = strSql & "VALUES ( " & vbCrLf
                strSql = strSql & Utility.CIdToDB(lngCodContribuente) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(lngIDDataAnagrafica) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.Cognome) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.Nome) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodiceFiscale) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.PartitaIva) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodiceComuneNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ComuneNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ProvinciaNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(ModDate.GiraData(oDettaglioAnagrafica.DataNascita)) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(ModDate.GiraData(oDettaglioAnagrafica.DataMorte)) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.NazionalitaNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.Sesso) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodiceComuneResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ComuneResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ProvinciaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CapResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(oDettaglioAnagrafica.CodViaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ViaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.PosizioneCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.EsponenteCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ScalaCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.InternoCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.FrazioneResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.NazionalitaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.Professione) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.Note) & "," & vbCrLf
                strSql = strSql & Utility.CToBit(oDettaglioAnagrafica.DaRicontrollare) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.Operatore) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(oDettaglioAnagrafica.CodContribuenteRappLegale) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodEnte) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(oDettaglioAnagrafica.CodIndividuale) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.NucleoFamiliare) & vbCrLf
                strSql = strSql & " )"

                Try

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.BeginTrans()
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    intRetVal = objDBAccess.CmdExec()


                    If intRetVal = objCONST.INIT_VALUE_NUMBER Then

                        Throw New Exception("INSERIMENTO ANAGRAFICA FALLITO")
                    End If

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    'Salvataggio Dei Contatti
                    '===============================================================================
                    If Not IsNothing(oDettaglioAnagrafica.dsContatti) Then

                        Dim dr As DataRow

                        For Each dr In oDettaglioAnagrafica.dsContatti.Tables(0).Rows

                            strSql = ""
                            strSql = "INSERT INTO CONTATTI"
                            strSql = strSql & "(TipoRiferimento,DatiRiferimento,COD_CONTRIBUENTE,IDDATAANAGRAFICA)" & vbCrLf
                            strSql = strSql & "VALUES ( " & vbCrLf
                            strSql = strSql & Utility.CIdToDB(dr("TipoRiferimento")) & "," & vbCrLf
                            strSql = strSql & Utility.CStrToDB(dr("DatiRiferimento")) & "," & vbCrLf
                            strSql = strSql & Utility.CIdToDB(lngCodContribuente) & "," & vbCrLf
                            strSql = strSql & Utility.CIdToDB(lngIDDataAnagrafica) & vbCrLf
                            strSql = strSql & " )"
                            '===============================================================================
                            'WorkFlow
                            '===============================================================================
                            objDBAccess.CmdCreateWithTransaction(strSql)
                            intRetVal = objDBAccess.CmdExec()

                            If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                                Throw New Exception("INSERIMENTO CONTATTI FALLITO")
                            End If

                            '===============================================================================
                            'WorkFlow
                            '===============================================================================


                        Next
                    End If

                    'INDIRZZO DI SPEDIZIONE
                    '===============================================================================
                    If Len(oDettaglioAnagrafica.CognomeInvio) > 0 Then

                        'objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)

                        lngIDData = Utility.GetNewId("DATA_VALIDITA_SPEDIZIONE", m_oSession, m_IDSottoAttivita)

                        strSql = "INSERT INTO INDIRIZZI_SPEDIZIONE" & vbCrLf
                        strSql = strSql & "(COD_TRIBUTO,COD_CONTRIBUENTE,IDDATA,COGNOME_INVIO,NOME_INVIO," & vbCrLf
                        strSql = strSql & "COD_COMUNE_RCP,COMUNE_RCP,LOC_RCP,PROVINCIA_RCP,CAP_RCP, " & vbCrLf
                        strSql = strSql & "COD_VIA_RCP,VIA_RCP,POSIZIONE_CIV_RCP, " & vbCrLf
                        strSql = strSql & "CIVICO_RCP,ESPONENTE_CIVICO_RCP,SCALA_CIVICO_RCP,INTERNO_CIVICO_RCP,FRAZIONE_RCP, OPERATORE)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodTributo) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngCodContribuente) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CognomeInvio) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.NomeInvio) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodComuneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ComuneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.LocRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ProvinciaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CapRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(oDettaglioAnagrafica.CodViaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ViaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.PosizioneCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(oDettaglioAnagrafica.EsponenteCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.ScalaCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.InternoCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.FrazioneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.OperatoreSpedizione) & vbCrLf
                        strSql = strSql & " )"
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()

                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERIMENTO INDIRIZZI SPEDIZIONE FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================


                        strSql = "INSERT INTO DATA_VALIDITA_SPEDIZIONE" & vbCrLf
                        strSql = strSql & "(IDDATA,COD_CONTRIBUENTE,COD_TRIBUTO,DATA_INIZIO_VALIDITA)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngCodContribuente) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioAnagrafica.CodTributo) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                        strSql = strSql & " )"
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()
                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERIMENTO DATA VALIDITA SPEDIZIONE FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                    End If
                    'INDIRZZO DI SPEDIZIONE
                    '===============================================================================
                    'TABELLA DATA_VALIDITA_ANAGRAFICA
                    '===============================================================================
                    strSql = "INSERT INTO DATA_VALIDITA_ANAGRAFICA" & vbCrLf
                    strSql = strSql & "(IDDATAANAGRAFICA,COD_CONTRIBUENTE,DATA_INIZIO_VALIDITA)" & vbCrLf
                    strSql = strSql & "VALUES ( " & vbCrLf
                    strSql = strSql & Utility.CIdToDB(lngIDDataAnagrafica) & "," & vbCrLf
                    strSql = strSql & Utility.CIdToDB(lngCodContribuente) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                    strSql = strSql & " )"
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    intRetVal = objDBAccess.CmdExec()
                    If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                        Throw New Exception("INSERIMENTO DATA_VALIDITA_ANAGRAFICA FALLITO")
                    End If
                    objDBAccess.CommitTrans()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    '//Se tutto e andato bene ritorno il COD_CONTRIBUENTE
                    SetAnagrafica = lngCodContribuente

                Catch ex As Exception
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.RollbackTrans()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    Throw New Exception("ANAGRAFICA::SetAnagrafica::" & ex.Message)


                Finally

                    '********************Gestione Anagrafiche massive****************************
                    objDBAccess.DisposeConnection()
                    objDBAccess.Dispose()
                    '********************Gestione Anagrafiche massive****************************
                End Try

            End If

        End Function

        '===============================================================================
        'Gestione Inserimento e Modifica Dati Anagrafici nel DataBase Anagrafica
        '===============================================================================

#End Region

        Public Sub DeleteAnagrafica(ByVal COD_CONTRIBUENTE As Long, ByVal ID_DATA_ANAGRAFICA As Long)

            '  Dim strSql As String
            '  Dim oConn As New SqlConnection()
            '  Dim oComm As SqlCommand
            '  Dim blnAnagraficaUsata As Boolean = False
            '  Dim drApplicazioni As SqlDataReader
            '  Dim drTabelleDelete As SqlDataReader
            '  Dim drVerificaEsistenza As SqlDataReader
            '  Dim strNomeTabella As String
            '  Dim strNomeApplicazione As String

            '  Dim sqlTrans As SqlTransaction

            '  '===============================================================================
            '  'WorkFlow
            '  '===============================================================================
            '  Dim objDBAccess As New RIBESFrameWork.DBManager()
            '  objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '  '===============================================================================
            '  'WorkFlow
            '  '===============================================================================
            '  Try
            '    strSql = "SELECT * FROM APPLICAZIONI"
            '    '===============================================================================
            '    'WorkFlow
            '    '===============================================================================
            '    drApplicazioni = objDBAccess.GetPrivateDataReader(strSql)
            '    '===============================================================================
            '    'WorkFlow
            '    '===============================================================================
            '    Dim xmlDom As New XmlDocument()

            '    While drApplicazioni.Read
            '      Dim userName, strErrore As String
            '      userName = "GIUSI"
            '      Dim WFSessione As New CreateSessione(drApplicazioni.Item("SUFFISSO_WORKFLOW"), userName, "STARTAPP")
            '      WFSessione.CreaSessione(userName, strErrore)
            '      'WFSessione.oSM.Initialize(userName, drApplicazioni.Item("SUFFISSO_WORKFLOW"))
            '      Dim nApplicazioni As Integer = 0

            '      'Leggo le applicazioni dal WorkFlow
            '      xmlDom = WFSessione.oSM.GetAppList.XMLDOM
            '      nApplicazioni = xmlDom.SelectSingleNode("//NRecord").InnerText
            '      Dim indice As Integer
            '      Dim idApplicazione As String
            '      For indice = 1 To CInt(nApplicazioni)
            '        idApplicazione = ""
            '        idApplicazione = xmlDom.SelectSingleNode("//GetApp[@ID='" & indice & "']/ID_APPLICAZIONE").InnerText
            '        If (StrComp(idApplicazione, "STARTAPP") <> 0) Then

            '          WFSessione = New CreateSessione(drApplicazioni.Item("SUFFISSO_WORKFLOW"), "TEST", idApplicazione)
            '          oConn.ConnectionString = WFSessione.oSession.oAppDB.GetConnection.ConnectionString
            '          oConn.Open()

            '          strSql = "SELECT * FROM TABELLEDELETE WHERE IDAPPLICAZIONE=" & Utility.CIdToDB(drApplicazioni.Item("IDAPPLICAZIONE"))
            '          '===============================================================================
            '          'WorkFlow
            '          '===============================================================================
            '          objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '          drTabelleDelete = objDBAccess.GetPrivateDataReader(strSql)
            '          '===============================================================================
            '          'WorkFlow
            '          '===============================================================================

            '          'Ciclo su tutta la tabella per verificare la presenza di anagrafiche
            '          While drTabelleDelete.Read

            '            strSql = "SELECT * FROM " & Utility.GetParametro(drTabelleDelete.Item("NOMETABELLA")) & " WHERE " & Utility.GetParametro(drTabelleDelete.Item("NOMECAMPO")) & " =" & COD_CONTRIBUENTE

            '            oComm = New SqlCommand(strSql, oConn)

            '            drVerificaEsistenza = oComm.ExecuteReader

            '            If drVerificaEsistenza.Read() Then

            '              strNomeTabella = Utility.GetParametro(drTabelleDelete.Item("NOMETABELLA"))
            '              strNomeApplicazione = Utility.GetParametro(drApplicazioni.Item("NOMEAPPLICAZIONE"))
            '              blnAnagraficaUsata = True
            '              drTabelleDelete.Close()
            '              drApplicazioni.Close()
            '              drVerificaEsistenza.Close()
            '              oComm.Dispose()
            '              oConn.Close()
            '              oConn.Dispose()

            '              Throw New Exception("Impossibile Cancellare l'Anagrafica è associata alla tabella" & strNomeTabella & " Applicazione: " & strNomeApplicazione)

            '            End If

            '            drVerificaEsistenza.Close()

            '          End While

            '          oConn.Close()
            '          oConn.Dispose()
            '        End If
            '      Next
            '    End While

            '    drTabelleDelete.Close()

            '    drApplicazioni.Close()

            '  Catch ex As Exception

            '    Throw

            '  End Try

            '  If Not blnAnagraficaUsata Then

            '    Try

            '      strSql = "DELETE FROM ANAGRAFICA WHERE COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
            '      strSql = strSql & "AND" & vbCrLf
            '      strSql = strSql & "IDDATAANAGRAFICA=" & ID_DATA_ANAGRAFICA
            '      '===============================================================================
            '      'WorkFlow
            '      '===============================================================================
            '      objDBAccess.BeginTrans()
            '      objDBAccess.CmdCreateWithTransaction(strSql)
            '      objDBAccess.CmdExec()
            '      '===============================================================================
            '      'WorkFlow
            '      '===============================================================================

            '      'Cancellazione dati da Indirizzi spedizione

            '      strSql = "DELETE FROM INDIRIZZI_SPEDIZIONE "
            '      strSql = strSql & "WHERE" & vbCrLf
            '      strSql = strSql & "INDIRIZZI_SPEDIZIONE.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
            '      strSql = strSql & "AND" & vbCrLf
            '      strSql = strSql & "(INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA IS NULL OR INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA='')"
            '      '===============================================================================
            '      'WorkFlow
            '      '===============================================================================
            '      objDBAccess.CmdCreateWithTransaction(strSql)
            '      objDBAccess.CmdExec()
            '      '===============================================================================
            '      'WorkFlow
            '      '===============================================================================

            '      objDBAccess.CommitTrans()

            '    Catch ex As Exception
            '      objDBAccess.RollbackTrans()
            '      Throw New Exception("Errore durante la cancellazione di un contribuente dalla tabella ANAGRAFICA")
            '    End Try
            '  End If
            Dim strSql As String
            Dim oConn As New SqlConnection
            Dim oComm As SqlCommand
            Dim blnAnagraficaUsata As Boolean = False
            Dim drApplicazioni As SqlDataReader
            Dim drTabelleDelete As SqlDataReader
            Dim drVerificaEsistenza As SqlDataReader
            Dim strNomeTabella As String
            Dim strNomeApplicazione As String

            Dim sqlTrans As SqlTransaction

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Try
                strSql = "SELECT * FROM APPLICAZIONI"
                '===============================================================================
                'WorkFlow
                '===============================================================================
                drApplicazioni = objDBAccess.GetPrivateDataReader(strSql)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                While drApplicazioni.Read

                    oConn.ConnectionString = Utility.GetParametro(drApplicazioni.Item("STRINGADICONNESSIONE"))
                    oConn.Open()

                    strSql = "SELECT * FROM TABELLEDELETE WHERE IDAPPLICAZIONE=" & Utility.CIdToDB(drApplicazioni.Item("IDAPPLICAZIONE"))
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                    drTabelleDelete = objDBAccess.GetPrivateDataReader(strSql)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    While drTabelleDelete.Read

                        strSql = "SELECT * FROM " & Utility.GetParametro(drTabelleDelete.Item("NOMETABELLA")) & " WHERE " & Utility.GetParametro(drTabelleDelete.Item("NOMECAMPO")) & " =" & COD_CONTRIBUENTE

                        oComm = New SqlCommand(strSql, oConn)

                        drVerificaEsistenza = oComm.ExecuteReader

                        If drVerificaEsistenza.Read() Then

                            strNomeTabella = Utility.GetParametro(drTabelleDelete.Item("NOMETABELLA"))
                            strNomeApplicazione = Utility.GetParametro(drApplicazioni.Item("NOMEAPPLICAZIONE"))
                            blnAnagraficaUsata = True
                            drTabelleDelete.Close()
                            drApplicazioni.Close()
                            drVerificaEsistenza.Close()
                            oComm.Dispose()
                            oConn.Close()
                            oConn.Dispose()

                            '********************Gestione Anagrafiche massive****************************
                            objDBAccess.DisposeConnection()
                            objDBAccess.Dispose()
                            '********************Gestione Anagrafiche massive****************************

                            'Throw New Exception("Impossibile Cancellare l'Anagrafica è associata alla tabella" & strNomeTabella & " Applicazione: " & strNomeApplicazione)
                            Throw New Exception("00000")

                        End If

                        drVerificaEsistenza.Close()

                    End While

                    oConn.Close()
                    oConn.Dispose()

                End While

                drTabelleDelete.Close()

                drApplicazioni.Close()

            Catch ex As Exception
                'Finally
                ''********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                ''********************Gestione Anagrafiche massive****************************
                Throw New Exception("Anagrafica::DeleteAnagrafica::" & ex.Message)
            End Try

            If Not blnAnagraficaUsata Then

                Try

                    strSql = "DELETE FROM ANAGRAFICA WHERE COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "IDDATAANAGRAFICA=" & ID_DATA_ANAGRAFICA
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.BeginTrans()
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    objDBAccess.CmdExec()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    'Cancellazione dati da Indirizzi spedizione

                    strSql = "DELETE FROM INDIRIZZI_SPEDIZIONE "
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "INDIRIZZI_SPEDIZIONE.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "(INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA IS NULL OR INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA='')"
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    objDBAccess.CmdExec()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    objDBAccess.CommitTrans()

                Catch ex As Exception
                    objDBAccess.RollbackTrans()

                    objDBAccess.DisposeConnection()
                    objDBAccess.Dispose()

                    Throw New Exception("Errore durante la cancellazione di un contribuente dalla tabella ANAGRAFICA")
                End Try
            End If

            ''********************Gestione Anagrafiche massive****************************
            objDBAccess.DisposeConnection()
            objDBAccess.Dispose()
            ''********************Gestione Anagrafiche massive****************************

        End Sub

        Public Function SetContatti(ByVal dsTemp As dsContatti, ByVal IDTipoRiferimento As Integer, ByVal DatiRiFerimento As String, ByVal hdIDRIFERIMENTO As Integer) As dsContatti
            Try
                Dim Row As dsContatti.CONTATTIRow

                SetContatti = dsTemp

                If hdIDRIFERIMENTO = Costant.INIT_VALUE_NUMBER Then

                    Row = SetContatti.CONTATTI.NewCONTATTIRow

                    Row.TipoRiferimento = Convert.ToInt32(IDTipoRiferimento)
                    Row.DatiRiferimento = DatiRiFerimento
                    If Row.IDRIFERIMENTO = 0 Then
                        Row.IDRIFERIMENTO = Row.IDRIFERIMENTO + 1
                    End If
                    Row.IDRIFERIMENTO = Row.IDRIFERIMENTO

                    SetContatti.CONTATTI.AddCONTATTIRow(Row)

                Else

                    Dim DBTable As DataTable = SetContatti.Tables("CONTATTI")
                    Dim DBRow As DataRow
                    For Each DBRow In DBTable.Rows
                        If DBRow("IDRIFERIMENTO") = hdIDRIFERIMENTO Then
                            DBRow("TipoRiferimento") = Convert.ToInt32(IDTipoRiferimento)
                            DBRow("DatiRiferimento") = DatiRiFerimento
                            Exit For
                        End If
                    Next


                End If

                SetContatti.AcceptChanges()

                dsTemp = SetContatti



                Return dsTemp
            Catch ex As Exception
                Throw New Exception("Errore durante l'aggiornamento del contatto: funzione-->SetContatti")
            End Try
        End Function

        Public Function DeleteContatti(ByVal dsTemp As dsContatti, ByVal hdIDRIFERIMENTO As Integer) As dsContatti

            Try

                Dim Row As dsContatti.CONTATTIRow
                Dim intRow As Long = 0

                DeleteContatti = dsTemp

                Dim DBTable As DataTable = DeleteContatti.Tables("CONTATTI")
                Dim DBRow As DataRow
                For Each DBRow In DBTable.Rows
                    If DBRow("IDRIFERIMENTO") = hdIDRIFERIMENTO Then
                        DeleteContatti.Tables("CONTATTI").Rows(intRow).Delete()
                        Exit For
                    End If
                    intRow = intRow + 1
                Next


                DeleteContatti.AcceptChanges()

                dsTemp = DeleteContatti



                Return dsTemp
            Catch ex As Exception
                Throw New Exception("Anagrafica::DeleteContatti::" & ex.Message)
            End Try
        End Function


        Public Function DescrizioneTipoContatto(ByVal IDTipoRiferimento As Object) As String

            Dim strSql As String
            Dim rdTemp As SqlDataReader
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                strSql = "SELECT DESCRIZIONE FROM TIPI_CONTATTI WHERE IDTipoRiferimento=" & IDTipoRiferimento
                '===============================================================================
                'WorkFlow
                '===============================================================================
                rdTemp = objDBAccess.GetPrivateDataReader(strSql)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                If rdTemp.Read Then
                    DescrizioneTipoContatto = Utility.GetParametro(rdTemp("DESCRIZIONE"))
                End If
                rdTemp.Close()

                Return DescrizioneTipoContatto
            Catch ex As Exception
                Throw New Exception("Anagrafica::DescrizioneTipoContatto::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try

        End Function

        Private Function AddItemToDataSet(ByRef _dsTemp As DataSet) As DataSet

            Dim dt As DataTable = _dsTemp.Tables(0)
            Dim rowNull As DataRow = dt.NewRow()
            rowNull("Descrizione") = "..."
            rowNull("IDTipoRiferimento") = "-1"
            _dsTemp.Tables(0).Rows.InsertAt(rowNull, 0)
            Return _dsTemp

        End Function


        Public Function SetIndirizziSpedizione(ByVal oDettaglioanagrafica As DettaglioAnagrafica, ByVal COD_CONTRIBUENTE As Int32, ByVal COD_TRIBUTO As String, ByVal IDDATASPEDIZIONE As Int32)

            Dim strSql As String
            Dim lngIDData As Long

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================


            Try
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.BeginTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

                If IDDATASPEDIZIONE <> Costant.INIT_VALUE_NUMBER Then

                    strSql = "UPDATE INDIRIZZI_SPEDIZIONE SET "
                    strSql = strSql & "DATA_FINE_VALIDITA =" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                    strSql = strSql & ",OPERATORE=" & Utility.CStrToDB(oDettaglioanagrafica.Operatore) & vbCrLf
                    strSql = strSql & ",DATA_ULTIMA_MODIFICA=" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "COD_TRIBUTO=" & Utility.CStrToDB(COD_TRIBUTO) & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "IDDATA=" & IDDATASPEDIZIONE & vbCrLf

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    objDBAccess.CmdExec()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                End If
                If Len(oDettaglioanagrafica.CognomeInvio) > 0 Then

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    lngIDData = Utility.GetNewId("DATA_VALIDITA_SPEDIZIONE", m_oSession, m_IDSottoAttivita)
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    strSql = "INSERT INTO INDIRIZZI_SPEDIZIONE" & vbCrLf
                    strSql = strSql & "(COD_TRIBUTO,COD_CONTRIBUENTE,IDDATA,DATA_INIZIO_VALIDITA,COGNOME_INVIO,NOME_INVIO," & vbCrLf
                    strSql = strSql & "COD_COMUNE_RCP,COMUNE_RCP,LOC_RCP,PROVINCIA_RCP,CAP_RCP, " & vbCrLf
                    strSql = strSql & "COD_VIA_RCP,VIA_RCP,POSIZIONE_CIV_RCP, " & vbCrLf
                    strSql = strSql & "CIVICO_RCP,ESPONENTE_CIVICO_RCP,SCALA_CIVICO_RCP,INTERNO_CIVICO_RCP,FRAZIONE_RCP, OPERATORE)" & vbCrLf
                    strSql = strSql & "VALUES ( " & vbCrLf
                    strSql = strSql & Utility.CStrToDB(COD_TRIBUTO) & "," & vbCrLf
                    strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                    strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CognomeInvio) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.NomeInvio) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodComuneRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ComuneRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.LocRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ProvinciaRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CapRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.CodViaRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ViaRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.PosizioneCivicoRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CivicoRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.EsponenteCivicoRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ScalaCivicoRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.InternoCivicoRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.FrazioneRCP) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.OperatoreSpedizione) & vbCrLf
                    strSql = strSql & " )"

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    objDBAccess.CmdExec()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    strSql = "INSERT INTO DATA_VALIDITA_SPEDIZIONE" & vbCrLf
                    strSql = strSql & "(IDDATA,COD_CONTRIBUENTE,COD_TRIBUTO,DATA_INIZIO_VALIDITA)" & vbCrLf
                    strSql = strSql & "VALUES ( " & vbCrLf
                    strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                    strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(COD_TRIBUTO) & "," & vbCrLf
                    strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                    strSql = strSql & " )"

                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    objDBAccess.CmdExec()
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================)


                End If

                strSql = ""
                strSql = "DELETE  FROM CONTATTI  " & vbCrLf
                strSql = strSql & "WHERE" & vbCrLf
                strSql = strSql & "CONTATTI.COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                strSql = strSql & "AND" & vbCrLf
                strSql = strSql & "CONTATTI.IDDATAANAGRAFICA=" & oDettaglioanagrafica.ID_DATA_ANAGRAFICA & vbCrLf

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.CmdCreateWithTransaction(strSql)
                objDBAccess.CmdExec()
                '===============================================================================
                'WorkFlow
                '===============================================================================
                If Not IsNothing(oDettaglioanagrafica.dsContatti) Then

                    Dim dr As DataRow

                    For Each dr In oDettaglioanagrafica.dsContatti.Tables(0).Rows

                        strSql = ""
                        strSql = "INSERT INTO CONTATTI"
                        strSql = strSql & "(TipoRiferimento,DatiRiferimento,COD_CONTRIBUENTE,IDDATAANAGRAFICA)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CIdToDB(dr("TipoRiferimento")) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(dr("DatiRiferimento")) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.ID_DATA_ANAGRAFICA) & vbCrLf
                        strSql = strSql & " )"

                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        objDBAccess.CmdExec()
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                    Next

                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.CommitTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

            Catch ex As Exception
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.RollbackTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Throw New Exception("Anagrafica::SetIndirizziSpedizione" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try


        End Function
        '/====================================================
        '//Funzione di salvataggio Dati in Anagrafica e tabella Indirizzi di Spedizione
        '//==================================================== 
        Public Function SetAnagraficaIndirizziSpedizione(ByVal oDettaglioanagrafica As DettaglioAnagrafica, ByVal COD_CONTRIBUENTE As Long, _
        ByVal COD_TRIBUTO As String, ByVal IDDATAANAGRAFICA As Long, ByVal IDDATASPEDIZIONE As Long, _
        Optional ByVal blnDatiInizioUgualiDatiFineSpedizione As Boolean = True) As Long

            Dim strSql As String
            Dim lngIDData As Long
            Dim lngCodContribuente As Long
            Dim lngIDDataAnagrafica As Long
            Dim intRetVal As Integer
            Dim objCONST As New Costanti
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================

            Try

                '===============================================================================
                'WorkFlow
                '===============================================================================
                lngIDDataAnagrafica = Utility.GetNewId("DATA_VALIDITA_ANAGRAFICA", m_oSession, m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                strSql = "INSERT INTO ANAGRAFICA" & vbCrLf
                strSql = strSql & "(COD_CONTRIBUENTE,IDDATAANAGRAFICA,COGNOME_DENOMINAZIONE,NOME,COD_FISCALE,PARTITA_IVA," & vbCrLf
                strSql = strSql & "COD_COMUNE_NASCITA,COMUNE_NASCITA,PROV_NASCITA,DATA_NASCITA,DATA_MORTE,NAZIONALITA_NASCITA,SESSO, " & vbCrLf
                strSql = strSql & "COD_COMUNE_RES,COMUNE_RES,PROVINCIA_RES,CAP_RES,COD_VIA_RES,VIA_RES,POSIZIONE_CIVICO_RES, " & vbCrLf
                strSql = strSql & "CIVICO_RES,ESPONENTE_CIVICO_RES,SCALA_CIVICO_RES,INTERNO_CIVICO_RES,FRAZIONE_RES,NAZIONALITA_RES, " & vbCrLf
                strSql = strSql & "PROFESSIONE,NOTE,DA_RICONTROLLARE,OPERATORE,COD_CONTRIBUENTE_RAPP_LEGALE,COD_ENTE,COD_INDIVIDUALE,NUCLEO_FAMILIARE )" & vbCrLf
                strSql = strSql & "VALUES ( " & vbCrLf
                strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(lngIDDataAnagrafica) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.Cognome) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.Nome) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodiceFiscale) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.PartitaIva) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodiceComuneNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ComuneNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ProvinciaNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(ModDate.GiraData(oDettaglioanagrafica.DataNascita)) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(ModDate.GiraData(oDettaglioanagrafica.DataMorte)) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.NazionalitaNascita) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.Sesso) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodiceComuneResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ComuneResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ProvinciaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CapResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.CodViaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ViaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.PosizioneCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.EsponenteCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ScalaCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.InternoCivicoResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.FrazioneResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.NazionalitaResidenza) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.Professione) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.Note) & "," & vbCrLf
                strSql = strSql & Utility.CToBit(oDettaglioanagrafica.DaRicontrollare) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.Operatore) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.CodContribuenteRappLegale) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodEnte) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.CodIndividuale) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.NucleoFamiliare) & vbCrLf
                strSql = strSql & " )"
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.BeginTrans()
                objDBAccess.CmdCreateWithTransaction(strSql)
                intRetVal = objDBAccess.CmdExec()
                If intRetVal = objCONST.INIT_VALUE_NUMBER Then

                    Throw New Exception("INSERIMENTO ANAGRAFICA FALLITO")
                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================

                If Not IsNothing(oDettaglioanagrafica.dsContatti) Then

                    Dim dr As DataRow

                    For Each dr In oDettaglioanagrafica.dsContatti.Tables(0).Rows

                        strSql = ""
                        strSql = "INSERT INTO CONTATTI"
                        strSql = strSql & "(TipoRiferimento,DatiRiferimento,COD_CONTRIBUENTE,IDDATAANAGRAFICA)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CIdToDB(dr("TipoRiferimento")) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(dr("DatiRiferimento")) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDDataAnagrafica) & vbCrLf
                        strSql = strSql & " )"

                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()
                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERIMENTO CONTATTI FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================


                    Next

                End If
                '****************************************************************
                strSql = "INSERT INTO DATA_VALIDITA_ANAGRAFICA" & vbCrLf
                strSql = strSql & "(IDDATAANAGRAFICA,COD_CONTRIBUENTE,DATA_INIZIO_VALIDITA)" & vbCrLf
                strSql = strSql & "VALUES ( " & vbCrLf
                strSql = strSql & Utility.CIdToDB(lngIDDataAnagrafica) & "," & vbCrLf
                strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                strSql = strSql & " )"

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.CmdCreateWithTransaction(strSql)
                intRetVal = objDBAccess.CmdExec()

                If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                    Throw New Exception("INSERIMENTO DATA VALIDITA ANAGRAFICA FALLITO")
                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================


                strSql = "UPDATE ANAGRAFICA SET "
                strSql = strSql & "DATA_FINE_VALIDITA =" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                strSql = strSql & ",OPERATORE=" & Utility.CStrToDB(oDettaglioanagrafica.Operatore) & vbCrLf
                strSql = strSql & ",DATA_ULTIMA_MODIFICA=" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                strSql = strSql & ",IDCODCONTRIBUENTEFIGLIO=" & COD_CONTRIBUENTE & vbCrLf
                strSql = strSql & ",IDDATAANAGRAFICAFIGLIO=" & lngIDDataAnagrafica & vbCrLf
                strSql = strSql & "WHERE" & vbCrLf
                strSql = strSql & "COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                strSql = strSql & "AND" & vbCrLf
                strSql = strSql & "IDDATAANAGRAFICA=" & IDDATAANAGRAFICA & vbCrLf

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.CmdCreateWithTransaction(strSql)
                intRetVal = objDBAccess.CmdExec()
                If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                    Throw New Exception("UPDATE ANAGRAFICA FALLITO")
                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================



                If blnDatiInizioUgualiDatiFineSpedizione Then

                    If Len(oDettaglioanagrafica.CognomeInvio) > 0 Then

                        strSql = "UPDATE INDIRIZZI_SPEDIZIONE SET "
                        strSql = strSql & "DATA_FINE_VALIDITA =" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                        strSql = strSql & ",OPERATORE=" & Utility.CStrToDB(oDettaglioanagrafica.Operatore) & vbCrLf
                        strSql = strSql & ",DATA_ULTIMA_MODIFICA=" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                        strSql = strSql & "WHERE" & vbCrLf
                        strSql = strSql & "COD_TRIBUTO=" & Utility.CStrToDB(COD_TRIBUTO) & vbCrLf
                        strSql = strSql & "AND" & vbCrLf
                        strSql = strSql & "COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                        strSql = strSql & "AND" & vbCrLf
                        strSql = strSql & "IDDATA=" & IDDATASPEDIZIONE & vbCrLf

                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()

                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("UPDATE INDIRIZZI SPEDIZIONE FALLITO")
                        End If

                        lngIDData = Utility.GetNewId("DATA_VALIDITA_SPEDIZIONE", m_oSession, m_IDSottoAttivita)
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                        strSql = "INSERT INTO INDIRIZZI_SPEDIZIONE" & vbCrLf
                        strSql = strSql & "(COD_TRIBUTO,COD_CONTRIBUENTE,IDDATA,COGNOME_INVIO,NOME_INVIO," & vbCrLf
                        strSql = strSql & "COD_COMUNE_RCP,COMUNE_RCP,LOC_RCP,PROVINCIA_RCP,CAP_RCP, " & vbCrLf
                        strSql = strSql & "COD_VIA_RCP,VIA_RCP,POSIZIONE_CIV_RCP, " & vbCrLf
                        strSql = strSql & "CIVICO_RCP,ESPONENTE_CIVICO_RCP,SCALA_CIVICO_RCP,INTERNO_CIVICO_RCP,FRAZIONE_RCP, OPERATORE)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CStrToDB(COD_TRIBUTO) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CognomeInvio) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.NomeInvio) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodComuneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ComuneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.LocRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ProvinciaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CapRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.CodViaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ViaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.PosizioneCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.EsponenteCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ScalaCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.InternoCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.FrazioneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.OperatoreSpedizione) & vbCrLf
                        strSql = strSql & " )"
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)

                        intRetVal = objDBAccess.CmdExec()
                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERT INDIRIZZI SPEDIZIONE FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                        strSql = "INSERT INTO DATA_VALIDITA_SPEDIZIONE" & vbCrLf
                        strSql = strSql & "(IDDATA,COD_CONTRIBUENTE,COD_TRIBUTO,DATA_INIZIO_VALIDITA)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(COD_TRIBUTO) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                        strSql = strSql & " )"
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()

                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERT DATA VALIDITA SPEDIZIONE  FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                    End If
                End If
                If Not blnDatiInizioUgualiDatiFineSpedizione Then
                    '**************************************************************
                    'I campi sono stati ripuliti ma al caricamento erano presenti

                    strSql = "UPDATE INDIRIZZI_SPEDIZIONE SET "
                    strSql = strSql & "DATA_FINE_VALIDITA =" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                    strSql = strSql & ",OPERATORE=" & Utility.CStrToDB(oDettaglioanagrafica.Operatore) & vbCrLf
                    strSql = strSql & ",DATA_ULTIMA_MODIFICA=" & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                    strSql = strSql & "WHERE" & vbCrLf
                    strSql = strSql & "COD_TRIBUTO=" & Utility.CStrToDB(COD_TRIBUTO) & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                    strSql = strSql & "AND" & vbCrLf
                    strSql = strSql & "IDDATA=" & IDDATASPEDIZIONE & vbCrLf
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================
                    objDBAccess.CmdCreateWithTransaction(strSql)
                    intRetVal = objDBAccess.CmdExec()

                    If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                        Throw New Exception("UPDATE INDIRIZZI SPEDIZIONE FALLITO")
                    End If
                    '===============================================================================
                    'WorkFlow
                    '===============================================================================

                    'Dati sono stati semplicemente modifcati
                    If Len(oDettaglioanagrafica.CognomeInvio) > 0 Then
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        lngIDData = Utility.GetNewId("DATA_VALIDITA_SPEDIZIONE", m_oSession, m_IDSottoAttivita)
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                        strSql = "INSERT INTO INDIRIZZI_SPEDIZIONE" & vbCrLf
                        strSql = strSql & "(COD_TRIBUTO,COD_CONTRIBUENTE,IDDATA,COGNOME_INVIO,NOME_INVIO," & vbCrLf
                        strSql = strSql & "COD_COMUNE_RCP,COMUNE_RCP,LOC_RCP,PROVINCIA_RCP,CAP_RCP, " & vbCrLf
                        strSql = strSql & "COD_VIA_RCP,VIA_RCP,POSIZIONE_CIV_RCP, " & vbCrLf
                        strSql = strSql & "CIVICO_RCP,ESPONENTE_CIVICO_RCP,SCALA_CIVICO_RCP,INTERNO_CIVICO_RCP,FRAZIONE_RCP, OPERATORE)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CStrToDB(COD_TRIBUTO) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CognomeInvio) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.NomeInvio) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CodComuneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ComuneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.LocRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ProvinciaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CapRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(oDettaglioanagrafica.CodViaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ViaRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.PosizioneCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.CivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.EsponenteCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.ScalaCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.InternoCivicoRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.FrazioneRCP) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(oDettaglioanagrafica.OperatoreSpedizione) & vbCrLf
                        strSql = strSql & " )"

                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()
                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERT INDIRIZZI SPEDIZIONE FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                        strSql = "INSERT INTO DATA_VALIDITA_SPEDIZIONE" & vbCrLf
                        strSql = strSql & "(IDDATA,COD_CONTRIBUENTE,COD_TRIBUTO,DATA_INIZIO_VALIDITA)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CIdToDB(lngIDData) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(COD_TRIBUTO) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(DateTime.Now.ToString("yyyyMMdd")) & vbCrLf
                        strSql = strSql & " )"

                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()
                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERT DATA VALIDITA SPEDIZIONE FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                    End If
                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.CommitTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

                '//SE NON SI è VARIFICATO ALCUN ERROR RITORNO IL CODICE_CONTRIBUENTE APPENA INSERITO
                SetAnagraficaIndirizziSpedizione = COD_CONTRIBUENTE

            Catch ex As Exception

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.RollbackTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Throw New Exception("Anagrafica::SetAnagraficaIndirizziSpedizione::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try


        End Function

        Public Sub GestContattiAnagrafica(ByVal oDettaglioanagrafica As DettaglioAnagrafica, ByVal COD_CONTRIBUENTE As Int32, ByVal IDDATAANAGRAFICA As Int32)

            Dim strSql As String
            Dim lngIDDataAnagrafica As Long
            Dim intRetVal As Integer
            Dim objCONST As New Costanti
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================

            strSql = ""
            strSql = "DELETE  FROM CONTATTI  " & vbCrLf
            strSql = strSql & "WHERE" & vbCrLf
            strSql = strSql & "CONTATTI.COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
            strSql = strSql & "AND" & vbCrLf
            strSql = strSql & "CONTATTI.IDDATAANAGRAFICA=" & IDDATAANAGRAFICA & vbCrLf
            Try

                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.BeginTrans()
                objDBAccess.CmdCreateWithTransaction(strSql)
                intRetVal = objDBAccess.CmdExec()
                If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                    Throw New Exception("DELETE CONTATTI FALLITO")
                End If
                '===============================================================================
                'WorkFlow
                '===============================================================================

                If Not IsNothing(oDettaglioanagrafica.dsContatti) Then

                    Dim dr As DataRow

                    For Each dr In oDettaglioanagrafica.dsContatti.Tables(0).Rows

                        strSql = ""
                        strSql = "INSERT INTO CONTATTI"
                        strSql = strSql & "(TipoRiferimento,DatiRiferimento,COD_CONTRIBUENTE,IDDATAANAGRAFICA)" & vbCrLf
                        strSql = strSql & "VALUES ( " & vbCrLf
                        strSql = strSql & Utility.CIdToDB(dr("TipoRiferimento")) & "," & vbCrLf
                        strSql = strSql & Utility.CStrToDB(dr("DatiRiferimento")) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(COD_CONTRIBUENTE) & "," & vbCrLf
                        strSql = strSql & Utility.CIdToDB(IDDATAANAGRAFICA) & vbCrLf
                        strSql = strSql & " )"
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================
                        objDBAccess.CmdCreateWithTransaction(strSql)
                        intRetVal = objDBAccess.CmdExec()
                        If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                            Throw New Exception("INSERT CONTATTI FALLITO")
                        End If
                        '===============================================================================
                        'WorkFlow
                        '===============================================================================

                    Next

                End If

                strSql = "UPDATE ANAGRAFICA SET "
                strSql = strSql & "NOTE =" & Utility.CStrToDB(oDettaglioanagrafica.Note) & vbCrLf
                strSql = strSql & ",DA_RICONTROLLARE=" & Utility.CToBit(oDettaglioanagrafica.DaRicontrollare) & vbCrLf
                strSql = strSql & "WHERE" & vbCrLf
                strSql = strSql & "COD_CONTRIBUENTE=" & COD_CONTRIBUENTE & vbCrLf
                strSql = strSql & "AND" & vbCrLf
                strSql = strSql & "IDDATAANAGRAFICA=" & IDDATAANAGRAFICA & vbCrLf
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.CmdCreateWithTransaction(strSql)

                intRetVal = objDBAccess.CmdExec()
                If intRetVal = objCONST.INIT_VALUE_NUMBER Then
                    Throw New Exception("UPDATE ANAGRAFICA")
                End If
                objDBAccess.CommitTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

            Catch ex As Exception
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objDBAccess.RollbackTrans()
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Throw New Exception("Anagrafica::GestContattiAnagrafica::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try

        End Sub



        '*****************************************************************************************
        '////Parametri:Cognome,Nome,CodiceFiscale,PartitaIva
        '////La funzione serve per esporre i dati all'interno di un dataset 
        '////Da usarsi per esempio in una pagina di ricerca dove è necessario visualizzare i dati in una tabella o griglia
        '////Deve essere usata per estrapolare le Anagrafiche attive
        '******************************************************************************************
        Public Function GetListaPersone(ByVal Cognome As String, ByVal Nome As String, ByVal CodiceFiscale As String, ByVal PartitaIVA As String) As DataSet


            Dim StringaCerca As String

            StringaCerca = " SELECT ANAGRAFICA.COD_CONTRIBUENTE AS CODICE,ANAGRAFICA.IDDATAANAGRAFICA ,ANAGRAFICA.COGNOME_DENOMINAZIONE , ANAGRAFICA.NOME,"
            StringaCerca = StringaCerca & " DN = CASE WHEN DATA_NASCITA <>'' THEN RIGHT(DATA_NASCITA,2) +'/' + RIGHT(LEFT(DATA_NASCITA,6),2) +'/' +LEFT(DATA_NASCITA,4) ELSE '' END,"
            StringaCerca = StringaCerca & " CFPI = CASE WHEN SESSO <>'G' THEN COD_FISCALE ELSE PARTITA_IVA END FROM ANAGRAFICA"
            StringaCerca = StringaCerca & " INNER JOIN"
            StringaCerca = StringaCerca & " DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE"
            StringaCerca = StringaCerca & " AND"
            StringaCerca = StringaCerca & " ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA"
            StringaCerca = StringaCerca & " AND"
            StringaCerca = StringaCerca & " (ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')"
            StringaCerca = StringaCerca & " WHERE 1=1 "

            If Trim(Cognome) <> "" Then
                StringaCerca = StringaCerca & " AND (COGNOME_DENOMINAZIONE LIKE '" & Replace(Replace(Trim(Cognome), "'", "''"), "*", "%") & "%')"
            End If
            If Trim(Nome) <> "" Then
                StringaCerca = StringaCerca & " AND (NOME LIKE '" & Replace(Replace(Trim(Nome), "'", "''"), "*", "%") & "%')"
            End If
            If Trim(CodiceFiscale) <> "" Then
                StringaCerca = StringaCerca & " AND (COD_FISCALE LIKE '" & Replace(Trim(CodiceFiscale), "*", "%") & "%')"
            End If
            If Trim(PartitaIVA) <> "" Then
                StringaCerca = StringaCerca & " AND (PARTITA_IVA LIKE '" & Replace(Trim(PartitaIVA), "*", "%") & "%')"
            End If
            StringaCerca = StringaCerca & " ORDER BY ANAGRAFICA.COGNOME_DENOMINAZIONE, ANAGRAFICA.NOME"
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================
                GetListaPersone = objDBAccess.GetPrivateDataSet(StringaCerca)

                Return GetListaPersone
            Catch ex As Exception
                Throw New Exception("Anagrafica::GetListaPersone::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Function

        Public Function GetListaPersone(ByVal Cognome As String, ByVal Nome As String, ByVal CodiceFiscale As String, ByVal PartitaIVA As String, ByVal CodEnte As String) As DataSet


            Dim StringaCerca As String

            StringaCerca = " SELECT ANAGRAFICA.COD_CONTRIBUENTE AS CODICE,ANAGRAFICA.IDDATAANAGRAFICA ,ANAGRAFICA.COGNOME_DENOMINAZIONE , ANAGRAFICA.NOME,"
            StringaCerca = StringaCerca & " DN = CASE WHEN DATA_NASCITA <>'' THEN RIGHT(DATA_NASCITA,2) +'/' + RIGHT(LEFT(DATA_NASCITA,6),2) +'/' +LEFT(DATA_NASCITA,4) ELSE '' END,"
            StringaCerca = StringaCerca & " CFPI = CASE WHEN SESSO <>'G' THEN COD_FISCALE ELSE PARTITA_IVA END FROM ANAGRAFICA"
            StringaCerca = StringaCerca & " INNER JOIN"
            StringaCerca = StringaCerca & " DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE"
            StringaCerca = StringaCerca & " AND"
            StringaCerca = StringaCerca & " ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA"
            StringaCerca = StringaCerca & " AND"
            StringaCerca = StringaCerca & " (ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')"
            StringaCerca = StringaCerca & " WHERE 1=1 "

            If Trim(Cognome) <> "" Then
                StringaCerca = StringaCerca & " AND (COGNOME_DENOMINAZIONE LIKE '" & Replace(Replace(Trim(Cognome), "'", "''"), "*", "%") & "%')"
            End If
            If Trim(Nome) <> "" Then
                StringaCerca = StringaCerca & " AND (NOME LIKE '" & Replace(Replace(Trim(Nome), "'", "''"), "*", "%") & "%')"
            End If
            If Trim(CodiceFiscale) <> "" Then
                StringaCerca = StringaCerca & " AND (COD_FISCALE LIKE '" & Replace(Trim(CodiceFiscale), "*", "%") & "%')"
            End If
            If Trim(PartitaIVA) <> "" Then
                StringaCerca = StringaCerca & " AND (PARTITA_IVA LIKE '" & Replace(Trim(PartitaIVA), "*", "%") & "%')"
            End If
            If Trim(CodEnte) <> "" Then
                StringaCerca = StringaCerca & " AND (COD_ENTE = '" & Trim(CodEnte) & "')"
            End If
            StringaCerca = StringaCerca & " ORDER BY ANAGRAFICA.COGNOME_DENOMINAZIONE, ANAGRAFICA.NOME"
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================
                GetListaPersone = objDBAccess.GetPrivateDataSet(StringaCerca)

                Return GetListaPersone
            Catch ex As Exception
                Throw New Exception("Anagrafica::GetListaPersone::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Function
        '*****************************************************************************************
        '////Parametri:Cognome,Nome,CodiceFiscale,PartitaIva
        '////La funzione serve per esporre i dati all'interno di un dataset 
        '////Da usarsi per esempio in una pagina di ricerca dove è necessario visualizzare i dati in una tabella o griglia
        '////Deve essere usata per estrapolare le Anagrafiche stroricizzate dal sistema
        '******************************************************************************************
        Public Function GetListaPersoneStorico(ByVal Cognome As String, _
        ByVal Nome As String, _
        ByVal CodiceFiscale As String, ByVal PartitaIVA As String) As DataSet

            Dim StringaCerca As String

            StringaCerca = " SELECT ANAGRAFICA.COD_CONTRIBUENTE AS CODICE,ANAGRAFICA.IDDATAANAGRAFICA ,ANAGRAFICA.COGNOME_DENOMINAZIONE , ANAGRAFICA.NOME,"
            StringaCerca = StringaCerca & " DN = CASE WHEN DATA_NASCITA <>'' THEN RIGHT(DATA_NASCITA,2) +'/' + RIGHT(LEFT(DATA_NASCITA,6),2) +'/' +LEFT(DATA_NASCITA,4) ELSE '' END,"
            StringaCerca = StringaCerca & " CFPI = CASE WHEN SESSO <>'G' THEN COD_FISCALE ELSE PARTITA_IVA END FROM ANAGRAFICA"
            StringaCerca = StringaCerca & " INNER JOIN"
            StringaCerca = StringaCerca & " DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE"
            StringaCerca = StringaCerca & " AND"
            StringaCerca = StringaCerca & " ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA"
            StringaCerca = StringaCerca & " AND"
            StringaCerca = StringaCerca & " (ANAGRAFICA.DATA_FINE_VALIDITA IS NOT  NULL OR ANAGRAFICA.DATA_FINE_VALIDITA<>'')"
            StringaCerca = StringaCerca & " WHERE 1=1 "

            If Trim(Cognome) <> "" Then
                StringaCerca = StringaCerca & " AND (COGNOME_DENOMINAZIONE LIKE '" & Replace(Replace(Trim(Cognome), "'", "''"), "*", "%") & "%')"
            End If
            If Trim(Nome) <> "" Then
                StringaCerca = StringaCerca & " AND (NOME LIKE '" & Replace(Replace(Trim(Nome), "'", "''"), "*", "%") & "%')"
            End If
            If Trim(CodiceFiscale) <> "" Then
                StringaCerca = StringaCerca & " AND (COD_FISCALE LIKE '" & Replace(Trim(CodiceFiscale), "*", "%") & "%')"
            End If
            If Trim(PartitaIVA) <> "" Then
                StringaCerca = StringaCerca & " AND (PARTITA_IVA LIKE '" & Replace(Trim(PartitaIVA), "*", "%") & "%')"
            End If
            StringaCerca = StringaCerca & " ORDER BY ANAGRAFICA.COGNOME_DENOMINAZIONE, ANAGRAFICA.NOME"
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                GetListaPersoneStorico = objDBAccess.GetPrivateDataSet(StringaCerca)

                Return GetListaPersoneStorico
            Catch ex As Exception
                Throw New Exception("Anagrafica::GetListaPersoneStorico::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Function
        '*****************************************************************************************
        '////Parametri:Cognome,Nome,CodiceFiscale,PartitaIva
        '////Da usarsi per esempio in una pagina di ricerca dove è necessario visualizzare i dati in una tabella o griglia
        '////Usa una Stroed Procedure per caricare un oggetto DataGrid di tipo RibesDataGrid (ASP.Net)
        '////La Griglia sarà caricata con le anagrafiche attive presenti a sitema 
        '******************************************************************************************
        Public Function GetListaPersoneForRibesDataGrid(ByVal Cognome As String, _
        ByVal Nome As String, _
        ByVal CodiceFiscale As String, ByVal PartitaIVA As String, ByVal Cod_Ente As String) As GetList

            Dim GetList As New GetList
            Dim objConn As New SqlConnection
            Dim objCmd As New SqlCommand

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================
                objCmd.Parameters.Clear()

                GetList.lngRecordCount = objDBAccess.GetPrivateRunSPForRibesDataGrid("sp_RicercaPersone", objConn, objCmd, _
            New SqlParameter("@Cognome", Cognome), _
            New SqlParameter("@Nome", Nome), _
            New SqlParameter("@CodiceFiscale", CodiceFiscale), _
            New SqlParameter("@PartitaIVA", PartitaIVA), _
            New SqlParameter("@ParametroRicerca", ""), _
            New SqlParameter("@CodContribuente", ""), _
            New SqlParameter("@DaRicontrollare", ""), _
            New SqlParameter("@Da", ""), _
            New SqlParameter("@A", ""), _
            New SqlParameter("@CodEnte", Cod_Ente))


                GetList.oConn = objConn
                GetList.oComm = objCmd
                'objCmd.Parameters.Clear() 'commentata per OpenTerritorio : errore di index nella datagrid Ale-Fabio 12/01/2005


                Return GetList
            Catch ex As Exception

                Throw New Exception("Anagrafica::GetListaPersoneForRibesDataGrid::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************

            End Try
        End Function
        '*****************************************************************************************
        '////Parametri:Cognome,Nome,CodiceFiscale,PartitaIva
        '////Da usarsi per esempio in una pagina di ricerca dove è necessario visualizzare i dati in una tabella o griglia
        '////Usa una Stroed Procedure per caricare un oggetto DataGrid di tipo RibesDataGrid (ASP.Net)
        '////La Griglia sarà caricata con le anagrafiche storicizzate presenti a sistema
        '******************************************************************************************
        Public Function GetListaPersoneForRibesDataGridStorico(ByVal Cognome As String, _
        ByVal Nome As String, _
        ByVal CodiceFiscale As String, ByVal PartitaIVA As String, ByVal Cod_Ente As String) As GetList

            Dim GetList As New GetList
            Dim objConn As New SqlConnection
            Dim objCmd As New SqlCommand

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================

            GetList.lngRecordCount = objDBAccess.GetPrivateRunSPForRibesDataGrid("sp_RicercaPersoneStorico", objConn, objCmd, _
        New SqlParameter("@Cognome", Cognome), _
        New SqlParameter("@Nome", Nome), _
        New SqlParameter("@CodiceFiscale", CodiceFiscale), _
        New SqlParameter("@PartitaIVA", PartitaIVA), _
            New SqlParameter("@CodEnte", Cod_Ente))


            GetList.oConn = objConn
            GetList.oComm = objCmd
            'objCmd.Parameters.Clear() 'commentata per OpenTerritorio : errore di index nella datagrid Ale-Fabio 12/01/2005

            Return GetList

        End Function

        Public Sub VerificaEsistenzaPersona(ByVal oDettaglioAnagrafica As DettaglioAnagrafica)

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================
                Dim status As Int32 = objDBAccess.GetPrivateRunSPReturnRowCount("VerificaEsistenzaPersona", New SqlParameter("@CodiceFiscale", oDettaglioAnagrafica.CodiceFiscale))

                Select Case status
                    Case Is > 0
                        '********************Gestione Anagrafiche massive****************************
                        objDBAccess.DisposeConnection()
                        objDBAccess.Dispose()
                        '********************Gestione Anagrafiche massive****************************

                        Throw New Exception("l'Anagrafica che si cerca di inserire è già esistente")


                End Select
            Catch ex As Exception
                Throw New Exception("Anagrafica::VerificaEsistenzaPersona::" & ex.Message)
            Finally

                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Sub

        Public Sub VerificaEsistenzaAzienda(ByVal oDettaglioAnagrafica As DettaglioAnagrafica)


            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Dim status As Int32 = objDBAccess.GetPrivateRunSPReturnRowCount("VerificaEsistenzaAzienda", New SqlParameter("@PartitaIva", oDettaglioAnagrafica.PartitaIva))

                Select Case status
                    Case Is > 0
                        Throw New Exception("l'Azienda che si cerca di inserire è già esistente")
                End Select
            Catch ex As Exception
                Throw New Exception("Anagrafica::VerificaEsistenzaAzienda::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Sub
        '******************************************************************************************************
        '////Parametri:oDettaglioAnagrafica As DettaglioAnagrafica
        '////La funzione serve per verificare se un record e' stato modifcato mentre un altro utente era n modifica sullo stesso record
        '////Usa una Stored Procedure che verifica se c'è una concorrenza o se il record è stato eleiminato dal DataBase
        '*******************************************************************************************************

        Public Sub UpdateForLock(ByVal oDettaglioAnagrafica As DettaglioAnagrafica)
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Dim status As Int32 = objDBAccess.GetPrivateRunSPReturnRowCount("AnagraficaUpdate", New SqlParameter("@CodContribuente", oDettaglioAnagrafica.COD_CONTRIBUENTE), _
            New SqlParameter("@IdDataAnagrafica", oDettaglioAnagrafica.ID_DATA_ANAGRAFICA), New SqlParameter("@Concurrency", oDettaglioAnagrafica.Concurrency))

                Select Case status
                    Case UpdateRecordStatus.Concurrency
                        Throw New DBConcurrencyException("Il Record è stato  modificato da un altro utente")
                    Case UpdateRecordStatus.Deleted
                        Throw New DeletedRowInaccessibleException("Il Record è stato Eliminato dal Data Base")
                End Select
            Catch ex As Exception
                Throw New Exception("Anagrafica::UpdateForLock::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Sub

        '*****************************************************************************************************************
        '////Parametri:COD_CONTRIBUENTE,IDDATAANAGRAFICA valori da passare
        '////IDCOD_CONTRIBUENTE_FIGLIO,IDDATAANAGRAFICA_FIGLIO valori ritornati
        '////La funzione serve per estrarre gli ID necessari per la ricerca dei campi modificati da un altro utente in caso di concorrenza sullo stesso
        '////record da parte di due o più utenti
        '*****************************************************************************************************************
        Public Sub GetAnagraficaConcurrency(ByVal COD_CONTRIBUENTE As Integer, _
        ByRef IDCOD_CONTRIBUENTE_FIGLIO As Integer, ByRef IDDATAANAGRAFICA_FIGLIO As Integer)

            Dim strSql As String

            strSql = ""
            strSql = "SELECT * FROM ANAGRAFICA" & vbCrLf
            strSql = strSql & "INNER JOIN" & vbCrLf
            strSql = strSql & "DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE" & vbCrLf
            strSql = strSql & "AND" & vbCrLf
            strSql = strSql & "ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA" & vbCrLf
            strSql = strSql & "WHERE" & vbCrLf
            strSql = strSql & "DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE & vbCrLf
            strSql = strSql & "AND" & vbCrLf
            strSql = strSql & "(ANAGRAFICA.DATA_FINE_VALIDITA IS NOT NULL OR ANAGRAFICA.DATA_FINE_VALIDITA <>'')"

            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Try
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                '===============================================================================
                'WorkFlow
                '===============================================================================

                Dim drDetailsAnagrafica As SqlDataReader = objDBAccess.GetPrivateDataReader(strSql)

                If drDetailsAnagrafica.Read Then
                    IDCOD_CONTRIBUENTE_FIGLIO = Utility.cTolng(drDetailsAnagrafica.Item("IDCODCONTRIBUENTEFIGLIO"))
                    IDDATAANAGRAFICA_FIGLIO = Utility.cTolng(drDetailsAnagrafica.Item("IDDATAANAGRAFICAFIGLIO"))
                End If

                drDetailsAnagrafica.Close()
            Catch ex As Exception
                Throw New Exception("Anagrafica::GetAnagraficaConcurrency::" & ex.Message)
            Finally
                '********************Gestione Anagrafiche massive****************************
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
                '********************Gestione Anagrafiche massive****************************
            End Try
        End Sub

        '===================================================================================
        'Se vi sono state delle storicizzazione per un determinato contribuente vengono visulizzate in una lista
        '===================================================================================
        Public Function ViewListaPersoneStorico(ByVal Cognome As String, _
        ByVal Nome As String) As GetList

            Dim GetList As New GetList
            Dim objConn As New SqlConnection
            Dim objCmd As New SqlCommand
            '===============================================================================
            'WorkFlow
            '===============================================================================
            Dim objDBAccess As New RIBESFrameWork.DBManager
            objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
            '===============================================================================
            'WorkFlow
            '===============================================================================
            GetList.lngRecordCount = objDBAccess.GetPrivateRunSPForRibesDataGrid("sp_ListaPersoneStorico", objConn, objCmd, _
        New SqlParameter("@Cognome", Cognome), _
        New SqlParameter("@Nome", Nome))


            GetList.oConn = objConn
            GetList.oComm = objCmd
            'objCmd.Parameters.Clear() 'commentata per OpenTerritorio : errore di index nella datagrid Ale-Fabio 12/01/2005

            Return GetList

        End Function

        Public Function hasImmobiliICI(ByVal COD_CONTRIBUENTE As Integer, ByVal COD_ENTE As String) As Boolean
            Dim strSql As String
            Dim NOME_DATABASE_ICI As String
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Dim drDetailsAnagrafica As SqlDataReader
            Try
                hasImmobiliICI = False

                NOME_DATABASE_ICI = ConfigurationSettings.AppSettings("NOME_DATABASE_ICI")
                strSql = "SELECT *"
                strSql = strSql & " FROM " & NOME_DATABASE_ICI & ".dbo.TblOggetti  inner JOIN"
                strSql = strSql & " " & NOME_DATABASE_ICI & ".dbo.TblTestata ON " & NOME_DATABASE_ICI & ".dbo.TblOggetti.IdTestata = " & NOME_DATABASE_ICI & ".dbo.TblTestata.ID"
                strSql = strSql & " where " & NOME_DATABASE_ICI & ".dbo.TblTestata.Ente = '" & COD_ENTE & "'"
                strSql = strSql & " AND " & NOME_DATABASE_ICI & ".dbo.TblTestata.IDContribuente = " & COD_CONTRIBUENTE

                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                drDetailsAnagrafica = objDBAccess.GetPrivateDataReader(strSql)
                If drDetailsAnagrafica.HasRows Then
                    hasImmobiliICI = True
                End If
                drDetailsAnagrafica.Close()

            Catch ex As Exception
                Throw New Exception("Anagrafica::hasImmobiliICI::" & ex.Message)
                hasImmobiliICI = False
            Finally
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
            End Try

        End Function

        Public Function hasImmobiliTARSU(ByVal COD_CONTRIBUENTE As Integer, ByVal COD_ENTE As String) As Boolean
            Dim strSql As String
            Dim NOME_DATABASE_TARSU As String
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Dim drDetailsAnagrafica As SqlDataReader
            Try
                hasImmobiliTARSU = False

                NOME_DATABASE_TARSU = ConfigurationSettings.AppSettings("NOME_DATABASE_TARSU")
                strSql = "SELECT *"
                strSql = strSql & " FROM " & NOME_DATABASE_TARSU & ".dbo.TBLDETTAGLIOTESTATATARSU INNER JOIN"
                strSql = strSql & " " & NOME_DATABASE_TARSU & ".dbo.TBLTESTATATARSU ON " & NOME_DATABASE_TARSU & ".dbo.TBLDETTAGLIOTESTATATARSU.IDTESTATA = " & NOME_DATABASE_TARSU & ".dbo.TBLTESTATATARSU.ID"
                strSql = strSql & " WHERE " & NOME_DATABASE_TARSU & ".dbo.TBLTESTATATARSU.IDCONTRIBUENTE = " & COD_CONTRIBUENTE
                strSql = strSql & "  AND " & NOME_DATABASE_TARSU & ".dbo.TBLTESTATATARSU.IDENTE = '" & COD_ENTE & "'"
                strSql = strSql & "  AND " & NOME_DATABASE_TARSU & ".dbo.TBLDETTAGLIOTESTATATARSU.DATA_VARIAZIONE IS NULL"
                strSql = strSql & "  AND " & NOME_DATABASE_TARSU & ".dbo.TBLTESTATATARSU.DATA_VARIAZIONE IS NULL"

                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                drDetailsAnagrafica = objDBAccess.GetPrivateDataReader(strSql)
                If drDetailsAnagrafica.HasRows Then
                    hasImmobiliTARSU = True
                End If
                drDetailsAnagrafica.Close()

            Catch ex As Exception
                Throw New Exception("Anagrafica::hasImmobiliTARSU::" & ex.Message)
                hasImmobiliTARSU = False
            Finally
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
            End Try

        End Function
        Public Function hasContatoriH2O(ByVal COD_CONTRIBUENTE As Integer, ByVal COD_ENTE As String) As Boolean
            Dim strSql As String
            Dim NOME_DATABASE_H20 As String
            Dim objDBAccess As New RIBESFrameWork.DBManager
            Dim drDetailsAnagrafica As SqlDataReader
            Try
                hasContatoriH2O = False

                NOME_DATABASE_H20 = ConfigurationSettings.AppSettings("NOME_DATABASE_H20")
                strSql = "SELECT *"
                strSql = strSql & " FROM " & NOME_DATABASE_H20 & ".dbo.TP_CONTATORI INNER JOIN"
                strSql = strSql & " " & NOME_DATABASE_H20 & ".dbo.TR_CONTATORI_UTENTE ON " & NOME_DATABASE_H20 & ".dbo.TP_CONTATORI.CODCONTATORE = " & NOME_DATABASE_H20 & ".dbo.TR_CONTATORI_UTENTE.CODCONTATORE"
                strSql = strSql & " WHERE " & NOME_DATABASE_H20 & ".dbo.TR_CONTATORI_UTENTE.COD_CONTRIBUENTE =  " & COD_CONTRIBUENTE
                strSql = strSql & " AND " & NOME_DATABASE_H20 & ".dbo.TP_CONTATORI.CODENTE =  " & COD_ENTE

                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)
                drDetailsAnagrafica = objDBAccess.GetPrivateDataReader(strSql)
                If drDetailsAnagrafica.HasRows Then
                    hasContatoriH2O = True
                End If
                drDetailsAnagrafica.Close()

                strSql = "SELECT *"
                strSql = strSql & " FROM " & NOME_DATABASE_H20 & ".dbo.TP_CONTATORI INNER JOIN"
                strSql = strSql & " " & NOME_DATABASE_H20 & ".dbo.TR_CONTATORI_INTESTATARIO ON " & NOME_DATABASE_H20 & ".dbo.TP_CONTATORI.CODCONTATORE = " & NOME_DATABASE_H20 & ".dbo.TR_CONTATORI_INTESTATARIO.CODCONTATORE"
                strSql = strSql & " WHERE " & NOME_DATABASE_H20 & ".dbo.TP_CONTATORI.CODENTE = " & COD_ENTE
                strSql = strSql & " AND " & NOME_DATABASE_H20 & ".dbo.TR_CONTATORI_INTESTATARIO.COD_CONTRIBUENTE = " & COD_CONTRIBUENTE
                objDBAccess = m_oSession.GetPrivateDBManager(m_IDSottoAttivita)

                drDetailsAnagrafica = objDBAccess.GetPrivateDataReader(strSql)
                If drDetailsAnagrafica.HasRows Then
                    hasContatoriH2O = True
                End If
                drDetailsAnagrafica.Close()

            Catch ex As Exception
                Throw New Exception("Anagrafica::hasContatoriH2O::" & ex.Message)
                hasContatoriH2O = False
            Finally
                objDBAccess.DisposeConnection()
                objDBAccess.Dispose()
            End Try

        End Function


        Public Sub New()

        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
