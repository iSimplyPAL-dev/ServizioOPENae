Imports log4net
'Imports OPENUtility
'Imports RIBESFrameWork

Namespace OPENgov_AgenziaEntrate
    Public Class LevelDB
        Private Log As ILog = LogManager.GetLogger("ClsLevelDB")
        Private cmdMyCommand As New SqlClient.SqlCommand
        Private myAdapter As New SqlClient.SqlDataAdapter
        Private dtMyDati As New DataTable()

#Region "RIBESFRAMEWORK"
        'Protected Function GetTitoloOccupazione(ByVal WFSession As CreateSessione, ByVal sTributo As String) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_TITOLO_OCCUPAZIONE.ID AS NVARCHAR)+ ' - ' + DBO.AE_TIPO_TITOLO_OCCUPAZIONE.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_TITOLO_OCCUPAZIONE.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_TITOLO_OCCUPAZIONE"
        '        cmdMyCommand.CommandText += " WHERE (DBO.AE_TIPO_TITOLO_OCCUPAZIONE.COD_TRIBUTO='" & sTributo & "')"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_TITOLO_OCCUPAZIONE.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetTitoloOccupazione::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetTitoloOccupazione::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetNaturaOccupazione(ByVal WFSession As CreateSessione) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_NATURA_OCCUPANTE.ID AS NVARCHAR) +' - '+DBO.AE_TIPO_NATURA_OCCUPANTE.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_NATURA_OCCUPANTE.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_NATURA_OCCUPANTE"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_NATURA_OCCUPANTE.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetNaturaOccupazione::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetNaturaOccupazione::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetTipoUtenza(ByVal WFSession As CreateSessione) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_TIPOLOGIA_UTENZA.ID AS NVARCHAR) +' - '+DBO.AE_TIPO_TIPOLOGIA_UTENZA.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_TIPOLOGIA_UTENZA.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_TIPOLOGIA_UTENZA"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_TIPOLOGIA_UTENZA.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoUtenza::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetTipoUtenza::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetDestUso(ByVal WFSession As CreateSessione) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_DESTINAZIONE_USO.ID AS NVARCHAR)+' - '+DBO.AE_TIPO_DESTINAZIONE_USO.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_DESTINAZIONE_USO.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_DESTINAZIONE_USO"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_DESTINAZIONE_USO.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetDestUso::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetDestUso::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetTipoUnita(ByVal WFSession As CreateSessione) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_TIPOLOGIA_UNITA.ID AS NVARCHAR)+' - '+DBO.AE_TIPO_TIPOLOGIA_UNITA.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_TIPOLOGIA_UNITA.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_TIPOLOGIA_UNITA"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_TIPOLOGIA_UNITA.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoUnita::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetTipoUnita::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetTipoParticella(ByVal WFSession As CreateSessione) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.ID AS NVARCHAR)+' - '+DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_TIPOLOGIA_PARTICELLA"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoParticella::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetTipoParticella::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetAssenzaDatiCat(ByVal WFSession As CreateSessione, ByVal sTributo As String) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT CAST(DBO.AE_TIPO_ASSENZA_DATI_CATASTALI.ID AS NVARCHAR)+' - '+DBO.AE_TIPO_ASSENZA_DATI_CATASTALI.DESCRIZIONE AS DESCRIZIONE, DBO.AE_TIPO_ASSENZA_DATI_CATASTALI.ID"
        '        cmdMyCommand.CommandText += " FROM DBO.AE_TIPO_ASSENZA_DATI_CATASTALI"
        '        cmdMyCommand.CommandText += " WHERE (DBO.AE_TIPO_ASSENZA_DATI_CATASTALI.COD_TRIBUTO='" & sTributo & "')"
        '        cmdMyCommand.CommandText += " ORDER BY DBO.AE_TIPO_ASSENZA_DATI_CATASTALI.ID"
        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoParticella::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetTipoParticella::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        'Protected Function GetDatiMancanti(ByVal sMyIdEnte As String, ByVal sMyAnno As String, ByVal sMyCognome As String, ByVal sMyNome As String, ByVal nMyTypeMancanti As Integer, ByVal WFSession As CreateSessione) As DataView
        '    Try
        '        'Valorizzo la connessione
        '        cmdMyCommand.Connection = WFSession.oSession.oAppDB.GetConnection
        '        'Valorizzo il commandtext:
        '        cmdMyCommand.CommandText = "SELECT DBO.OPENae_GET_DATIMANCATI.COD_CONTRIBUENTE, DBO.OPENae_GET_DATIMANCATI.ID_RUOLO, DBO.OPENae_GET_DATIMANCATI.ANNO, "
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.COGNOME, DBO.OPENae_GET_DATIMANCATI.NOME, DBO.OPENae_GET_DATIMANCATI.CFPIVA, "
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.VIA_RES, DBO.OPENae_GET_DATIMANCATI.CIVICO_RES, DBO.OPENae_GET_DATIMANCATI.CAP_RES,  DBO.OPENae_GET_DATIMANCATI.COMUNE_RES,  DBO.OPENae_GET_DATIMANCATI.PROVINCIA_RES, "
        '        cmdMyCommand.CommandText += " DBO.INDIRIZZI_SPEDIZIONE.COGNOME_INVIO+' '+DBO.INDIRIZZI_SPEDIZIONE.NOME_INVIO AS NOMINATIVO, DBO.INDIRIZZI_SPEDIZIONE.VIA_RCP, DBO.INDIRIZZI_SPEDIZIONE.CIVICO_RCP, DBO.INDIRIZZI_SPEDIZIONE.CAP_RCP, DBO.INDIRIZZI_SPEDIZIONE.COMUNE_RCP, DBO.INDIRIZZI_SPEDIZIONE.PROVINCIA_RCP,"
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.IND_IMMO, DBO.OPENae_GET_DATIMANCATI.FOGLIO, DBO.OPENae_GET_DATIMANCATI.PARTICELLA, DBO.OPENae_GET_DATIMANCATI.ESTENSIONE_PARTICELLA,"
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.DATA_INIZIO, DBO.AE_TIPO_TITOLO_OCCUPAZIONE.DESCRIZIONE AS TITOCCUP, DBO.AE_TIPO_NATURA_OCCUPANTE.DESCRIZIONE AS NATOCCUP, DBO.AE_TIPO_DESTINAZIONE_USO.DESCRIZIONE AS DESTUSO,"
        '        cmdMyCommand.CommandText += " DBO.AE_TIPO_TIPOLOGIA_UNITA.DESCRIZIONE AS TIPOUNITA, DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.DESCRIZIONE AS TIPOPARTICELLA,"
        '        cmdMyCommand.CommandText += " SUM(DBO.OPENae_GET_DATIMANCATI.IDANOMALIA) AS ANOMALIA"
        '        cmdMyCommand.CommandText += " FROM DBO.OPENae_GET_DATIMANCATI"
        '        cmdMyCommand.CommandText += " LEFT JOIN DBO.INDIRIZZI_SPEDIZIONE ON DBO.OPENae_GET_DATIMANCATI.COD_CONTRIBUENTE=DBO.INDIRIZZI_SPEDIZIONE.COD_CONTRIBUENTE"
        '        cmdMyCommand.CommandText += " LEFT JOIN DBO.AE_TIPO_TITOLO_OCCUPAZIONE ON DBO.OPENae_GET_DATIMANCATI.ID_TITOLO_OCCUPAZIONE=DBO.AE_TIPO_TITOLO_OCCUPAZIONE.ID"
        '        cmdMyCommand.CommandText += " LEFT JOIN DBO.AE_TIPO_NATURA_OCCUPANTE ON DBO.OPENae_GET_DATIMANCATI.ID_NATURA_OCCUPANTE=DBO.AE_TIPO_NATURA_OCCUPANTE.ID"
        '        cmdMyCommand.CommandText += " LEFT JOIN DBO.AE_TIPO_DESTINAZIONE_USO ON DBO.OPENae_GET_DATIMANCATI.ID_DESTINAZIONE_USO=DBO.AE_TIPO_DESTINAZIONE_USO.ID"
        '        cmdMyCommand.CommandText += " LEFT JOIN DBO.AE_TIPO_TIPOLOGIA_UNITA ON DBO.OPENae_GET_DATIMANCATI.ID_TIPO_UNITA=DBO.AE_TIPO_TIPOLOGIA_UNITA.ID"
        '        cmdMyCommand.CommandText += " LEFT JOIN DBO.AE_TIPO_TIPOLOGIA_PARTICELLA ON DBO.OPENae_GET_DATIMANCATI.ID_TIPO_PARTICELLA=DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.ID"
        '        cmdMyCommand.CommandText += " WHERE (DBO.INDIRIZZI_SPEDIZIONE.DATA_FINE_VALIDITA IS NULL) "
        '        'Valorizzo i parameters:
        '        cmdMyCommand.Parameters.Clear()
        '        cmdMyCommand.CommandText += " AND (DBO.OPENae_GET_DATIMANCATI.CODICE_ISTAT=@CODISTAT)"
        '        cmdMyCommand.Parameters.Add(New SqlParameter("@CODISTAT", SqlDbType.NVarChar)).Value = sMyIdEnte
        '        If nMyTypeMancanti <> 0 Then
        '            cmdMyCommand.CommandText += " AND (DBO.OPENae_GET_DATIMANCATI.IDANOMALIA=@TIPOANOMALIA)"
        '            cmdMyCommand.Parameters.Add(New SqlParameter("@TIPOANOMALIA", SqlDbType.Int)).Value = nMyTypeMancanti
        '        End If
        '        If sMyAnno <> "" Then
        '            cmdMyCommand.CommandText += " AND (DBO.OPENae_GET_DATIMANCATI.ANNO=@ANNO)"
        '            cmdMyCommand.Parameters.Add(New SqlParameter("@ANNO", SqlDbType.NVarChar)).Value = sMyAnno
        '        End If
        '        If sMyCognome <> "" Then
        '            cmdMyCommand.CommandText += " AND (DBO.OPENae_GET_DATIMANCATI.COGNOME=@COGNOME)"
        '            cmdMyCommand.Parameters.Add(New SqlParameter("@COGNOME", SqlDbType.NVarChar)).Value = sMyCognome
        '        End If
        '        If sMyNome <> "" Then
        '            cmdMyCommand.CommandText += " AND (DBO.OPENae_GET_DATIMANCATI.NOME=@NOME)"
        '            cmdMyCommand.Parameters.Add(New SqlParameter("@NOME", SqlDbType.NVarChar)).Value = sMyNome
        '        End If
        '        cmdMyCommand.CommandText += " GROUP BY  DBO.OPENae_GET_DATIMANCATI.COD_CONTRIBUENTE, DBO.OPENae_GET_DATIMANCATI.ID_RUOLO, DBO.OPENae_GET_DATIMANCATI.ANNO, "
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.COGNOME, DBO.OPENae_GET_DATIMANCATI.NOME, DBO.OPENae_GET_DATIMANCATI.CFPIVA, "
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.VIA_RES, DBO.OPENae_GET_DATIMANCATI.CIVICO_RES, DBO.OPENae_GET_DATIMANCATI.CAP_RES,  DBO.OPENae_GET_DATIMANCATI.COMUNE_RES,  DBO.OPENae_GET_DATIMANCATI.PROVINCIA_RES, "
        '        cmdMyCommand.CommandText += " DBO.INDIRIZZI_SPEDIZIONE.COGNOME_INVIO+' '+DBO.INDIRIZZI_SPEDIZIONE.NOME_INVIO, DBO.INDIRIZZI_SPEDIZIONE.VIA_RCP, DBO.INDIRIZZI_SPEDIZIONE.CIVICO_RCP, DBO.INDIRIZZI_SPEDIZIONE.CAP_RCP, DBO.INDIRIZZI_SPEDIZIONE.COMUNE_RCP, DBO.INDIRIZZI_SPEDIZIONE.PROVINCIA_RCP,"
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.IND_IMMO, DBO.OPENae_GET_DATIMANCATI.FOGLIO, DBO.OPENae_GET_DATIMANCATI.PARTICELLA, DBO.OPENae_GET_DATIMANCATI.ESTENSIONE_PARTICELLA,"
        '        cmdMyCommand.CommandText += " DBO.OPENae_GET_DATIMANCATI.DATA_INIZIO, DBO.AE_TIPO_TITOLO_OCCUPAZIONE.DESCRIZIONE, DBO.AE_TIPO_NATURA_OCCUPANTE.DESCRIZIONE, DBO.AE_TIPO_DESTINAZIONE_USO.DESCRIZIONE,"
        '        cmdMyCommand.CommandText += " DBO.AE_TIPO_TIPOLOGIA_UNITA.DESCRIZIONE, DBO.AE_TIPO_TIPOLOGIA_PARTICELLA.DESCRIZIONE"

        '        'eseguo la query
        '        DvDati = WFSession.oSession.oAppDB.GetPrivateDataview(cmdMyCommand)
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetDatiMancanti::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsLevelDB::GetDatiMancanti::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
#End Region

#Region "Select"
        Protected Function GetTitoloOccupazione(ByVal myStringConnection As String, ByVal sTributo As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAETitoloOccupazione"
                cmdMyCommand.Parameters.Clear()
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IdTributo", SqlDbType.VarChar)).Value = sTributo
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::GetTitoloOccupazione::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
        End Function
        Protected Function GetNaturaOccupazione(ByVal myStringConnection As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAENaturaOccupazione"
                cmdMyCommand.Parameters.Clear()
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::GetNaturaOccupazione::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
        End Function
        Protected Function GetTipoUtenza(ByVal myStringConnection As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAETipoUtenza"
                cmdMyCommand.Parameters.Clear()
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoUtenza::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
          End Function
        Protected Function GetDestUso(ByVal myStringConnection As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAEDestUso"
                cmdMyCommand.Parameters.Clear()
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::GetDestUso::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
         End Function
        Protected Function GetTipoUnita(ByVal myStringConnection As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAETipoUnita"
                cmdMyCommand.Parameters.Clear()
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoUnita::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
        End Function
        Protected Function GetTipoParticella(ByVal myStringConnection As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAETipoParticella"
                cmdMyCommand.Parameters.Clear()
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::GetTipoParticella::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
        End Function
        Protected Function GetAssenzaDatiCat(ByVal myStringConnection As String, ByVal sTributo As String) As DataView
            Try
                cmdMyCommand = New SqlClient.SqlCommand
                myAdapter = New SqlClient.SqlDataAdapter
                dtMyDati = New DataTable
                cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
                cmdMyCommand.Connection.Open()
                cmdMyCommand.CommandTimeout = 0
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                cmdMyCommand.CommandText = "prc_GetAEAssenzaDatiCat"
                cmdMyCommand.Parameters.Clear()
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IdTributo", SqlDbType.VarChar)).Value = sTributo
                myAdapter.SelectCommand = cmdMyCommand
                myAdapter.Fill(dtMyDati)
                myAdapter.Dispose()
                Return dtMyDati.DefaultView
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsLevelDB::::GetAssenzaDatiCat" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
         End Function
        'Protected Function GetDatiMancanti(ByVal sMyIdEnte As String, ByVal sMyAnno As String, ByVal sMyCognome As String, ByVal sMyNome As String, ByVal nMyTypeMancanti As Integer, ByVal myStringConnection As String) As DataView
        '    Try
        '        cmdMyCommand = New SqlClient.SqlCommand
        '        myAdapter = New SqlClient.SqlDataAdapter
        '        dtMyDati = New DataTable
        '        cmdMyCommand.Connection = New SqlClient.SqlConnection(myStringConnection)
        '        cmdMyCommand.Connection.Open()
        '        cmdMyCommand.CommandTimeout = 0
        '        cmdMyCommand.CommandType = CommandType.StoredProcedure
        '        cmdMyCommand.CommandText = "prc_GetAEDatiMancanti"
        '        cmdMyCommand.Parameters.Clear()
        '        cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODISTAT", SqlDbType.NVarChar)).Value = sMyIdEnte
        '        cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TIPOANOMALIA", SqlDbType.Int)).Value = nMyTypeMancanti
        '        cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNO", SqlDbType.NVarChar)).Value = sMyAnno
        '        cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COGNOME", SqlDbType.NVarChar)).Value = sMyCognome
        '        cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOME", SqlDbType.NVarChar)).Value = sMyNome
        '        myAdapter.SelectCommand = cmdMyCommand
        '        myAdapter.Fill(dtMyDati)
        '        myAdapter.Dispose()
        '        Return dtMyDati.DefaultView
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsLevelDB::GetDatiMancanti::" & Err.Message)
        '        Return Nothing
        '    Finally
        '        dtMyDati.Dispose()
        '        cmdMyCommand.Dispose()
        '        cmdMyCommand.Connection.Close()
        '    End Try
        'End Function
#End Region
    End Class
End Namespace