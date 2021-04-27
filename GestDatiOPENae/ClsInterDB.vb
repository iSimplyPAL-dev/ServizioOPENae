Imports log4net

Public Class ClsInterDB
    Private Shared Log As ILog = LogManager.GetLogger("ClsInterDB")
    Private cmdMyCommand As SqlClient.SqlCommand
    Private myResult As Integer = -1
    Private FncGen As New General

#Region "Query di SELEZIONE"
    Public Function GetFlussiCodIstat(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sTributo As String) As DataView
        Dim dvResult As DataView
        Try
            'mySQLCommand = New SqlClient.SqlCommand
            'mySQLCommand.CommandType = CommandType.Text
            'mySQLCommand.CommandText = "SELECT AE_FLUSSI_ESTRATTI.ID, AE_FLUSSI_ESTRATTI.CODICE_ISTAT, AE_FLUSSI_ESTRATTI.ANNO, AE_FLUSSI_ESTRATTI.NOME_FILE, AE_FLUSSI_ESTRATTI.DATA_ESTRAZIONE, AE_FLUSSI_ESTRATTI.NUTENTI, AE_FLUSSI_ESTRATTI.NRECORD, AE_FLUSSI_ESTRATTI.NARTICOLI"
            'mySQLCommand.CommandText += " FROM AE_FLUSSI_ESTRATTI"
            'mySQLCommand.CommandText += " WHERE (NOT AE_FLUSSI_ESTRATTI.NUTENTI IS NULL)"
            'mySQLCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CodiceISTAT)"
            'mySQLCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.COD_TRIBUTO=@CodTributo)"
            'mySQLCommand.CommandText += " ORDER BY AE_FLUSSI_ESTRATTI.ANNO"
            'Log.Debug("GetFlussiCodIstat->" & mySQLCommand.CommandText)
            'Log.Debug("@CodiceISTAT->" & sCodIstat)
            'Log.Debug("@CodTributo->" & sTributo)
            'Log.Debug("")
            ''valorizzo i parameters
            'mySQLCommand.Parameters.Clear()
            'mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("CodiceISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("CodTributo", SqlDbType.NVarChar)).Value = sTributo
            'dvResult = oDBManager.GetDataView(mySQLCommand.CommandText, "RESULT")

            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_GetFlussiEstratti", "CodiceISTAT", "CodTributo")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CodiceISTAT", sCodIstat), oDBManager.GetParam("CodTributo", sTributo))

            Return dvResult
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetFlussiCodIstat::" & Err.Message)
            Return Nothing
        End Try
    End Function

    Public Function GetNUtenti(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String, ByVal sTributo As String) As Integer
        'Dim drDati As SqlClient.SqlDataReader

        Try
            myResult = 0
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "SELECT COUNT(TMPCOUNT.COD_CONTRIBUENTE) AS NUTENTI"
            'cmdMyCommand.CommandText += " FROM ("
            'cmdMyCommand.CommandText += " SELECT DISTINCT AE_DATI_FILE.COD_CONTRIBUENTE "
            'cmdMyCommand.CommandText += " FROM AE_DATI_FILE WITH (NOLOCK)"
            'cmdMyCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CODICEISTAT)"
            'cmdMyCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@ANNORIF)"
            'cmdMyCommand.CommandText += " AND (AE_DATI_FILE.COD_TRIBUTO=@CODTRIBUTO)) AS TMPCOUNT"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            'drDati = oDBManager.GetDataReader(cmdMyCommand.CommandText)
            'Do While drDati.Read
            '    If Not IsDBNull(drDati("nutenti")) Then
            '        myResult += CInt(drDati("nutenti"))
            '    End If
            'Loop
            'drDati.Close()
            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_GetNUtenti", "CODICEISTAT", "ANNORIF", "CODTRIBUTO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo))
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next

        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetNUtenti::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function GetNArticoli(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String, ByVal sTributo As String) As Integer
        'Dim drDati As SqlClient.SqlDataReader

        Try
            myResult = 0
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "SELECT COUNT(TMPCOUNT.ID_RUOLO) AS NARTICOLI"
            'cmdMyCommand.CommandText += " FROM ("
            'cmdMyCommand.CommandText += " SELECT DISTINCT AE_DATI_FILE.ID_RUOLO "
            'cmdMyCommand.CommandText += " FROM AE_DATI_FILE WITH (NOLOCK)"
            'cmdMyCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CODICEISTAT)"
            'cmdMyCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@ANNORIF)"
            'cmdMyCommand.CommandText += " AND (AE_DATI_FILE.COD_TRIBUTO=@CODTRIBUTO)) AS TMPCOUNT"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            'drDati = oDBManager.GetDataReader(cmdMyCommand.CommandText)
            'Do While drDati.Read
            '    If Not IsDBNull(drDati("narticoli")) Then
            '        myResult += CInt(drDati("narticoli"))
            '    End If
            'Loop
            'drDati.Close()
            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_GetNArticoli", "CODICEISTAT", "ANNORIF", "CODTRIBUTO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo))
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetNArticoli::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function GetDisposizione(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sTributo As String, ByVal sAnno As String) As DataView
        Dim dvResult As DataView
        Try
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "SELECT *"
            'cmdMyCommand.CommandText += " FROM V_GETDISPOSIZIONE_DATIFILE"
            'cmdMyCommand.CommandText += " WHERE (CODICE_ISTAT=@CODICEISTAT)"
            'cmdMyCommand.CommandText += " AND (COD_TRIBUTO=@CODTRIBUTO)"
            'cmdMyCommand.CommandText += " AND (ANNO=@ANNORIF)"
            'cmdMyCommand.CommandText += " ORDER BY COD_CONTRIBUENTE, ID_RUOLO"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'dvResult = oDBManager.GetDataView(cmdMyCommand.CommandText, "RESULT")
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_GetDisposizione", "CODICEISTAT", "ANNORIF", "CODTRIBUTO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo))

            Return dvResult
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetDisposizione::" & Err.Message)
            Return Nothing
        End Try
    End Function

    Public Function GetAnagrafe(ByVal oDBManager As Utility.DBModel, ByVal sEnte As String, ByVal nIdContrib As Integer) As DataView
        Dim dvResult As DataView
        Try
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "SELECT *"
            cmdMyCommand.CommandText += " FROM AE_ANAGRAFICA_ICI"
            cmdMyCommand.CommandText += " WHERE (AE_ANAGRAFICA_ICI.COD_CONTRIBUENTE=@CODCONTRIBUENTE)"
            cmdMyCommand.CommandText += " AND (AE_ANAGRAFICA_ICI.CODICE_ISTAT=@CODICEISTAT)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODCONTRIBUENTE", SqlDbType.Int)).Value = nIdContrib
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sEnte
            dvResult = oDBManager.GetDataView(cmdMyCommand.CommandText, "RESULT")

            Return dvResult
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetAnagrafe::" & Err.Message)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::GetAnagrafe::" & Err.Message)
            Return Nothing
        End Try
    End Function

    Public Function GetProgressivoInvioICI(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sTributo As String, ByVal sAnno As String) As Integer
        Dim drDati As SqlClient.SqlDataReader

        Try
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "SELECT TOP 1 AE_FLUSSI_ESTRATTI.PROGRESSIVO_INVIO"
            cmdMyCommand.CommandText += " FROM AE_FLUSSI_ESTRATTI"
            cmdMyCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.ANNO_IMPOSTA=@ANNO)"
            cmdMyCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.COD_TRIBUTO=@CODTRIBUTO)"
            cmdMyCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CODICEISTAT)"
            cmdMyCommand.CommandText += " ORDER BY AE_FLUSSI_ESTRATTI.PROGRESSIVO_INVIO DESC"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sAnno
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            drDati = oDBManager.GetDataReader(cmdMyCommand.CommandText)
            Do While drDati.Read
                If Not IsDBNull(drDati("progressivo_invio")) Then
                    myResult += CInt(drDati("progressivo_invio"))
                End If
            Loop
            drDati.Close()
            myResult += 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetProgressivoInvioICI::" & Err.Message)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::GetProgressivoInvioICI::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function GetTotaliInvioICI(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String, ByVal nNumQuietanza As Integer, ByVal sDataRiversamento As String, ByVal nProvenienza As Integer, ByRef sTipoRiscossioni As String, ByRef nTotRiversato As Double, ByRef nDisposizioni As Integer) As Boolean
        Dim dvResult As DataView
        Dim x As Integer
        Dim IsOrdinari As Integer = 0
        Dim IsViolazioni As Integer = 0

        Try
            cmdMyCommand = New SqlClient.SqlCommand
            nDisposizioni = 0 : nTotRiversato = 0
            cmdMyCommand.CommandText = "SELECT COUNT(AE_PAGAMENTI_ICI.IDDISPOSIZIONE) AS NDISPOSIZIONI, SUM(AE_PAGAMENTI_ICI.IMPORTO) AS TOTRIVERSATO,"
            cmdMyCommand.CommandText += " CASE WHEN DATA_SANZIONE IS NULL THEN '' ELSE DATA_SANZIONE END+CASE WHEN N_SANZIONE IS NULL THEN '' ELSE N_SANZIONE END+CASE WHEN TIPO_BOLLETTINO_VIOLAZIONI IS  NULL THEN '' ELSE TIPO_BOLLETTINO_VIOLAZIONI END AS TIPORISCOSSIONI"
            cmdMyCommand.CommandText += " FROM AE_PAGAMENTI_ICI"
            cmdMyCommand.CommandText += " WHERE (AE_PAGAMENTI_ICI.CODICE_ISTAT=@CODICEISTAT) AND (AE_PAGAMENTI_ICI.ANNO=@ANNO)"
            cmdMyCommand.CommandText += " AND (AE_PAGAMENTI_ICI.NUMERO_QUIETANZA=@NUMQUIETANZA)"
            cmdMyCommand.CommandText += " AND (AE_PAGAMENTI_ICI.DATA_ACCREDITO=@DATAACCREDITO)"
            If nProvenienza = 2 Then
                cmdMyCommand.CommandText += " AND (AE_PAGAMENTI_ICI.PROVENIENZA=@PROVENIENZA)"
            Else
                cmdMyCommand.CommandText += " AND (AE_PAGAMENTI_ICI.PROVENIENZA<>@PROVENIENZA)"
            End If
            cmdMyCommand.CommandText += " GROUP BY CASE WHEN DATA_SANZIONE IS NULL THEN '' ELSE DATA_SANZIONE END+CASE WHEN N_SANZIONE IS NULL THEN '' ELSE N_SANZIONE END+CASE WHEN TIPO_BOLLETTINO_VIOLAZIONI IS  NULL THEN '' ELSE TIPO_BOLLETTINO_VIOLAZIONI END"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNO", SqlDbType.NVarChar)).Value = sAnno
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NUMQUIETANZA", SqlDbType.Int)).Value = nNumQuietanza
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAACCREDITO", SqlDbType.NVarChar)).Value = sDataRiversamento
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PROVENIENZA", SqlDbType.NVarChar)).Value = FncGen.ICI_RIVERSAMENTOCONTOCORRENTE
            dvResult = oDBManager.GetDataView(cmdMyCommand.CommandText, "RESULT")
            If Not IsNothing(dvResult) Then
                For x = 0 To dvResult.Count - 1
                    nDisposizioni += dvResult.Item(x)("ndisposizioni")
                    nTotRiversato += dvResult.Item(x)("totriversato")
                    If Not IsDBNull(dvResult.Item(x)("tiporiscossioni")) Then
                        If CStr(dvResult.Item(x)("tiporiscossioni")) <> "" And CStr(dvResult.Item(x)("tiporiscossioni")) <> "0" Then
                            IsViolazioni = 1
                        Else
                            IsOrdinari = 1
                        End If
                    Else
                        IsOrdinari = 1
                    End If
                Next
            End If
            If IsOrdinari = 1 And IsViolazioni = 1 Then
                sTipoRiscossioni = "M"
            ElseIf IsOrdinari = 1 And IsViolazioni = 0 Then
                sTipoRiscossioni = "O"
            ElseIf IsOrdinari = 0 And IsViolazioni = 1 Then
                sTipoRiscossioni = "V"
            Else
                Return False
            End If
            Return True
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetTotaliInvioICI::" & Err.Message)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::GetTotaliInvioICI::" & Err.Message)
            Return False
        End Try
    End Function

    Public Function GetDisposizioneICI(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String) As DataView
        Dim dvResult As DataView
        Try
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "SELECT *"
            cmdMyCommand.CommandText += " FROM AE_ANAGRAFICA_ICI"
            cmdMyCommand.CommandText += " INNER JOIN AE_PAGAMENTI_ICI ON AE_ANAGRAFICA_ICI.COD_CONTRIBUENTE=AE_PAGAMENTI_ICI.COD_CONTRIBUENTE"
            cmdMyCommand.CommandText += " AND AE_ANAGRAFICA_ICI.CODICE_ISTAT=AE_PAGAMENTI_ICI.CODICE_ISTAT"
            cmdMyCommand.CommandText += " WHERE (AE_PAGAMENTI_ICI.CODICE_ISTAT=@CODICEISTAT) AND (AE_PAGAMENTI_ICI.ANNO_RIFERIMENTO=@ANNO)"
            cmdMyCommand.CommandText += " ORDER BY AE_PAGAMENTI_ICI.ANNO, AE_PAGAMENTI_ICI.NUMERO_QUIETANZA, AE_PAGAMENTI_ICI.DATA_ACCREDITO, CASE WHEN AE_PAGAMENTI_ICI.PROVENIENZA='POSTE' THEN 2 ELSE 0 END, AE_ANAGRAFICA_ICI.COD_CONTRIBUENTE"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNO", SqlDbType.NVarChar)).Value = sAnno
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            dvResult = oDBManager.GetDataView(cmdMyCommand.CommandText, "RESULT")

            Return dvResult
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::GetDisposizioneICI::" & Err.Message)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::GetDisposizioneICI::" & Err.Message)
            Return Nothing
        End Try
    End Function
#End Region

#Region "Query di INSERIMENTO/UPDATE"
    Public Function SetIdFlussoEstratto(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String, ByVal sTributo As String) As Integer
        Dim drDati As SqlClient.SqlDataReader

        Try
            myResult = -1
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "INSERT INTO AE_FLUSSI_ESTRATTI(CODICE_ISTAT, COD_TRIBUTO, ANNO)"
            'cmdMyCommand.CommandText += " VALUES (@CODICEISTAT,@CODTRIBUTO,@ANNORIF)"
            'cmdMyCommand.CommandText += " SELECT @@IDENTITY"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            ''eseguo la query
            'drDati = oDBManager.GetDataReader(cmdMyCommand.CommandText)
            'Do While drDati.Read
            '    myResult = drDati(0)
            'Loop
            'drDati.Close()
            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_AE_FLUSSI_ESTRATTI_IU", "CODICEISTAT", "ANNORIF", "CODTRIBUTO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo))
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetIdFlusso::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function SetFlussoEstratto(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, sTributo As String, ByVal sAnno As String, ByVal nUtenti As Integer, ByVal nArticoli As Integer) As Integer
        Try
            myResult = -1
            cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "UPDATE AE_FLUSSI_ESTRATTI"
            'cmdMyCommand.CommandText += " SET NUTENTI=@NUTENTI, NARTICOLI=@NARTICOLI"
            'cmdMyCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CODICEISTAT)"
            'cmdMyCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.ANNO=@ANNORIF)"
            'cmdMyCommand.CommandText = cmdMyCommand.CommandText
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NUTENTI", SqlDbType.Int)).Value = nUtenti
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NARTICOLI", SqlDbType.Int)).Value = nArticoli
            'oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)

            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_AE_FLUSSI_ESTRATTI_IU", "CODICEISTAT", "ANNORIF", "CODTRIBUTO", "NUTENTI", "NARTICOLI")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo), oDBManager.GetParam("NUTENTI", nUtenti), oDBManager.GetParam("NARTICOLI", nArticoli))
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next

            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetFlusso::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function SetFlussoEstratto(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, sTributo As String, ByVal sAnno As String, ByVal sNomeFileTracciati As String, ByVal nRcFile As Integer) As Integer
        Try
            myResult = -1
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "UPDATE AE_FLUSSI_ESTRATTI"
            'cmdMyCommand.CommandText += " SET NOME_FILE=@NOMEFILE, DATA_ESTRAZIONE=@DATAESTRAZIONE, NUTENTI=@NRC, NRECORD=@NRC"
            'cmdMyCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CODICEISTAT)"
            'cmdMyCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.ANNO=@ANNORIF)"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMEFILE", SqlDbType.NVarChar)).Value = FncGen.ReplaceChar(sNomeFileTracciati)
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAESTRAZIONE", SqlDbType.NVarChar)).Value = DateTime.Now.ToString("yyyyMMdd")
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NRC", SqlDbType.Int)).Value = nRcFile
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)

            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_AE_FLUSSI_ESTRATTI_IU", "CODICEISTAT", "ANNORIF", "CODTRIBUTO", "NUTENTI", "NARTICOLI", "NOMEFILE", "DATAESTRAZIONE", "NRC")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo) _
                    , oDBManager.GetParam("NUTENTI", 0), oDBManager.GetParam("NARTICOLI", 0) _
                    , oDBManager.GetParam("NOMEFILE", FncGen.ReplaceChar(sNomeFileTracciati)), oDBManager.GetParam("DATAESTRAZIONE", DateTime.Now.ToString("yyyyMMdd")), oDBManager.GetParam("NRC", nRcFile)
                )
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next
            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae:: SetFlusso::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function SetDisposizione(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String, ByVal nFlusso As Integer) As Integer
        Try
            myResult = -1
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "INSERT INTO AE_DATI_FILE("
            cmdMyCommand.CommandText += " ID_FLUSSO, CODICE_ISTAT, COD_TRIBUTO, ANNO,"
            cmdMyCommand.CommandText += " COD_FISCALE_ENTE, COGNOME_ENTE, NOME_ENTE, SESSO_ENTE, DATA_NASCITA_ENTE, COMUNE_NASCITA_SEDE_ENTE, PV_NASCITA_SEDE_ENTE,"
            cmdMyCommand.CommandText += " COD_CONTRIBUENTE, ID_RUOLO, COD_FISCALE,"
            cmdMyCommand.CommandText += " COGNOME, NOME, COMUNE_SEDE, PV_SEDE,"
            cmdMyCommand.CommandText += " COMUNE_UBICAZIONE, PV_UBICAZIONE, COD_COMUNE_UBICAZIONE_CATAST,"
            cmdMyCommand.CommandText += " DATA_INIZIO, DATA_FINE,"
            cmdMyCommand.CommandText += " ID_TITOLO_OCCUPAZIONE, ID_TIPO_OCCUPANTE, ID_DESTINAZIONE_USO, ID_TIPO_UNITA,"
            cmdMyCommand.CommandText += " SEZIONE, FOGLIO, PARTICELLA, ESTENSIONE_PARTICELLA, ID_TIPO_PARTICELLA, SUBALTERNO,"
            cmdMyCommand.CommandText += " INDIRIZZO, CIVICO, INTERNO, SCALA,"
            cmdMyCommand.CommandText += " ID_ASSENZA_DATI_CATASTALI)"

            cmdMyCommand.CommandText += " SELECT " & nFlusso & ","
            cmdMyCommand.CommandText += " RUOLO_TARSU.CODICE_ISTAT, RUOLO_TARSU.COD_TRIBUTO, RUOLO_TARSU.ANNO,"
            cmdMyCommand.CommandText += " ENTI_IN_LAVORAZIONE.COD_FISCALE, ENTI_IN_LAVORAZIONE.COGNOME, ENTI_IN_LAVORAZIONE.NOME, ENTI_IN_LAVORAZIONE.SESSO, ENTI_IN_LAVORAZIONE.DATA_NASCITA, ENTI_IN_LAVORAZIONE.COMUNE_NASCITA_SEDE, ENTI_IN_LAVORAZIONE.PV_NASCITA_SEDE,"
            cmdMyCommand.CommandText += " RUOLO_TARSU.COD_CONTRIBUENTE, RUOLO_TARSU.ID_RUOLO, CASE WHEN ANAGRAFICA.PARTITA_IVA IS NULL OR ANAGRAFICA.PARTITA_IVA='' THEN ANAGRAFICA.COD_FISCALE ELSE ANAGRAFICA.PARTITA_IVA END,"
            cmdMyCommand.CommandText += " SUBSTRING(ANAGRAFICA.COGNOME_DENOMINAZIONE,1,50), SUBSTRING(ANAGRAFICA.NOME,1,25), ANAGRAFICA.COMUNE_NASCITA, ANAGRAFICA.PROV_NASCITA,"
            cmdMyCommand.CommandText += " ENTI_IN_LAVORAZIONE.DESCRIZIONE, ENTI_IN_LAVORAZIONE.PROVINCIA, ENTI_IN_LAVORAZIONE.COD_BELFIORE,"
            cmdMyCommand.CommandText += " RUOLO_TARSU.DATA_INIZIO, RUOLO_TARSU.DATA_FINE,"
            cmdMyCommand.CommandText += " RUOLO_TARSU.ID_TITOLO_OCCUPAZIONE, RUOLO_TARSU.ID_NATURA_OCCUPANTE, RUOLO_TARSU.ID_DESTINAZIONE_USO, RUOLO_TARSU.ID_TIPO_UNITA,"
            cmdMyCommand.CommandText += " RUOLO_TARSU.SEZIONE, RUOLO_TARSU.FOGLIO, RUOLO_TARSU.PARTICELLA, RUOLO_TARSU.ESTENSIONE_PARTICELLA, RUOLO_TARSU.ID_TIPO_PARTICELLA, RUOLO_TARSU.SUBALTERNO,"
            cmdMyCommand.CommandText += " SUBSTRING(RUOLO_TARSU.UBICAZIONE,1,30), RUOLO_TARSU.CIVICO, RUOLO_TARSU.INTERNO, RUOLO_TARSU.SCALA,"
            cmdMyCommand.CommandText += " CASE WHEN (RUOLO_TARSU.ANNO>='2009' AND (RUOLO_TARSU.SEZIONE+RUOLO_TARSU.FOGLIO+RUOLO_TARSU.PARTICELLA IS NULL OR RUOLO_TARSU.SEZIONE+RUOLO_TARSU.FOGLIO+RUOLO_TARSU.PARTICELLA='')) THEN 3 ELSE "
            cmdMyCommand.CommandText += " CASE WHEN (FOGLIO<>'' AND PARTICELLA='') THEN 3 ELSE "
            cmdMyCommand.CommandText += " CASE WHEN (RUOLO_TARSU.FOGLIO+RUOLO_TARSU.PARTICELLA IS NULL OR RUOLO_TARSU.FOGLIO+RUOLO_TARSU.PARTICELLA='') THEN 3 ELSE NULL END END END"
            cmdMyCommand.CommandText += " FROM ANAGRAFICA"
            cmdMyCommand.CommandText += " INNER JOIN RUOLO_TARSU ON ANAGRAFICA.COD_CONTRIBUENTE = RUOLO_TARSU.COD_CONTRIBUENTE"
            cmdMyCommand.CommandText += " INNER JOIN ENTI_IN_LAVORAZIONE ON RUOLO_TARSU.CODICE_ISTAT=ENTI_IN_LAVORAZIONE.CODICE_ISTAT"
            cmdMyCommand.CommandText += " WHERE (ANAGRAFICA.DATA_FINE_VALIDITA IS NULL)"
            cmdMyCommand.CommandText += " AND ((CASE WHEN ANAGRAFICA.PARTITA_IVA IS NULL OR ANAGRAFICA.PARTITA_IVA='' THEN ANAGRAFICA.COD_FISCALE ELSE ANAGRAFICA.PARTITA_IVA END)<>'')"
            cmdMyCommand.CommandText += " AND (RUOLO_TARSU.CODICE_ISTAT=@CODICEISTAT)"
            cmdMyCommand.CommandText += " AND (RUOLO_TARSU.ANNO=@ANNORIF)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)
            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetDisposizione::" & Err.Message & vbCrLf & "SQL::" & cmdMyCommand.CommandText)
        End Try
        Return myResult
    End Function

    Public Function SetDisposizione(ByVal oDBManager As Utility.DBModel, ByVal oMyDati As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Integer
        Try
            myResult = -1
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandText = "INSERT INTO AE_DATI_FILE("
            'cmdMyCommand.CommandText += " ID_FLUSSO, CODICE_ISTAT, COD_TRIBUTO, ANNO,"
            'cmdMyCommand.CommandText += " COD_FISCALE_ENTE, COGNOME_ENTE, NOME_ENTE, SESSO_ENTE, DATA_NASCITA_ENTE, COMUNE_NASCITA_SEDE_ENTE, PV_NASCITA_SEDE_ENTE,"
            'cmdMyCommand.CommandText += " COD_CONTRIBUENTE, ID_RUOLO, COD_FISCALE,"
            'cmdMyCommand.CommandText += " COGNOME, NOME, SESSO, DATA_NASCITA, COMUNE_SEDE, PV_SEDE,"
            'cmdMyCommand.CommandText += " COMUNE_DOMICILIOFISC, PV_DOMICILIOFISC,"
            'cmdMyCommand.CommandText += " COMUNE_UBICAZIONE, PV_UBICAZIONE, COMUNE_UBICAZIONE_CATAST, COD_COMUNE_UBICAZIONE_CATAST,"
            'cmdMyCommand.CommandText += " ESTREMI_CONTRATTO, TIPOCONTRATTO, DATA_INIZIO, DATA_FINE,"
            'cmdMyCommand.CommandText += " ID_TITOLO_OCCUPAZIONE, ID_TIPO_OCCUPANTE, ID_TIPO_UTENZA, ID_DESTINAZIONE_USO, ID_TIPO_UNITA,"
            'cmdMyCommand.CommandText += " SEZIONE, FOGLIO, PARTICELLA, ESTENSIONE_PARTICELLA, ID_TIPO_PARTICELLA, SUBALTERNO,"
            'cmdMyCommand.CommandText += " INDIRIZZO, CIVICO, INTERNO, SCALA,"
            'cmdMyCommand.CommandText += " ID_ASSENZA_DATI_CATASTALI, MESI_FATTURAZIONE, SEGNO_SPESA, CONSUMO,IMPORTOFATTURATO)"

            'cmdMyCommand.CommandText += " VALUES (@IDFLUSSO, @CODISTAT, @TRIBUTO, @ANNO,"
            'cmdMyCommand.CommandText += "@CODFISCALEENTE, @COGNOMEENTE, @NOMEENTE, @SESSOENTE, @DATANASCITAENTE, @COMUNENASCITASEDEENTE, @PVNASCITASEDEENTE,"
            'cmdMyCommand.CommandText += "@IDCONTRIBUENTE, @IDCOLLEGAMENTO, @CODFISCALE,"
            'cmdMyCommand.CommandText += "@COGNOME, @NOME, @SESSO, @DATANASCITA, @COMUNENASCITASEDE, @PVNASCITASEDE,"
            'cmdMyCommand.CommandText += "@COMUNEDOMFISC, @PVDOMFISC,"
            'cmdMyCommand.CommandText += "@COMUNEAMMUBICAZIONE, @PVAMMUBICAZIONE, @COMUNECATASTUBICAZIONE, @CODCOMUNEUBICAZIONECATAST,"
            'cmdMyCommand.CommandText += "@ESTREMICONTRATTO, @TIPOCONTRATTO, @DATAINIZIO, @DATAFINE,"
            'cmdMyCommand.CommandText += "@IDTITOLOOCCUPAZIONE, @IDTITOLOOCCUPANTE, @IDTIPOUTENZA, @IDDESTINAZIONEUSO, @IDTIPOUNITA,"
            'cmdMyCommand.CommandText += "@SEZIONE, @FOGLIO, @PARTICELLA, @ESTENSIONEPARTICELLA, @IDTIPOPARTICELLA, @SUBALTERNO,"
            'cmdMyCommand.CommandText += "@INDIRIZZO, @CIVICO, @INTERNO, @SCALA,"
            'cmdMyCommand.CommandText += "@IDASSENZADATICATASTALI, @MESIFATTURAZIONE, @SEGNO, @CONSUMO,@IMPORTOFATTURATO)"
            ''Valorizzo i parameters:
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDFLUSSO", SqlDbType.Int)).Value = oMyDati.nIDFlusso
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODISTAT", SqlDbType.NVarChar)).Value = oMyDati.sCodISTAT
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TRIBUTO", SqlDbType.NVarChar)).Value = oMyDati.sTributo
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNO", SqlDbType.NVarChar)).Value = oMyDati.sAnno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODFISCALEENTE", SqlDbType.NVarChar)).Value = oMyDati.sCodFiscaleEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COGNOMEENTE", SqlDbType.NVarChar)).Value = oMyDati.sCognomeEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMEENTE", SqlDbType.NVarChar)).Value = oMyDati.sNomeEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SESSOENTE", SqlDbType.NVarChar)).Value = oMyDati.sSessoEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATANASCITAENTE", SqlDbType.NVarChar)).Value = oMyDati.sDataNascitaEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNENASCITASEDEENTE", SqlDbType.NVarChar)).Value = oMyDati.sComuneNascitaSedeEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVNASCITASEDEENTE", SqlDbType.NVarChar)).Value = oMyDati.sPVNascitaSedeEnte
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCONTRIBUENTE", SqlDbType.Int)).Value = oMyDati.nIDContribuente
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCOLLEGAMENTO", SqlDbType.Int)).Value = oMyDati.nIDCollegamento
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODFISCALE", SqlDbType.NVarChar)).Value = oMyDati.sCodFiscale
            'If oMyDati.sCognome.Length > 50 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COGNOME", SqlDbType.NVarChar)).Value = oMyDati.sCognome.Substring(0, 50)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COGNOME", SqlDbType.NVarChar)).Value = oMyDati.sCognome
            'End If
            'If oMyDati.sNome.Length > 25 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOME", SqlDbType.NVarChar)).Value = oMyDati.sNome.Substring(0, 25)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOME", SqlDbType.NVarChar)).Value = oMyDati.sNome
            'End If
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SESSO", SqlDbType.NVarChar)).Value = oMyDati.sSesso
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATANASCITA", SqlDbType.NVarChar)).Value = oMyDati.sDataNascita
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNENASCITASEDE", SqlDbType.NVarChar)).Value = oMyDati.sComuneNascitaSede
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVNASCITASEDE", SqlDbType.NVarChar)).Value = oMyDati.sPVNascitaSede
            'If oMyDati.sComuneDomFisc.Length > 20 Then
            '    oMyDati.sComuneDomFisc = oMyDati.sComuneDomFisc.Substring(0, 20)
            'End If
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNEDOMFISC", SqlDbType.NVarChar)).Value = oMyDati.sComuneDomFisc
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVDOMFISC", SqlDbType.NVarChar)).Value = oMyDati.sPVDomFisc
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNEAMMUBICAZIONE", SqlDbType.NVarChar)).Value = oMyDati.sComuneAmmUbicazione
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVAMMUBICAZIONE", SqlDbType.NVarChar)).Value = oMyDati.sPVAmmUbicazione
            ''valorizzare solo se diverso da comune amministrativo
            'If oMyDati.sComuneCatastUbicazione <> oMyDati.sComuneAmmUbicazione Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNECATASTUBICAZIONE", SqlDbType.NVarChar)).Value = oMyDati.sComuneCatastUbicazione
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNECATASTUBICAZIONE", SqlDbType.NVarChar)).Value = String.Empty
            'End If
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODCOMUNEUBICAZIONECATAST", SqlDbType.NVarChar)).Value = oMyDati.sCodComuneUbicazioneCatast
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ESTREMICONTRATTO", SqlDbType.NVarChar)).Value = oMyDati.sEstremiContratto
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TIPOCONTRATTO", SqlDbType.NVarChar)).Value = oMyDati.sTipoContratto
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAINIZIO", SqlDbType.NVarChar)).Value = oMyDati.sDataInizio
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAFINE", SqlDbType.NVarChar)).Value = oMyDati.sDataFine
            'If CInt(oMyDati.nIDTitoloOccupazione) <> -1 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTITOLOOCCUPAZIONE", SqlDbType.Int)).Value = oMyDati.nIDTitoloOccupazione
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTITOLOOCCUPAZIONE", SqlDbType.Int)).Value = DBNull.Value
            'End If
            'If CInt(oMyDati.nIDTipoOccupante) <> -1 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTITOLOOCCUPANTE", SqlDbType.Int)).Value = oMyDati.nIDTipoOccupante
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTITOLOOCCUPANTE", SqlDbType.Int)).Value = DBNull.Value
            'End If
            ''If CInt(oMyDati.nIDTipoUtenza) <> -1 Then
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTIPOUTENZA", SqlDbType.Int)).Value = oMyDati.nIDTipoUtenza
            ''Else
            ''    mySQLCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTIPOUTENZA", SqlDbType.Int)).Value = DBNull.Value
            ''End If
            'If CInt(oMyDati.nIDDestinazioneUso) <> -1 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDDESTINAZIONEUSO", SqlDbType.Int)).Value = oMyDati.nIDDestinazioneUso
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDDESTINAZIONEUSO", SqlDbType.Int)).Value = DBNull.Value
            'End If
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTIPOUNITA", SqlDbType.NVarChar)).Value = oMyDati.sIDTipoUnita
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SEZIONE", SqlDbType.NVarChar)).Value = oMyDati.sSezione
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@FOGLIO", SqlDbType.NVarChar)).Value = oMyDati.sFoglio
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PARTICELLA", SqlDbType.NVarChar)).Value = oMyDati.sParticella
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ESTENSIONEPARTICELLA", SqlDbType.NVarChar)).Value = oMyDati.sEstensioneParticella
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDTIPOPARTICELLA", SqlDbType.NVarChar)).Value = oMyDati.sIDTipoParticella
            'If oMyDati.sSubalterno.Length > 4 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SUBALTERNO", SqlDbType.NVarChar)).Value = oMyDati.sSubalterno.Substring(0, 4)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SUBALTERNO", SqlDbType.NVarChar)).Value = oMyDati.sSubalterno
            'End If
            'If oMyDati.sIndirizzo.Length > 30 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@INDIRIZZO", SqlDbType.NVarChar)).Value = oMyDati.sIndirizzo.Substring(0, 30)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@INDIRIZZO", SqlDbType.NVarChar)).Value = oMyDati.sIndirizzo
            'End If
            'If oMyDati.sCivico.Length > 6 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CIVICO", SqlDbType.NVarChar)).Value = oMyDati.sCivico.Substring(0, 6)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CIVICO", SqlDbType.NVarChar)).Value = oMyDati.sCivico
            'End If
            'If oMyDati.sInterno.Length > 2 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@INTERNO", SqlDbType.NVarChar)).Value = oMyDati.sInterno.Substring(0, 2)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@INTERNO", SqlDbType.NVarChar)).Value = oMyDati.sInterno
            'End If
            'If oMyDati.sScala.Length > 1 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SCALA", SqlDbType.NVarChar)).Value = oMyDati.sScala.Substring(0, 1)
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SCALA", SqlDbType.NVarChar)).Value = oMyDati.sScala
            'End If
            'If CInt(oMyDati.nIDAssenzaDatiCatastali) <> -1 Then
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDASSENZADATICATASTALI", SqlDbType.Int)).Value = oMyDati.nIDAssenzaDatiCatastali
            'Else
            '    cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDASSENZADATICATASTALI", SqlDbType.Int)).Value = DBNull.Value
            'End If
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@MESIFATTURAZIONE", SqlDbType.Int)).Value = oMyDati.nMesiFatturazione
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SEGNO", SqlDbType.NVarChar)).Value = oMyDati.sSegno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CONSUMO", SqlDbType.Int)).Value = oMyDati.nConsumo
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPORTOFATTURATO", SqlDbType.Float)).Value = oMyDati.nImportoFatturato

            'myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)

            If oMyDati.sCognome.Length > 50 Then
                oMyDati.sCognome = oMyDati.sCognome.Substring(0, 50)
            End If
            If oMyDati.sNome.Length > 25 Then
                oMyDati.sNome = oMyDati.sNome.Substring(0, 25)
            End If
            If oMyDati.sComuneDomFisc.Length > 20 Then
                oMyDati.sComuneDomFisc = oMyDati.sComuneDomFisc.Substring(0, 20)
            End If
            'valorizzare solo se diverso da comune amministrativo                                      
            If oMyDati.sComuneCatastUbicazione <> oMyDati.sComuneAmmUbicazione Then
                oMyDati.sComuneCatastUbicazione = oMyDati.sComuneCatastUbicazione
            Else
                oMyDati.sComuneCatastUbicazione = String.Empty
            End If
            If oMyDati.sSubalterno.Length > 4 Then
                oMyDati.sSubalterno = oMyDati.sSubalterno.Substring(0, 4)
            End If
            If oMyDati.sIndirizzo.Length > 30 Then
                oMyDati.sIndirizzo = oMyDati.sIndirizzo.Substring(0, 30)
            End If
            If oMyDati.sCivico.Length > 6 Then
                oMyDati.sCivico = oMyDati.sCivico.Substring(0, 6)
            End If
            If oMyDati.sInterno.Length > 2 Then
                oMyDati.sInterno = oMyDati.sInterno.Substring(0, 2)
            End If
            If oMyDati.sScala.Length > 1 Then
                oMyDati.sScala = oMyDati.sScala.Substring(0, 1)
            End If

            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_AE_DATI_FILE_IU", "IDFLUSSO", "CODISTAT", "TRIBUTO", "ANNO", "CODFISCALEENTE", "COGNOMEENTE", "NOMEENTE", "SESSOENTE", "DATANASCITAENTE", "COMUNENASCITASEDEENTE", "PVNASCITASEDEENTE", "IDCONTRIBUENTE", "IDCOLLEGAMENTO", "CODFISCALE", "COGNOME", "NOME", "SESSO", "DATANASCITA", "COMUNENASCITASEDE", "PVNASCITASEDE", "COMUNEDOMFISC", "PVDOMFISC", "COMUNEAMMUBICAZIONE", "PVAMMUBICAZIONE", "COMUNECATASTUBICAZIONE", "CODCOMUNEUBICAZIONECATAST", "ESTREMICONTRATTO", "TIPOCONTRATTO", "DATAINIZIO", "DATAFINE", "IDTITOLOOCCUPAZIONE", "IDTITOLOOCCUPANTE", "IDTIPOUTENZA", "IDDESTINAZIONEUSO", "IDTIPOUNITA", "SEZIONE", "FOGLIO", "PARTICELLA", "ESTENSIONEPARTICELLA", "IDTIPOPARTICELLA", "SUBALTERNO", "INDIRIZZO", "CIVICO", "INTERNO", "SCALA", "IDASSENZADATICATASTALI", "MESIFATTURAZIONE", "SEGNO", "CONSUMO", "IMPORTOFATTURATO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("IDFLUSSO", oMyDati.nIDFlusso) _
                , oDBManager.GetParam("CODISTAT", oMyDati.sCodISTAT) _
                , oDBManager.GetParam("TRIBUTO", oMyDati.sTributo) _
                , oDBManager.GetParam("ANNO", oMyDati.sAnno) _
                , oDBManager.GetParam("CODFISCALEENTE", oMyDati.sCodFiscaleEnte) _
                , oDBManager.GetParam("COGNOMEENTE", oMyDati.sCognomeEnte) _
                , oDBManager.GetParam("NOMEENTE", oMyDati.sNomeEnte) _
                , oDBManager.GetParam("SESSOENTE", oMyDati.sSessoEnte) _
                , oDBManager.GetParam("DATANASCITAENTE", oMyDati.sDataNascitaEnte) _
                , oDBManager.GetParam("COMUNENASCITASEDEENTE", oMyDati.sComuneNascitaSedeEnte) _
                , oDBManager.GetParam("PVNASCITASEDEENTE", oMyDati.sPVNascitaSedeEnte) _
                , oDBManager.GetParam("IDCONTRIBUENTE", oMyDati.nIDContribuente) _
                , oDBManager.GetParam("IDCOLLEGAMENTO", oMyDati.nIDCollegamento) _
                , oDBManager.GetParam("CODFISCALE", oMyDati.sCodFiscale) _
                , oDBManager.GetParam("COGNOME", oMyDati.sCognome) _
                , oDBManager.GetParam("NOME", oMyDati.sNome) _
                , oDBManager.GetParam("SESSO", oMyDati.sSesso) _
                , oDBManager.GetParam("DATANASCITA", oMyDati.sDataNascita) _
                , oDBManager.GetParam("COMUNENASCITASEDE", oMyDati.sComuneNascitaSede) _
                , oDBManager.GetParam("PVNASCITASEDE", oMyDati.sPVNascitaSede) _
                , oDBManager.GetParam("COMUNEDOMFISC", oMyDati.sComuneDomFisc) _
                , oDBManager.GetParam("PVDOMFISC", oMyDati.sPVDomFisc) _
                , oDBManager.GetParam("COMUNEAMMUBICAZIONE", oMyDati.sComuneAmmUbicazione) _
                , oDBManager.GetParam("PVAMMUBICAZIONE", oMyDati.sPVAmmUbicazione) _
                , oDBManager.GetParam("COMUNECATASTUBICAZIONE", oMyDati.sComuneCatastUbicazione) _
                , oDBManager.GetParam("CODCOMUNEUBICAZIONECATAST", oMyDati.sCodComuneUbicazioneCatast) _
                , oDBManager.GetParam("ESTREMICONTRATTO", oMyDati.sEstremiContratto) _
                , oDBManager.GetParam("TIPOCONTRATTO", oMyDati.sTipoContratto) _
                , oDBManager.GetParam("DATAINIZIO", oMyDati.sDataInizio) _
                , oDBManager.GetParam("DATAFINE", oMyDati.sDataFine) _
                , oDBManager.GetParam("IDTITOLOOCCUPAZIONE", oMyDati.nIDTitoloOccupazione) _
                , oDBManager.GetParam("IDTITOLOOCCUPANTE", oMyDati.nIDTipoOccupante) _
                , oDBManager.GetParam("IDTIPOUTENZA", oMyDati.nIDTipoUtenza) _
                , oDBManager.GetParam("IDDESTINAZIONEUSO", oMyDati.nIDDestinazioneUso) _
                , oDBManager.GetParam("IDTIPOUNITA", oMyDati.sIDTipoUnita) _
                , oDBManager.GetParam("SEZIONE", oMyDati.sSezione) _
                , oDBManager.GetParam("FOGLIO", oMyDati.sFoglio) _
                , oDBManager.GetParam("PARTICELLA", oMyDati.sParticella) _
                , oDBManager.GetParam("ESTENSIONEPARTICELLA", oMyDati.sEstensioneParticella) _
                , oDBManager.GetParam("IDTIPOPARTICELLA", oMyDati.sIDTipoParticella) _
                , oDBManager.GetParam("SUBALTERNO", oMyDati.sSubalterno) _
                , oDBManager.GetParam("INDIRIZZO", oMyDati.sIndirizzo) _
                , oDBManager.GetParam("CIVICO", oMyDati.sCivico) _
                , oDBManager.GetParam("INTERNO", oMyDati.sInterno) _
                , oDBManager.GetParam("SCALA", oMyDati.sScala) _
                , oDBManager.GetParam("IDASSENZADATICATASTALI", oMyDati.nIDAssenzaDatiCatastali) _
                , oDBManager.GetParam("MESIFATTURAZIONE", oMyDati.nMesiFatturazione) _
                , oDBManager.GetParam("SEGNO", oMyDati.sSegno) _
                , oDBManager.GetParam("CONSUMO", oMyDati.nConsumo) _
                , oDBManager.GetParam("IMPORTOFATTURATO", oMyDati.nImportoFatturato)
            )
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next
            myResult = 1
        Catch Err As Exception
            'Dim sSQL As String
            'Dim x As Integer
            'For x = 0 To cmdMyCommand.Parameters.Count - 1
            '    sSQL += cmdMyCommand.Parameters(x).Value & ","
            'Next
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetDisposizione::" & Err.Message) ' & vbCrLf & "SQL::" & cmdMyCommand.CommandText & vbCrLf & " VALUES " & sSQL
        End Try
        Return myResult
    End Function

    Public Function SetFlussiAcq(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal oMyDati As ImportTestaICI) As Integer
        Dim drDati As SqlClient.SqlDataReader

        Try
            myResult = -1
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "INSERT INTO AE_FLUSSI_ACQUISITI_ICI (COD_FLUSSO_PAGAMENTI, NOME_FILE, DATA_CREAZIONE,"
            cmdMyCommand.CommandText += " DATA_INIZIO, DATA_FINE, COD_DIVISA, N_PAGAMENTI, ANNO, CODICE_ISTAT,"
            cmdMyCommand.CommandText += " TOTALE_IMPORTI_POSITIVI, TOTALE_IMPORTI_NEGATIVI, TOTALE_RIVERSATO,"
            cmdMyCommand.CommandText += " TOTALE_IMPORTI_SANZIONI, NUMERO_SANZIONI, DATA_IMPORTAZIONE)"
            cmdMyCommand.CommandText += " VALUES (@CODFLUSSO, @NOMEFLUSSO,"
            cmdMyCommand.CommandText += "@DATACREAZIONEFLUSSO, @DATAINIZIO, @DATAFINE, @DIVISA,"
            cmdMyCommand.CommandText += "@TOTPAGAMENTIACQ, @ANNOFLUSSO, @CODISTAT,"
            cmdMyCommand.CommandText += "@TOTIMPPOS, @TOTIMPNEG, @TOTALERIVERSATO, @TOTALEIMPSANZIONI, @TOTSANZIONI, @DATAIMPORTAZIONE)"
            cmdMyCommand.CommandText += " SELECT @@IDENTITY"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODFLUSSO", SqlDbType.Int)).Value = oMyDati.nCodFlussoPagamenti
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMEFLUSSO", SqlDbType.NVarChar)).Value = oMyDati.sNomeFlusso
            If oMyDati.sDataCreazioneFlusso <> "00000000" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATACREAZIONEFLUSSO", SqlDbType.NVarChar)).Value = oMyDati.sDataCreazioneFlusso
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATACREAZIONEFLUSSO", SqlDbType.NVarChar)).Value = DBNull.Value
            End If
            If oMyDati.sDataInizio <> "00000000" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAINIZIO", SqlDbType.NVarChar)).Value = oMyDati.sDataInizio
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAINIZIO", SqlDbType.NVarChar)).Value = DBNull.Value
            End If
            If oMyDati.sDataFine <> "00000000" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAFINE", SqlDbType.NVarChar)).Value = oMyDati.sDataFine
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAFINE", SqlDbType.NVarChar)).Value = DBNull.Value
            End If
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DIVISA", SqlDbType.NVarChar)).Value = oMyDati.sDivisa
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTPAGAMENTIACQ", SqlDbType.Float)).Value = oMyDati.nTotPagamentiAcq
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNOFLUSSO", SqlDbType.NVarChar)).Value = oMyDati.sAnnoFlusso
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODISTAT", SqlDbType.NVarChar)).Value = oMyDati.sCodISTAT
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPPOS", SqlDbType.Float)).Value = oMyDati.nTotImpPos
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPNEG", SqlDbType.Float)).Value = oMyDati.nTotImpNeg
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTALERIVERSATO", SqlDbType.Float)).Value = oMyDati.nTotaleRiversato
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTALEIMPSANZIONI", SqlDbType.Float)).Value = oMyDati.nTotaleImpSanzioni
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTSANZIONI", SqlDbType.Int)).Value = oMyDati.nTotSanzioni
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAIMPORTAZIONE", SqlDbType.DateTime)).Value = Now
            'eseguo la query
            drDati = oDBManager.GetDataReader(cmdMyCommand.CommandText)
            Do While drDati.Read
                myResult = drDati(0)
            Loop
            drDati.Close()
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetFlussiAcq::" & Err.Message & vbCrLf & "SQL::" & cmdMyCommand.CommandText)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::SetFlussiAcq::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function SetFlussiAcq(ByVal oDBManager As Utility.DBModel, ByVal oMyDati As ImportCodaICI) As Integer
        Try
            myResult = -1
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "UPDATE AE_FLUSSI_ACQUISITI_ICI SET"
            cmdMyCommand.CommandText += " TOT_ANAGRAFICHE=@TOTANAGRAFICHE,"
            cmdMyCommand.CommandText += " TOT_DISPOSIZIONI=@TOTDISPOSIZIONI,"
            cmdMyCommand.CommandText += " TOT_IMPABIPRIN=@TOTIMPABIPRIN,"
            cmdMyCommand.CommandText += " TOT_IMPALTRIFAB=@TOTIMPALTRIFAB,"
            cmdMyCommand.CommandText += " TOT_IMPAREEFAB=@TOTIMPAREEFAB,"
            cmdMyCommand.CommandText += " TOT_IMPDETRAZIONE=@TOTIMPDETRAZIONE,"
            cmdMyCommand.CommandText += " TOT_IMPTERAGR=@TOTIMPTERAGR,"
            cmdMyCommand.CommandText += " TOT_IMPVERSAMENTI=@TOTIMPVERSAMENTI,"
            cmdMyCommand.CommandText += " TOT_IMPVIOLAZIONI=@TOTIMPVIOLAZIONI,"
            cmdMyCommand.CommandText += " TOT_VERSAMENTI=@TOTVERSAMENTI,"
            cmdMyCommand.CommandText += " TOT_VIOLAZIONI=@TOTVIOLAZIONI"
            cmdMyCommand.CommandText += " WHERE (AE_FLUSSI_ACQUISITI_ICI.COD_FLUSSO_PAGAMENTI=@CODFLUSSO)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTANAGRAFICHE", SqlDbType.Int)).Value = oMyDati.nTotAnagrafiche
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTDISPOSIZIONI", SqlDbType.Int)).Value = oMyDati.nTotDisposizioni
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPABIPRIN", SqlDbType.Float)).Value = oMyDati.nTotImpAbiPrin
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPALTRIFAB", SqlDbType.Float)).Value = oMyDati.nTotImpAltriFab
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPAREEFAB", SqlDbType.Float)).Value = oMyDati.nTotImpAreeFab
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPDETRAZIONE", SqlDbType.Float)).Value = oMyDati.nTotImpDetrazione
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPTERAGR", SqlDbType.Float)).Value = oMyDati.nTotImpTerAgr
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPVERSAMENTI", SqlDbType.Float)).Value = oMyDati.nTotImpVersamenti
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTIMPVIOLAZIONI", SqlDbType.Float)).Value = oMyDati.nTotImpViolazioni
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTVERSAMENTI", SqlDbType.Int)).Value = oMyDati.nTotVersamenti
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TOTVIOLAZIONI", SqlDbType.Int)).Value = oMyDati.nTotViolazioni
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODFLUSSO", SqlDbType.Int)).Value = oMyDati.nCodFlussoPagamenti
            oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)

            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetFlussiAcq::" & Err.Message & vbCrLf & "SQL::" & cmdMyCommand.CommandText)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::SetFlussiAcq::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function SetAnagrafe(ByVal oDBManager As Utility.DBModel, ByVal oMyDati As ImportAnagICI) As Integer
        Try
            myResult = -1
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "INSERT INTO AE_ANAGRAFICA_ICI (CODICE_ISTAT, COD_CONTRIBUENTE, CF_PIVA, COGNOME, NOME, SESSO,"
            cmdMyCommand.CommandText += " DATA_NASCITA, COMUNE_NASCITA, PROV_NASCITA, NAZIONALITA,"
            cmdMyCommand.CommandText += " VIA_RES, FRAZIONE_RES, CIVICO_RES, CAP_RES, CITTA_RES, PROVINCIA_RES,"
            cmdMyCommand.CommandText += " NOMINATIVO_RCP, VIA_RCP, CIVICO_RCP, CAP_RCP, CITTA_RCP, PROVINCIA_RCP) "
            cmdMyCommand.CommandText += " VALUES (@CODISTAT, @IDCONTRIBUENTE, @CFPIVA, @COGNOME, @NOME, @SESSO,"
            cmdMyCommand.CommandText += "@DATANASCITA, @COMUNENASCITA, @PVNASCITA, @NAZIONALITA,"
            cmdMyCommand.CommandText += "@VIARES, @FRAZIONERES, @CIVICORES, @CAPRES, @CITTARES, @PVRES,"
            cmdMyCommand.CommandText += "@NOMINATIVOINVIO, @VIAINVIO, @CIVICOINVIO, @CAPINVIO, @CITTAINVIO, @PVINVIO)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODISTAT", SqlDbType.NVarChar)).Value = oMyDati.sCodISTAT
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCONTRIBUENTE", SqlDbType.Int)).Value = oMyDati.nIdContribuente
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CFPIVA", SqlDbType.NVarChar)).Value = oMyDati.sCFPIVA
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COGNOME", SqlDbType.NVarChar)).Value = oMyDati.sCognome
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOME", SqlDbType.NVarChar)).Value = oMyDati.sNome
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SESSO", SqlDbType.NVarChar)).Value = oMyDati.sSesso
            If oMyDati.sDataNascita = "00000000" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATANASCITA", SqlDbType.NVarChar)).Value = String.Empty
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATANASCITA", SqlDbType.NVarChar)).Value = oMyDati.sDataNascita
            End If
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@COMUNENASCITA", SqlDbType.NVarChar)).Value = oMyDati.sComuneNascita
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVNASCITA", SqlDbType.NVarChar)).Value = oMyDati.sPVNascita
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NAZIONALITA", SqlDbType.NVarChar)).Value = oMyDati.sNazionalita
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@VIARES", SqlDbType.NVarChar)).Value = oMyDati.sViaRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@FRAZIONERES", SqlDbType.NVarChar)).Value = oMyDati.sFrazioneRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CIVICORES", SqlDbType.NVarChar)).Value = oMyDati.sCivicoRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CAPRES", SqlDbType.NVarChar)).Value = oMyDati.sCAPRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CITTARES", SqlDbType.NVarChar)).Value = oMyDati.sCittaRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVRES", SqlDbType.NVarChar)).Value = oMyDati.sPVRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMINATIVOINVIO", SqlDbType.NVarChar)).Value = oMyDati.sNominativoInvio
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@VIAINVIO", SqlDbType.NVarChar)).Value = oMyDati.sViaInvio
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CIVICOINVIO", SqlDbType.NVarChar)).Value = oMyDati.sCivicoInvio
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CAPINVIO", SqlDbType.NVarChar)).Value = oMyDati.sCAPInvio
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CITTAINVIO", SqlDbType.NVarChar)).Value = oMyDati.sCittaInvio
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PVINVIO", SqlDbType.NVarChar)).Value = oMyDati.sPVInvio

            myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)
            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetAnagrafe::" & Err.Message & vbCrLf & "SQL::" & cmdMyCommand.CommandText)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::SetAnagrafe::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function SetDisposizioneICI(ByVal oDBManager As Utility.DBModel, ByVal oMyDati As ImportDisposizioneICI) As Integer
        Try
            myResult = -1
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandText = "INSERT INTO AE_PAGAMENTI_ICI (CODICE_ISTAT, ID_VERSAMENTO, COD_FLUSSO, CF_PIVA, COGNOME_NOME, DATA_ACCREDITO, DATA_PAGAMENTO, FLAG_ACCONTO_SALDO, ANNO, ANNO_RIFERIMENTO, N_FAB,"
            cmdMyCommand.CommandText += " IMPORTO, IMP_TER_AGR, IMP_AREE_FAB, IMP_ALTRI_FAB, IMP_ABI_PRIN, DETRAZIONE, INDIRIZZO_RES, CAP_RES, CITTA_RES, BOLLETTINO_EX_RURALE,"
            cmdMyCommand.CommandText += " DATA_SANZIONE, N_SANZIONE, N_MOVIMENTO, SPAZIO_LIBERO, DATA_FLUSSO_RENDICONTAZIONE, COD_CONTRIBUENTE, COD_CONTRIBUENTE_SIMILE, FLAG_TRATTATO,"
            cmdMyCommand.CommandText += " COD_DIVISA, FLAG_RAVVEDIMENTO_OPEROSO, COD_FLUSSO_AP, PROGRESSIVO_PAGAMENTO_AP, COD_TIPO_PAGAMENTO, CODICE_COMUNICO, IMPORTO_PAGATO_CONTRIBUENTE,"
            cmdMyCommand.CommandText += " NOME_IMMAGINE, PROVENIENZA, IMMAGINE_VISIBLE, N_PROG_RENDICONTAZIONE, NUMERO_QUIETANZA, TIPO_BOLLETTINO_VIOLAZIONI)"
            cmdMyCommand.CommandText += " VALUES (@CODISTAT, @IDVERSAMENTO, @CODFLUSSO, @CFPIVA, @NOMINATIVO,"
            cmdMyCommand.CommandText += "@DATAACCREDITO, @DATAPAGAMENTO, @FLAGAS, @ANNO, @ANNORIF, @NUMFAB,"
            cmdMyCommand.CommandText += "@IMPVERSAMENTO, @IMPTERAGR, @IMPAREEFAB, @IMPALTRIFAB, @IMPABIPRINC, @IMPDETRAZIONE,"
            cmdMyCommand.CommandText += "@INDIRIZZORES, @CAPRES, @CITTARES, @BOLLETTINOEXRURALE, @DATASANZIONE,"
            cmdMyCommand.CommandText += "@NUMSANZIONE, @NUMMOVIMENTO, @SPAZIOLIBERO, @DATAFLUSSOREND,"
            cmdMyCommand.CommandText += "@IDCONTRIBUENTE, @IDCONTRIBUENTESIMILE, @FLAGTRATTATO, @DIVISA, @FLAGRAVVOPEROSO,"
            cmdMyCommand.CommandText += "@CODFLUSSOAP, @PROGRPAGAP, @CODTIPOPAGAMENTO, @CODICECOMUNICO, @IMPVERSATO,"
            cmdMyCommand.CommandText += "@NOMEIMMAGINE, @PROVENIENZA, @VIEWIMMAGINE, @PROGRENDICONTAZ, @NUMEROQUIETANZA, @TIPOBOLLETTINOVIOLAZIONI)"
            'Valorizzo i parameters:
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODISTAT", SqlDbType.NVarChar)).Value = oMyDati.sCodISTAT
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDVERSAMENTO", SqlDbType.Int)).Value = oMyDati.nIdVersamento
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODFLUSSO", SqlDbType.Int)).Value = oMyDati.nCodFlusso
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CFPIVA", SqlDbType.NVarChar)).Value = oMyDati.sCFPIVA
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMINATIVO", SqlDbType.NVarChar)).Value = oMyDati.sNominativo
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAACCREDITO", SqlDbType.NVarChar)).Value = oMyDati.sDataAccredito
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAPAGAMENTO", SqlDbType.NVarChar)).Value = oMyDati.sDataPagamento
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@FLAGAS", SqlDbType.NVarChar)).Value = oMyDati.sFlagAS
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNO", SqlDbType.NVarChar)).Value = oMyDati.sAnno
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = oMyDati.sAnnoRif
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NUMFAB", SqlDbType.Int)).Value = oMyDati.nNumFab
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPVERSAMENTO", SqlDbType.Float)).Value = oMyDati.nImpVersamento
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPTERAGR", SqlDbType.Float)).Value = oMyDati.nImpTerAgr
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPAREEFAB", SqlDbType.Float)).Value = oMyDati.nImpAreeFab
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPALTRIFAB", SqlDbType.Float)).Value = oMyDati.nImpAltriFab
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPABIPRINC", SqlDbType.Float)).Value = oMyDati.nImpAbiPrinc
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPDETRAZIONE", SqlDbType.Float)).Value = oMyDati.nImpDetrazione
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@INDIRIZZORES", SqlDbType.NVarChar)).Value = oMyDati.sIndirizzoRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CAPRES", SqlDbType.NVarChar)).Value = oMyDati.sCapRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CITTARES", SqlDbType.NVarChar)).Value = oMyDati.sCittaRes
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@BOLLETTINOEXRURALE", SqlDbType.NVarChar)).Value = oMyDati.sBollettinoEXRurale
            If oMyDati.sDataSanzione <> "00000000" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATASANZIONE", SqlDbType.NVarChar)).Value = oMyDati.sDataSanzione
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATASANZIONE", SqlDbType.NVarChar)).Value = String.Empty
            End If
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NUMSANZIONE", SqlDbType.NVarChar)).Value = oMyDati.sNumSanzione
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NUMMOVIMENTO", SqlDbType.NVarChar)).Value = oMyDati.sNumMovimento
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@SPAZIOLIBERO", SqlDbType.NVarChar)).Value = oMyDati.sSpazioLibero
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DATAFLUSSOREND", SqlDbType.NVarChar)).Value = oMyDati.sDataFlussoRend
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCONTRIBUENTE", SqlDbType.Int)).Value = oMyDati.nIdContribuente
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCONTRIBUENTESIMILE", SqlDbType.Int)).Value = oMyDati.nIdContribuenteSimile
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@FLAGTRATTATO", SqlDbType.Int)).Value = oMyDati.nFlagTrattato
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@DIVISA", SqlDbType.NVarChar)).Value = oMyDati.sDivisa
            If oMyDati.sFlagRavvOperoso <> "" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@FLAGRAVVOPEROSO", SqlDbType.NVarChar)).Value = oMyDati.sFlagRavvOperoso
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@FLAGRAVVOPEROSO", SqlDbType.NVarChar)).Value = String.Empty
            End If
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODFLUSSOAP", SqlDbType.Int)).Value = oMyDati.nCodFlussoAP
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PROGRPAGAP", SqlDbType.Int)).Value = oMyDati.nProgrPagAP
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTIPOPAGAMENTO", SqlDbType.NVarChar)).Value = oMyDati.sCodTipoPagamento
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICECOMUNICO", SqlDbType.NVarChar)).Value = oMyDati.sCodiceComunico
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IMPVERSATO", SqlDbType.Float)).Value = oMyDati.nImpVersato
            If oMyDati.sNomeImmagine <> "" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMEIMMAGINE", SqlDbType.NVarChar)).Value = oMyDati.sNomeImmagine
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NOMEIMMAGINE", SqlDbType.NVarChar)).Value = String.Empty
            End If
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PROVENIENZA", SqlDbType.NVarChar)).Value = oMyDati.sProvenienza
            If oMyDati.sViewImmagine = "" Then
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@VIEWIMMAGINE", SqlDbType.NVarChar)).Value = "1"
            Else
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@VIEWIMMAGINE", SqlDbType.NVarChar)).Value = oMyDati.sViewImmagine
            End If
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@PROGRENDICONTAZ", SqlDbType.Int)).Value = oMyDati.nProgRendicontaz
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@NUMEROQUIETANZA", SqlDbType.Int)).Value = oMyDati.nNumQuietanza
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@TIPOBOLLETTINOVIOLAZIONI", SqlDbType.NVarChar)).Value = oMyDati.sTipoBollettinoViolazioni


            myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)
            myResult = 1
        Catch Err As Exception
            Dim FncGen As New General
            Dim sValParametri As String = FncGen.GetValParamCmd(cmdMyCommand)
            Log.Debug("Si è verificato un errore in GestDatiOPENae::SetDisposizioneICI::" & vbCrLf & "SQL::" & cmdMyCommand.CommandText & vbCrLf & " VALUES " & sValParametri)
            Log.Warn("Si è verificato un errore in GestDatiOPENae::SetDisposizioneICI::" & Err.Message)
        End Try
        Return myResult
    End Function
#End Region

#Region "Query di CANCELLAZIONE"
    Public Function DeleteDisposizione(ByVal oDBManager As Utility.DBModel, ByVal sTributo As String, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
        Try
            myResult = -1
            Dim dvResult As DataView
            'cmdMyCommand = New SqlClient.SqlCommand
            'mySQLCommand.CommandType = CommandType.Text
            'mySQLCommand.CommandText = "DELETE"
            'mySQLCommand.CommandText += " FROM AE_DATI_FILE"
            'mySQLCommand.CommandText += " WHERE (AE_DATI_FILE.CODICE_ISTAT=@CODICEISTAT)"
            'mySQLCommand.CommandText += " AND (AE_DATI_FILE.ANNO=@ANNORIF)"
            'mySQLCommand.CommandText += " AND (AE_DATI_FILE.COD_TRIBUTO=@CODTRIBUTO)"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_AE_DATI_FILE_D", "CODICEISTAT", "ANNORIF", "CODTRIBUTO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo))
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::DeleteDisposizione::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function DeleteDisposizione(ByVal oDBManager As Utility.DBModel, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
        Try
            myResult = -1
            'svuoto l'anagrafica
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandType = CommandType.Text
            cmdMyCommand.CommandText = "DELETE"
            cmdMyCommand.CommandText += " FROM AE_ANAGRAFICA_ICI"
            cmdMyCommand.CommandText += " WHERE (AE_ANAGRAFICA_ICI.COD_CONTRIBUENTE IN("
            cmdMyCommand.CommandText += " SELECT AE_PAGAMENTI_ICI.COD_CONTRIBUENTE"
            cmdMyCommand.CommandText += " FROM AE_PAGAMENTI_ICI"
            cmdMyCommand.CommandText += " WHERE (AE_PAGAMENTI_ICI.CODICE_ISTAT=@CODICEISTAT)"
            cmdMyCommand.CommandText += " AND (AE_PAGAMENTI_ICI.ANNO_RIFERIMENTO=@ANNORIF)))"
            cmdMyCommand.CommandText += " AND (AE_ANAGRAFICA_ICI.CODICE_ISTAT=@CODICEISTAT)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)

            'svuoto le disposizioni
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandType = CommandType.Text
            cmdMyCommand.CommandText = "DELETE"
            cmdMyCommand.CommandText += " FROM AE_PAGAMENTI_ICI"
            cmdMyCommand.CommandText += " WHERE (AE_PAGAMENTI_ICI.CODICE_ISTAT=@CODICEISTAT)"
            cmdMyCommand.CommandText += " AND (AE_PAGAMENTI_ICI.ANNO_RIFERIMENTO=@ANNORIF)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)

            'svuoto i flussi
            cmdMyCommand = New SqlClient.SqlCommand
            cmdMyCommand.CommandType = CommandType.Text
            cmdMyCommand.CommandText = "DELETE"
            cmdMyCommand.CommandText += " FROM AE_FLUSSI_ACQUISITI_ICI"
            cmdMyCommand.CommandText += " WHERE (AE_FLUSSI_ACQUISITI_ICI.CODICE_ISTAT=@CODICEISTAT)"
            cmdMyCommand.CommandText += " AND (AE_FLUSSI_ACQUISITI_ICI.ANNO=@ANNORIF)"
            'valorizzo i parameters
            cmdMyCommand.Parameters.Clear()
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)
            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::DeleteDisposizione::" & Err.Message)
        End Try
        Return myResult
    End Function

    Public Function DeleteFlusso(ByVal oDBManager As Utility.DBModel, ByVal sTributo As String, ByVal sCodIstat As String, ByVal sAnno As String) As Integer
        Try
            myResult = -1
            'cmdMyCommand = New SqlClient.SqlCommand
            'cmdMyCommand.CommandType = CommandType.Text
            'cmdMyCommand.CommandText = "DELETE"
            'cmdMyCommand.CommandText += " FROM AE_FLUSSI_ESTRATTI"
            'cmdMyCommand.CommandText += " WHERE (AE_FLUSSI_ESTRATTI.CODICE_ISTAT=@CodiceISTAT)"
            'cmdMyCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.ANNO=@AnnoRif)"
            'cmdMyCommand.CommandText += " AND (AE_FLUSSI_ESTRATTI.COD_TRIBUTO=@CODTRIBUTO)"
            ''valorizzo i parameters
            'cmdMyCommand.Parameters.Clear()
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODICEISTAT", SqlDbType.NVarChar)).Value = sCodIstat
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@ANNORIF", SqlDbType.NVarChar)).Value = sAnno
            'cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@CODTRIBUTO", SqlDbType.NVarChar)).Value = sTributo
            'myResult = oDBManager.ExecuteNonQuery(cmdMyCommand.CommandText)
            Dim dvResult As DataView
            Dim sSQL As String = oDBManager.GetSQL(Utility.DBModel.TypeQuery.StoredProcedure, "prc_AE_FLUSSI_ESTRATTI_D", "CODICEISTAT", "ANNORIF", "CODTRIBUTO")
            dvResult = oDBManager.GetDataView(sSQL, "TBL", oDBManager.GetParam("CODICEISTAT", sCodIstat), oDBManager.GetParam("ANNORIF", sAnno), oDBManager.GetParam("CODTRIBUTO", sTributo))
            For Each myRow As DataRowView In dvResult
                myResult = myRow(0)
            Next
            myResult = 1
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GestDatiOPENae::DeleteFlusso::" & Err.Message)
        End Try
        Return myResult
    End Function
#End Region
End Class
