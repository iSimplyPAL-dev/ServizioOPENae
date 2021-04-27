Imports log4net
Imports OPENUtility

Namespace OPENgov_AgenziaEntrate
    Public Class General
        Inherits LevelDB

        Private Log As ILog = LogManager.GetLogger(GetType(General))

        'Public Function LoadComboDati(ByVal sTypeDati As String, ByVal sTributo As String, ByVal WFSessione As CreateSessione) As DataView
        '    Dim DvDati As DataView
        '    'Dim FncDB As New LevelDB

        '    Try
        '        Select Case sTypeDati
        '            Case "TIT_OCCUPAZIONE"
        '                'Se sto caricando “TIT_OCCUPAZIONE”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTitoloOccupazione();
        '                DvDati = GetTitoloOccupazione(WFSessione, sTributo)
        '            Case "NAT_OCCUPAZIONE"
        '                'Se sto caricando “NAT_OCCUPAZIONE”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetNaturaOccupazione();
        '                DvDati = GetNaturaOccupazione(WFSessione)
        '            Case "TIPO_UTENZA"
        '                'Se sto caricando “TIPO_UTENZA”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTipoUtenza();
        '                DvDati = GetTIPOUTENZA(WFSessione)
        '            Case "DEST_USO"
        '                'Se sto caricando “DEST_USO”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetDestUso();
        '                DvDati = GetDestUso(WFSessione)
        '            Case "TIPO_UNITA"
        '                'Se sto caricando “TIPO_UNITA”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTipoUnita();
        '                DvDati = GetTipoUnita(WFSessione)
        '            Case "TIPO_PARTICELLA"
        '                'Se sto caricando “TIPO_PARTICELLA”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTipoParticella();
        '                DvDati = GetTipoParticella(WFSessione)
        '            Case "ASSENZA_DATI_CAT"
        '                'Se sto caricando “ASSENZA_DATI_CAT”:
        '                'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetAssenzaDatiCat();
        '                DvDati = GetAssenzaDatiCat(WFSessione, sTributo)
        '        End Select
        '        Return DvDati
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in AgenziaEntrate_General::LoadComboDati::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in AgenziaEntrate_General::LoadComboDati::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function
        Public Function LoadComboDati(ByVal sTypeDati As String, ByVal sTributo As String, ByVal myStringConnection As String) As DataView
            Dim DvDati As New DataView
            'Dim FncDB As New LevelDB

            Try
                Select Case sTypeDati
                    Case "TIT_OCCUPAZIONE"
                        'Se sto caricando “TIT_OCCUPAZIONE”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTitoloOccupazione();
                        DvDati = GetTitoloOccupazione(myStringConnection, sTributo)
                    Case "NAT_OCCUPAZIONE"
                        'Se sto caricando “NAT_OCCUPAZIONE”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetNaturaOccupazione();
                        DvDati = GetNaturaOccupazione(myStringConnection)
                    Case "TIPO_UTENZA"
                        'Se sto caricando “TIPO_UTENZA”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTipoUtenza();
                        DvDati = GetTipoUtenza(myStringConnection)
                    Case "DEST_USO"
                        'Se sto caricando “DEST_USO”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetDestUso();
                        DvDati = GetDestUso(myStringConnection)
                    Case "TIPO_UNITA"
                        'Se sto caricando “TIPO_UNITA”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTipoUnita();
                        DvDati = GetTipoUnita(myStringConnection)
                    Case "TIPO_PARTICELLA"
                        'Se sto caricando “TIPO_PARTICELLA”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetTipoParticella();
                        DvDati = GetTipoParticella(myStringConnection)
                    Case "ASSENZA_DATI_CAT"
                        'Se sto caricando “ASSENZA_DATI_CAT”:
                        'popolo il dataview con gli anni presenti a ruolo richiamando la funzione GetAssenzaDatiCat();
                        DvDati = GetAssenzaDatiCat(myStringConnection, sTributo)
                End Select
                Return DvDati
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in AgenziaEntrate_General::LoadComboDati::" & Err.Message)
                Return Nothing
            End Try
        End Function
    End Class
End Namespace