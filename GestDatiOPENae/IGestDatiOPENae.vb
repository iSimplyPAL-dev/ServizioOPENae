Public Interface IGestDatiOPENae
#Region "Importazione Dati"
    Function PopolaTabAppoggioAE(ByVal sTributo As String, ByVal sCodiceISTAT As String, ByVal sAnnoRif As String, ByVal sFileImport As String, ByVal sProvenienza As String) As Boolean
    Function PopolaTabAppoggioAE(ByVal oDati() As AgenziaEntrateDLL.AgenziaEntrate.DisposizioneAE) As Boolean
#End Region

#Region "Estrazione Dati"
    Function EstraiTracciato(ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sCodiceISTAT As String, ByRef sNomeFileTracciati As String) As String
    Function EstraiTracciato(ByVal sCodiceISTAT As String, ByVal sCodBelfiore As String, ByVal sDescrEnte As String, ByVal sCAPEnte As String, ByVal sTributo As String, ByVal sAnnoRif As String, ByVal sDataScadenza As String, ByVal nProgInvio As Integer, ByRef sNomeFileTracciati As String) As String
#End Region

#Region "Preleva Dati"
    Function GetFlussiTracciati(ByVal sTributo As String, ByVal sCodiceISTAT As String) As AgenziaEntrateDLL.AgenziaEntrate.objFlussoAE()
#End Region
End Interface
