Public Class FrmEstrazRDB
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
	Friend WithEvents Label16 As System.Windows.Forms.Label
	Friend WithEvents TxtDal As System.Windows.Forms.TextBox
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents TxtAnno As System.Windows.Forms.TextBox
	Friend WithEvents Label13 As System.Windows.Forms.Label
	Friend WithEvents TxtConnessioneDB As System.Windows.Forms.TextBox
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents TxtFile As System.Windows.Forms.TextBox
	Friend WithEvents LblAvanzamento As System.Windows.Forms.Label
	Friend WithEvents CmdEstrazione As System.Windows.Forms.Button
	Friend WithEvents CmdUscita As System.Windows.Forms.Button
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents TxtCodIstat As System.Windows.Forms.TextBox
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents TxtCodFiscale As System.Windows.Forms.TextBox
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents TxtDenominazione As System.Windows.Forms.TextBox
	Friend WithEvents Label6 As System.Windows.Forms.Label
	Friend WithEvents TxtComune As System.Windows.Forms.TextBox
	Friend WithEvents Label7 As System.Windows.Forms.Label
	Friend WithEvents TxtPv As System.Windows.Forms.TextBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmEstrazRDB))
		Me.Label16 = New System.Windows.Forms.Label
		Me.TxtDal = New System.Windows.Forms.TextBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.TxtAnno = New System.Windows.Forms.TextBox
		Me.Label13 = New System.Windows.Forms.Label
		Me.TxtConnessioneDB = New System.Windows.Forms.TextBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.TxtFile = New System.Windows.Forms.TextBox
		Me.LblAvanzamento = New System.Windows.Forms.Label
		Me.CmdEstrazione = New System.Windows.Forms.Button
		Me.CmdUscita = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.TxtCodIstat = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.TxtCodFiscale = New System.Windows.Forms.TextBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.TxtDenominazione = New System.Windows.Forms.TextBox
		Me.Label6 = New System.Windows.Forms.Label
		Me.TxtComune = New System.Windows.Forms.TextBox
		Me.Label7 = New System.Windows.Forms.Label
		Me.TxtPv = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		'
		'Label16
		'
		Me.Label16.BackColor = System.Drawing.Color.Transparent
		Me.Label16.Location = New System.Drawing.Point(152, 112)
		Me.Label16.Name = "Label16"
		Me.Label16.Size = New System.Drawing.Size(40, 20)
		Me.Label16.TabIndex = 35
		Me.Label16.Text = "Dal"
		'
		'TxtDal
		'
		Me.TxtDal.Location = New System.Drawing.Point(192, 112)
		Me.TxtDal.Name = "TxtDal"
		Me.TxtDal.Size = New System.Drawing.Size(80, 20)
		Me.TxtDal.TabIndex = 34
		Me.TxtDal.Text = ""
		'
		'Label4
		'
		Me.Label4.BackColor = System.Drawing.Color.Transparent
		Me.Label4.Location = New System.Drawing.Point(48, 112)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(40, 20)
		Me.Label4.TabIndex = 33
		Me.Label4.Text = "Anno"
		'
		'TxtAnno
		'
		Me.TxtAnno.Location = New System.Drawing.Point(88, 112)
		Me.TxtAnno.Name = "TxtAnno"
		Me.TxtAnno.Size = New System.Drawing.Size(56, 20)
		Me.TxtAnno.TabIndex = 32
		Me.TxtAnno.Text = ""
		'
		'Label13
		'
		Me.Label13.BackColor = System.Drawing.Color.Transparent
		Me.Label13.Location = New System.Drawing.Point(16, 184)
		Me.Label13.Name = "Label13"
		Me.Label13.Size = New System.Drawing.Size(72, 20)
		Me.Label13.TabIndex = 39
		Me.Label13.Text = "Percorso DB"
		'
		'TxtConnessioneDB
		'
		Me.TxtConnessioneDB.Location = New System.Drawing.Point(88, 176)
		Me.TxtConnessioneDB.Name = "TxtConnessioneDB"
		Me.TxtConnessioneDB.Size = New System.Drawing.Size(328, 20)
		Me.TxtConnessioneDB.TabIndex = 38
		Me.TxtConnessioneDB.Text = "E:\pal\rdb\DATABASE\rdb.mdb"
		'
		'Label5
		'
		Me.Label5.BackColor = System.Drawing.Color.Transparent
		Me.Label5.Location = New System.Drawing.Point(16, 144)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(72, 20)
		Me.Label5.TabIndex = 37
		Me.Label5.Text = "Percorso File"
		'
		'TxtFile
		'
		Me.TxtFile.Location = New System.Drawing.Point(88, 144)
		Me.TxtFile.Name = "TxtFile"
		Me.TxtFile.Size = New System.Drawing.Size(328, 20)
		Me.TxtFile.TabIndex = 36
		Me.TxtFile.Text = "C:\RDB\"
		'
		'LblAvanzamento
		'
		Me.LblAvanzamento.AutoSize = True
		Me.LblAvanzamento.BackColor = System.Drawing.Color.Transparent
		Me.LblAvanzamento.Location = New System.Drawing.Point(24, 224)
		Me.LblAvanzamento.Name = "LblAvanzamento"
		Me.LblAvanzamento.Size = New System.Drawing.Size(200, 16)
		Me.LblAvanzamento.TabIndex = 40
		Me.LblAvanzamento.Text = "Estrazione in corso...Attendere prego..."
		Me.LblAvanzamento.Visible = False
		'
		'CmdEstrazione
		'
		Me.CmdEstrazione.Image = CType(resources.GetObject("CmdEstrazione.Image"), System.Drawing.Image)
		Me.CmdEstrazione.Location = New System.Drawing.Point(320, 208)
		Me.CmdEstrazione.Name = "CmdEstrazione"
		Me.CmdEstrazione.Size = New System.Drawing.Size(40, 40)
		Me.CmdEstrazione.TabIndex = 42
		'
		'CmdUscita
		'
		Me.CmdUscita.BackColor = System.Drawing.Color.Transparent
		Me.CmdUscita.Image = CType(resources.GetObject("CmdUscita.Image"), System.Drawing.Image)
		Me.CmdUscita.Location = New System.Drawing.Point(376, 208)
		Me.CmdUscita.Name = "CmdUscita"
		Me.CmdUscita.Size = New System.Drawing.Size(40, 40)
		Me.CmdUscita.TabIndex = 41
		'
		'Label1
		'
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Location = New System.Drawing.Point(32, 16)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(56, 20)
		Me.Label1.TabIndex = 44
		Me.Label1.Text = "Cod.Istat"
		'
		'TxtCodIstat
		'
		Me.TxtCodIstat.Location = New System.Drawing.Point(88, 16)
		Me.TxtCodIstat.Name = "TxtCodIstat"
		Me.TxtCodIstat.Size = New System.Drawing.Size(72, 20)
		Me.TxtCodIstat.TabIndex = 43
		Me.TxtCodIstat.Text = ""
		'
		'Label2
		'
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Location = New System.Drawing.Point(168, 16)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(96, 20)
		Me.Label2.TabIndex = 46
		Me.Label2.Text = "Cod.Fiscale Ente"
		'
		'TxtCodFiscale
		'
		Me.TxtCodFiscale.Location = New System.Drawing.Point(264, 16)
		Me.TxtCodFiscale.Name = "TxtCodFiscale"
		Me.TxtCodFiscale.Size = New System.Drawing.Size(152, 20)
		Me.TxtCodFiscale.TabIndex = 45
		Me.TxtCodFiscale.Text = ""
		'
		'Label3
		'
		Me.Label3.BackColor = System.Drawing.Color.Transparent
		Me.Label3.Location = New System.Drawing.Point(5, 48)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(88, 20)
		Me.Label3.TabIndex = 48
		Me.Label3.Text = "Denominazione"
		'
		'TxtDenominazione
		'
		Me.TxtDenominazione.Location = New System.Drawing.Point(88, 48)
		Me.TxtDenominazione.Name = "TxtDenominazione"
		Me.TxtDenominazione.Size = New System.Drawing.Size(328, 20)
		Me.TxtDenominazione.TabIndex = 47
		Me.TxtDenominazione.Text = ""
		'
		'Label6
		'
		Me.Label6.BackColor = System.Drawing.Color.Transparent
		Me.Label6.Location = New System.Drawing.Point(32, 80)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(48, 20)
		Me.Label6.TabIndex = 50
		Me.Label6.Text = "Comune"
		'
		'TxtComune
		'
		Me.TxtComune.Location = New System.Drawing.Point(88, 80)
		Me.TxtComune.Name = "TxtComune"
		Me.TxtComune.Size = New System.Drawing.Size(216, 20)
		Me.TxtComune.TabIndex = 49
		Me.TxtComune.Text = ""
		'
		'Label7
		'
		Me.Label7.BackColor = System.Drawing.Color.Transparent
		Me.Label7.Location = New System.Drawing.Point(320, 80)
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(40, 20)
		Me.Label7.TabIndex = 52
		Me.Label7.Text = "Prov."
		'
		'TxtPv
		'
		Me.TxtPv.Location = New System.Drawing.Point(360, 80)
		Me.TxtPv.Name = "TxtPv"
		Me.TxtPv.Size = New System.Drawing.Size(56, 20)
		Me.TxtPv.TabIndex = 51
		Me.TxtPv.Text = ""
		'
		'FrmEstrazRDB
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(432, 262)
		Me.Controls.Add(Me.TxtDenominazione)
		Me.Controls.Add(Me.Label7)
		Me.Controls.Add(Me.TxtPv)
		Me.Controls.Add(Me.Label6)
		Me.Controls.Add(Me.TxtComune)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.TxtCodFiscale)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.TxtCodIstat)
		Me.Controls.Add(Me.CmdEstrazione)
		Me.Controls.Add(Me.CmdUscita)
		Me.Controls.Add(Me.LblAvanzamento)
		Me.Controls.Add(Me.Label13)
		Me.Controls.Add(Me.TxtConnessioneDB)
		Me.Controls.Add(Me.Label5)
		Me.Controls.Add(Me.TxtFile)
		Me.Controls.Add(Me.Label16)
		Me.Controls.Add(Me.TxtDal)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.TxtAnno)
		Me.Name = "FrmEstrazRDB"
		Me.Text = "FrmEstrazRDB"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub CmdEstrazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEstrazione.Click
		Dim MsgCaption As String = "Agenzia Entrate"
		Dim FncAE As New EstrazRDB
		Dim oAE As New OPENae_InterfVB6VBNET.ClsCall
		Dim sNameFile, sProvenienza As String

		Try
			LblAvanzamento.Visible = True
			If IsNumeric(TxtAnno.Text) = False Then
				MessageBox.Show("Attenzione! E' necessario selezionare un anno a ruolo!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
				Exit Sub
			End If
			Try
				TxtDal.Text = CDate(TxtDal.Text)
			Catch Err As Exception
				MessageBox.Show("Attenzione! E' necessario selezionare una data valida!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
				Exit Sub
			End Try
			If TxtConnessioneDB.Text = "" Or TxtConnessioneDB.Text.ToLower.IndexOf(".mdb") <= 0 Then
				MessageBox.Show("Attenzione! E' necessario selezionare un database!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
				Exit Sub
			End If
			If TxtFile.Text = "" Then
				MessageBox.Show("Attenzione! E' necessario selezionare un percorso dove scaricare il file!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
				Exit Sub
			End If
			If TxtCodIstat.Text = "" Then
				MessageBox.Show("Attenzione! E' necessario inserire il Codice Istat!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
				Exit Sub
			End If
			TxtConnessioneDB.Text = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + TxtConnessioneDB.Text

			LblAvanzamento.Text = "Preparazione banca dati in corso...Attendere prego..."
			If FncAE.PreparazioneDati(TxtAnno.Text, TxtDal.Text, TxtCodIstat.Text, txtCodFiscale.text, txtdenominazione.text, txtcomune.text, txtpv.text, TxtConnessioneDB.Text) = False Then
				MessageBox.Show("Errore nella preparazione della banca dati!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
				Exit Sub
			End If

			LblAvanzamento.Text = "Estrazione tracciato in corso...Attendere prego..."
			'If TxtTributo.Text = "8852" Then
			'	sNameFile = oTest.CreaFlussoMEF(TxtEnte.Text, TxtBelfiore.Text, TxtDescrEnte.Text, TxtCAPEnte.Text, TxtTributo.Text, TxtAnno.Text, TxtDataScadenza.Text, TxtProgInvio.Text, TxtFile.Text, TxtFileUpload.Text, TxtFile.Text.Replace(".txt", ".log"), "R", TxtFileDownload.Text)
			'Else
			sNameFile = oAE.CallServiceAE(TxtConnessioneDB.Text, TxtFile.Text, "LOG_CALLSERVICEAE.log", TxtFile.Text)
			'End If
			If sNameFile <> "" Then
				MessageBox.Show("Errore in estrazione tracciato!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
			Else
				MessageBox.Show("Estrazione terminata con successo!", MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
			End If
		Catch ex As Exception
			MessageBox.Show("Errore in estrazione tracciato!" + vbCrLf + ex.Message, MsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
		Finally
			LblAvanzamento.Visible = False
		End Try
	End Sub

	Private Sub CmdUscita_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdUscita.Click
		End
	End Sub
End Class
