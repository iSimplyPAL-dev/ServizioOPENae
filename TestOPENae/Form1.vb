Imports log4net
Imports Utility

Public Class Form1
	Inherits System.Windows.Forms.Form
	Private Shared Log As ILog = LogManager.GetLogger("Form1")

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
	Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents TxtTributo As System.Windows.Forms.TextBox
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents TxtEnte As System.Windows.Forms.TextBox
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents TxtAnno As System.Windows.Forms.TextBox
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents TxtFile As System.Windows.Forms.TextBox
	Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents CmdPopolaDaFile As System.Windows.Forms.Button
	Friend WithEvents Label6 As System.Windows.Forms.Label
	Friend WithEvents CmdCreaFile As System.Windows.Forms.Button
	Friend WithEvents Label7 As System.Windows.Forms.Label
	Friend WithEvents TxtCAPEnte As System.Windows.Forms.TextBox
	Friend WithEvents Label8 As System.Windows.Forms.Label
	Friend WithEvents TxtDescrEnte As System.Windows.Forms.TextBox
	Friend WithEvents Label9 As System.Windows.Forms.Label
	Friend WithEvents TxtBelfiore As System.Windows.Forms.TextBox
	Friend WithEvents Label10 As System.Windows.Forms.Label
	Friend WithEvents TxtProgInvio As System.Windows.Forms.TextBox
	Friend WithEvents Label11 As System.Windows.Forms.Label
	Friend WithEvents TxtDataScadenza As System.Windows.Forms.TextBox
	Friend WithEvents Label12 As System.Windows.Forms.Label
	Friend WithEvents CmdInterfaccia As System.Windows.Forms.Button
	Friend WithEvents OptRendicontazione As System.Windows.Forms.RadioButton
	Friend WithEvents OptRDB As System.Windows.Forms.RadioButton
	Friend WithEvents Label13 As System.Windows.Forms.Label
	Friend WithEvents TxtConnessioneDB As System.Windows.Forms.TextBox
	Friend WithEvents Label14 As System.Windows.Forms.Label
	Friend WithEvents Label15 As System.Windows.Forms.Label
	Friend WithEvents TxtFileUpload As System.Windows.Forms.TextBox
	Friend WithEvents TxtFileDownload As System.Windows.Forms.TextBox
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents TxtDal As System.Windows.Forms.TextBox
	Friend WithEvents Label16 As System.Windows.Forms.Label
	Friend WithEvents Label17 As System.Windows.Forms.Label
	Friend WithEvents TxtNomeFile As System.Windows.Forms.TextBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox
		Me.Label17 = New System.Windows.Forms.Label
		Me.TxtNomeFile = New System.Windows.Forms.TextBox
		Me.Label16 = New System.Windows.Forms.Label
		Me.TxtDal = New System.Windows.Forms.TextBox
		Me.Label15 = New System.Windows.Forms.Label
		Me.TxtFileDownload = New System.Windows.Forms.TextBox
		Me.Label14 = New System.Windows.Forms.Label
		Me.TxtFileUpload = New System.Windows.Forms.TextBox
		Me.Label13 = New System.Windows.Forms.Label
		Me.TxtConnessioneDB = New System.Windows.Forms.TextBox
		Me.OptRDB = New System.Windows.Forms.RadioButton
		Me.OptRendicontazione = New System.Windows.Forms.RadioButton
		Me.TxtDataScadenza = New System.Windows.Forms.TextBox
		Me.Label10 = New System.Windows.Forms.Label
		Me.TxtProgInvio = New System.Windows.Forms.TextBox
		Me.Label11 = New System.Windows.Forms.Label
		Me.TxtCAPEnte = New System.Windows.Forms.TextBox
		Me.TxtDescrEnte = New System.Windows.Forms.TextBox
		Me.TxtBelfiore = New System.Windows.Forms.TextBox
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.TxtFile = New System.Windows.Forms.TextBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.TxtAnno = New System.Windows.Forms.TextBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.TxtEnte = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.TxtTributo = New System.Windows.Forms.TextBox
		Me.GroupBox2 = New System.Windows.Forms.GroupBox
		Me.Button1 = New System.Windows.Forms.Button
		Me.Label12 = New System.Windows.Forms.Label
		Me.CmdInterfaccia = New System.Windows.Forms.Button
		Me.Label6 = New System.Windows.Forms.Label
		Me.CmdCreaFile = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.CmdPopolaDaFile = New System.Windows.Forms.Button
		Me.GroupBox1.SuspendLayout()
		Me.GroupBox2.SuspendLayout()
		Me.SuspendLayout()
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.Label17)
		Me.GroupBox1.Controls.Add(Me.TxtNomeFile)
		Me.GroupBox1.Controls.Add(Me.Label16)
		Me.GroupBox1.Controls.Add(Me.TxtDal)
		Me.GroupBox1.Controls.Add(Me.Label15)
		Me.GroupBox1.Controls.Add(Me.TxtFileDownload)
		Me.GroupBox1.Controls.Add(Me.Label14)
		Me.GroupBox1.Controls.Add(Me.TxtFileUpload)
		Me.GroupBox1.Controls.Add(Me.Label13)
		Me.GroupBox1.Controls.Add(Me.TxtConnessioneDB)
		Me.GroupBox1.Controls.Add(Me.OptRDB)
		Me.GroupBox1.Controls.Add(Me.OptRendicontazione)
		Me.GroupBox1.Controls.Add(Me.TxtDataScadenza)
		Me.GroupBox1.Controls.Add(Me.Label10)
		Me.GroupBox1.Controls.Add(Me.TxtProgInvio)
		Me.GroupBox1.Controls.Add(Me.Label11)
		Me.GroupBox1.Controls.Add(Me.TxtCAPEnte)
		Me.GroupBox1.Controls.Add(Me.TxtDescrEnte)
		Me.GroupBox1.Controls.Add(Me.TxtBelfiore)
		Me.GroupBox1.Controls.Add(Me.Label7)
		Me.GroupBox1.Controls.Add(Me.Label8)
		Me.GroupBox1.Controls.Add(Me.Label9)
		Me.GroupBox1.Controls.Add(Me.Label5)
		Me.GroupBox1.Controls.Add(Me.TxtFile)
		Me.GroupBox1.Controls.Add(Me.Label4)
		Me.GroupBox1.Controls.Add(Me.TxtAnno)
		Me.GroupBox1.Controls.Add(Me.Label3)
		Me.GroupBox1.Controls.Add(Me.TxtEnte)
		Me.GroupBox1.Controls.Add(Me.Label2)
		Me.GroupBox1.Controls.Add(Me.TxtTributo)
		Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(560, 256)
		Me.GroupBox1.TabIndex = 4
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Parametri"
		'
		'Label17
		'
		Me.Label17.BackColor = System.Drawing.Color.Transparent
		Me.Label17.Location = New System.Drawing.Point(8, 145)
		Me.Label17.Name = "Label17"
		Me.Label17.Size = New System.Drawing.Size(56, 20)
		Me.Label17.TabIndex = 33
		Me.Label17.Text = "Nome File"
		'
		'TxtNomeFile
		'
		Me.TxtNomeFile.Location = New System.Drawing.Point(72, 137)
		Me.TxtNomeFile.Name = "TxtNomeFile"
		Me.TxtNomeFile.Size = New System.Drawing.Size(480, 20)
		Me.TxtNomeFile.TabIndex = 32
		Me.TxtNomeFile.Text = "MEF_PAGAMENTI_ICI_2_04625_2012_0000001_20121031_133324.txt"
		'
		'Label16
		'
		Me.Label16.BackColor = System.Drawing.Color.Transparent
		Me.Label16.Location = New System.Drawing.Point(432, 24)
		Me.Label16.Name = "Label16"
		Me.Label16.Size = New System.Drawing.Size(40, 20)
		Me.Label16.TabIndex = 31
		Me.Label16.Text = "Dal"
		'
		'TxtDal
		'
		Me.TxtDal.Location = New System.Drawing.Point(472, 24)
		Me.TxtDal.Name = "TxtDal"
		Me.TxtDal.Size = New System.Drawing.Size(80, 20)
		Me.TxtDal.TabIndex = 30
		Me.TxtDal.Text = ""
		'
		'Label15
		'
		Me.Label15.BackColor = System.Drawing.Color.Transparent
		Me.Label15.Location = New System.Drawing.Point(8, 232)
		Me.Label15.Name = "Label15"
		Me.Label15.Size = New System.Drawing.Size(56, 20)
		Me.Label15.TabIndex = 29
		Me.Label15.Text = "Download"
		'
		'TxtFileDownload
		'
		Me.TxtFileDownload.Location = New System.Drawing.Point(72, 224)
		Me.TxtFileDownload.Name = "TxtFileDownload"
		Me.TxtFileDownload.Size = New System.Drawing.Size(480, 20)
		Me.TxtFileDownload.TabIndex = 28
		Me.TxtFileDownload.Text = "c:\"
		'
		'Label14
		'
		Me.Label14.BackColor = System.Drawing.Color.Transparent
		Me.Label14.Location = New System.Drawing.Point(8, 200)
		Me.Label14.Name = "Label14"
		Me.Label14.Size = New System.Drawing.Size(40, 20)
		Me.Label14.TabIndex = 27
		Me.Label14.Text = "Upload"
		'
		'TxtFileUpload
		'
		Me.TxtFileUpload.Location = New System.Drawing.Point(72, 192)
		Me.TxtFileUpload.Name = "TxtFileUpload"
		Me.TxtFileUpload.Size = New System.Drawing.Size(480, 20)
		Me.TxtFileUpload.TabIndex = 26
		Me.TxtFileUpload.Text = "https://sec.isimply.it/openae/ElaborazioneTracciati/importdati/"
		'
		'Label13
		'
		Me.Label13.BackColor = System.Drawing.Color.Transparent
		Me.Label13.Location = New System.Drawing.Point(8, 119)
		Me.Label13.Name = "Label13"
		Me.Label13.Size = New System.Drawing.Size(40, 20)
		Me.Label13.TabIndex = 25
		Me.Label13.Text = "DB"
		'
		'TxtConnessioneDB
		'
		Me.TxtConnessioneDB.Enabled = False
		Me.TxtConnessioneDB.Location = New System.Drawing.Point(72, 111)
		Me.TxtConnessioneDB.Name = "TxtConnessioneDB"
		Me.TxtConnessioneDB.Size = New System.Drawing.Size(480, 20)
		Me.TxtConnessioneDB.TabIndex = 24
		Me.TxtConnessioneDB.Text = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\RDB\D" & _
		"ATABASE\rdb.mdb"
		'
		'OptRDB
		'
		Me.OptRDB.Location = New System.Drawing.Point(432, 86)
		Me.OptRDB.Name = "OptRDB"
		Me.OptRDB.Size = New System.Drawing.Size(72, 16)
		Me.OptRDB.TabIndex = 23
		Me.OptRDB.Text = "RDB"
		'
		'OptRendicontazione
		'
		Me.OptRendicontazione.Checked = True
		Me.OptRendicontazione.Location = New System.Drawing.Point(280, 86)
		Me.OptRendicontazione.Name = "OptRendicontazione"
		Me.OptRendicontazione.Size = New System.Drawing.Size(136, 16)
		Me.OptRendicontazione.TabIndex = 22
		Me.OptRendicontazione.TabStop = True
		Me.OptRendicontazione.Text = "OPENrendicontazione"
		'
		'TxtDataScadenza
		'
		Me.TxtDataScadenza.Location = New System.Drawing.Point(64, 84)
		Me.TxtDataScadenza.Name = "TxtDataScadenza"
		Me.TxtDataScadenza.Size = New System.Drawing.Size(64, 20)
		Me.TxtDataScadenza.TabIndex = 18
		Me.TxtDataScadenza.Text = "20131031"
		'
		'Label10
		'
		Me.Label10.BackColor = System.Drawing.Color.Transparent
		Me.Label10.Location = New System.Drawing.Point(128, 88)
		Me.Label10.Name = "Label10"
		Me.Label10.Size = New System.Drawing.Size(56, 20)
		Me.Label10.TabIndex = 21
		Me.Label10.Text = "Prov.Invio"
		'
		'TxtProgInvio
		'
		Me.TxtProgInvio.Location = New System.Drawing.Point(184, 84)
		Me.TxtProgInvio.Name = "TxtProgInvio"
		Me.TxtProgInvio.Size = New System.Drawing.Size(80, 20)
		Me.TxtProgInvio.TabIndex = 20
		Me.TxtProgInvio.Text = "1"
		'
		'Label11
		'
		Me.Label11.BackColor = System.Drawing.Color.Transparent
		Me.Label11.Location = New System.Drawing.Point(8, 88)
		Me.Label11.Name = "Label11"
		Me.Label11.Size = New System.Drawing.Size(56, 20)
		Me.Label11.TabIndex = 19
		Me.Label11.Text = "Scadenza"
		'
		'TxtCAPEnte
		'
		Me.TxtCAPEnte.Location = New System.Drawing.Point(504, 54)
		Me.TxtCAPEnte.Name = "TxtCAPEnte"
		Me.TxtCAPEnte.Size = New System.Drawing.Size(48, 20)
		Me.TxtCAPEnte.TabIndex = 16
		Me.TxtCAPEnte.Text = "20015"
		'
		'TxtDescrEnte
		'
		Me.TxtDescrEnte.Location = New System.Drawing.Point(128, 54)
		Me.TxtDescrEnte.Name = "TxtDescrEnte"
		Me.TxtDescrEnte.Size = New System.Drawing.Size(336, 20)
		Me.TxtDescrEnte.TabIndex = 14
		Me.TxtDescrEnte.Text = "PARABIAGO"
		'
		'TxtBelfiore
		'
		Me.TxtBelfiore.Location = New System.Drawing.Point(48, 54)
		Me.TxtBelfiore.Name = "TxtBelfiore"
		Me.TxtBelfiore.Size = New System.Drawing.Size(40, 20)
		Me.TxtBelfiore.TabIndex = 12
		Me.TxtBelfiore.Text = "G324"
		'
		'Label7
		'
		Me.Label7.BackColor = System.Drawing.Color.Transparent
		Me.Label7.Location = New System.Drawing.Point(472, 59)
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(40, 20)
		Me.Label7.TabIndex = 17
		Me.Label7.Text = "CAP"
		'
		'Label8
		'
		Me.Label8.BackColor = System.Drawing.Color.Transparent
		Me.Label8.Location = New System.Drawing.Point(96, 59)
		Me.Label8.Name = "Label8"
		Me.Label8.Size = New System.Drawing.Size(40, 20)
		Me.Label8.TabIndex = 15
		Me.Label8.Text = "Ente"
		'
		'Label9
		'
		Me.Label9.BackColor = System.Drawing.Color.Transparent
		Me.Label9.Location = New System.Drawing.Point(8, 59)
		Me.Label9.Name = "Label9"
		Me.Label9.Size = New System.Drawing.Size(48, 20)
		Me.Label9.TabIndex = 13
		Me.Label9.Text = "Belfiore"
		'
		'Label5
		'
		Me.Label5.BackColor = System.Drawing.Color.Transparent
		Me.Label5.Location = New System.Drawing.Point(8, 168)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(64, 20)
		Me.Label5.TabIndex = 11
		Me.Label5.Text = "Path File"
		'
		'TxtFile
		'
		Me.TxtFile.Location = New System.Drawing.Point(72, 164)
		Me.TxtFile.Name = "TxtFile"
		Me.TxtFile.Size = New System.Drawing.Size(480, 20)
		Me.TxtFile.TabIndex = 10
		Me.TxtFile.Text = "\\192.168.14.112\f$\Siti-Rend\Rendicontazione_Ente\estrazione_dati\ici\ICI_MEF\"
		'
		'Label4
		'
		Me.Label4.BackColor = System.Drawing.Color.Transparent
		Me.Label4.Location = New System.Drawing.Point(288, 28)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(40, 20)
		Me.Label4.TabIndex = 9
		Me.Label4.Text = "Anno"
		'
		'TxtAnno
		'
		Me.TxtAnno.Location = New System.Drawing.Point(328, 24)
		Me.TxtAnno.Name = "TxtAnno"
		Me.TxtAnno.Size = New System.Drawing.Size(80, 20)
		Me.TxtAnno.TabIndex = 8
		Me.TxtAnno.Text = "2013"
		'
		'Label3
		'
		Me.Label3.BackColor = System.Drawing.Color.Transparent
		Me.Label3.Location = New System.Drawing.Point(160, 28)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(40, 20)
		Me.Label3.TabIndex = 7
		Me.Label3.Text = "Ente"
		'
		'TxtEnte
		'
		Me.TxtEnte.Location = New System.Drawing.Point(200, 24)
		Me.TxtEnte.Name = "TxtEnte"
		Me.TxtEnte.Size = New System.Drawing.Size(80, 20)
		Me.TxtEnte.TabIndex = 6
		Me.TxtEnte.Text = "015168"
		'
		'Label2
		'
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Location = New System.Drawing.Point(8, 28)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(40, 20)
		Me.Label2.TabIndex = 5
		Me.Label2.Text = "Tributo"
		'
		'TxtTributo
		'
		Me.TxtTributo.Location = New System.Drawing.Point(48, 24)
		Me.TxtTributo.Name = "TxtTributo"
		Me.TxtTributo.Size = New System.Drawing.Size(80, 20)
		Me.TxtTributo.TabIndex = 4
		Me.TxtTributo.Text = "8852"
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.Button1)
		Me.GroupBox2.Controls.Add(Me.Label12)
		Me.GroupBox2.Controls.Add(Me.CmdInterfaccia)
		Me.GroupBox2.Controls.Add(Me.Label6)
		Me.GroupBox2.Controls.Add(Me.CmdCreaFile)
		Me.GroupBox2.Controls.Add(Me.Label1)
		Me.GroupBox2.Controls.Add(Me.CmdPopolaDaFile)
		Me.GroupBox2.Location = New System.Drawing.Point(8, 280)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(424, 104)
		Me.GroupBox2.TabIndex = 5
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Funzioni"
		'
		'Button1
		'
		Me.Button1.Enabled = False
		Me.Button1.Location = New System.Drawing.Point(202, 42)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(142, 20)
		Me.Button1.TabIndex = 8
		Me.Button1.Text = "PREPARA RDB"
		'
		'Label12
		'
		Me.Label12.BackColor = System.Drawing.Color.Transparent
		Me.Label12.Location = New System.Drawing.Point(8, 72)
		Me.Label12.Name = "Label12"
		Me.Label12.Size = New System.Drawing.Size(232, 20)
		Me.Label12.TabIndex = 7
		Me.Label12.Text = "Interfaccia VB6-VB.NET"
		'
		'CmdInterfaccia
		'
		Me.CmdInterfaccia.Location = New System.Drawing.Point(388, 72)
		Me.CmdInterfaccia.Name = "CmdInterfaccia"
		Me.CmdInterfaccia.Size = New System.Drawing.Size(20, 20)
		Me.CmdInterfaccia.TabIndex = 6
		'
		'Label6
		'
		Me.Label6.BackColor = System.Drawing.Color.Transparent
		Me.Label6.Location = New System.Drawing.Point(8, 46)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(232, 20)
		Me.Label6.TabIndex = 5
		Me.Label6.Text = "Export File"
		'
		'CmdCreaFile
		'
		Me.CmdCreaFile.Enabled = False
		Me.CmdCreaFile.Location = New System.Drawing.Point(388, 46)
		Me.CmdCreaFile.Name = "CmdCreaFile"
		Me.CmdCreaFile.Size = New System.Drawing.Size(20, 20)
		Me.CmdCreaFile.TabIndex = 4
		'
		'Label1
		'
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Location = New System.Drawing.Point(8, 22)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(232, 20)
		Me.Label1.TabIndex = 3
		Me.Label1.Text = "Import da File"
		'
		'CmdPopolaDaFile
		'
		Me.CmdPopolaDaFile.Enabled = False
		Me.CmdPopolaDaFile.Location = New System.Drawing.Point(388, 22)
		Me.CmdPopolaDaFile.Name = "CmdPopolaDaFile"
		Me.CmdPopolaDaFile.Size = New System.Drawing.Size(20, 20)
		Me.CmdPopolaDaFile.TabIndex = 2
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.Color.Snow
		Me.ClientSize = New System.Drawing.Size(576, 398)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.GroupBox1)
		Me.Name = "Form1"
		Me.Text = "Form1"
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox2.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub CmdPopolaDaFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdPopolaDaFile.Click
		Try
			Dim oTest As New ServiceOPENae.ServiceOPENae
			Dim sProvenienza As String
			Dim sFile As String = TxtFile.Text + TxtNomeFile.Text

			If OptRendicontazione.Checked = True Then
				sProvenienza = "O"
			Else
				sProvenienza = "R"
            End If
            oTest.Url = System.Configuration.ConfigurationSettings.AppSettings("OPENae_InterfVB6VBNET.ServiceOPENae.ServiceOPENae")
			oTest.Timeout = 1800000
			If oTest.PopolaTabAppoggioAE(TxtTributo.Text, TxtAnno.Text, TxtEnte.Text, sFile, sProvenienza) = False Then
				MessageBox.Show("Errore", "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
			Else
				MessageBox.Show("Finito", "EVVIVA", MessageBoxButtons.OK, MessageBoxIcon.Information)
			End If
		Catch Err As Exception
			MessageBox.Show(Err.Message, "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
	End Sub

	Private Sub CmdCreaFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCreaFile.Click
		Try
			Dim oTest As New ServiceOPENae.ServiceOPENae
			Dim sFile As String = TxtFile.Text + TxtNomeFile.Text

            oTest.Url = System.Configuration.ConfigurationSettings.AppSettings("OPENae_InterfVB6VBNET.ServiceOPENae.ServiceOPENae")
			oTest.Timeout = 1800000
			If TxtTributo.Text = "8852" Then
				If oTest.EstraiTracciatoAE(TxtTributo.Text, TxtAnno.Text, TxtEnte.Text, TxtBelfiore.Text, TxtDescrEnte.Text, TxtCAPEnte.Text, TxtDataScadenza.Text, TxtProgInvio.Text, sFile) = "" Then
					MessageBox.Show("Errore", "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
				Else
					MessageBox.Show("Finito", "EVVIVA", MessageBoxButtons.OK, MessageBoxIcon.Information)
				End If
			Else
				If oTest.EstraiTracciatoAE(TxtTributo.Text, TxtAnno.Text, TxtEnte.Text, sFile) = "" Then
					MessageBox.Show("Errore", "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
				Else
					MessageBox.Show("Finito", "EVVIVA", MessageBoxButtons.OK, MessageBoxIcon.Information)
				End If
			End If
		Catch Err As Exception
			MessageBox.Show(Err.Message, "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
	End Sub

	Private Sub CmdInterfaccia_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdInterfaccia.Click
		Try
			Dim oTest As New OPENae_InterfVB6VBNET.ClsCall
			Dim sNameFile, sProvenienza As String
			Dim sFile As String = TxtFile.Text + TxtNomeFile.Text
			Dim sFileToUpload As String = TxtFileUpload.Text + TxtNomeFile.Text

			If OptRendicontazione.Checked = True Then
				sProvenienza = "O"
			Else
				sProvenienza = "R"
			End If
			If TxtTributo.Text = "8852" Then
				sNameFile = oTest.CreaFlussoMEF(TxtEnte.Text, TxtBelfiore.Text, TxtDescrEnte.Text, TxtCAPEnte.Text, TxtTributo.Text, TxtAnno.Text, TxtDataScadenza.Text, TxtProgInvio.Text, sFile, sfiletoupload, sFile.Replace(".txt", ".log"), sProvenienza, TxtFileDownload.Text)
			Else
				sNameFile = oTest.CallServiceAE(TxtConnessioneDB.Text, sFile, "LOG_CALLSERVICEAE.log", "e:\pal\")
			End If
			If sNameFile = "" Then
				MessageBox.Show("Errore", "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
			Else
				MessageBox.Show("Finito", "EVVIVA", MessageBoxButtons.OK, MessageBoxIcon.Information)
			End If
		Catch Err As Exception
			MessageBox.Show(Err.Message, "BACO", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End Try
	End Sub

	Private Sub OptRDB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptRDB.CheckedChanged
		If OptRDB.Checked = True Then
			OptRendicontazione.Checked = False
			TxtConnessioneDB.Enabled = True
			CmdCreaFile.Enabled = True : CmdPopolaDaFile.Enabled = True : Button1.Enabled = True
		Else
			OptRendicontazione.Checked = True
			TxtConnessioneDB.Enabled = False
			CmdCreaFile.Enabled = False : CmdPopolaDaFile.Enabled = False : Button1.Enabled = False
		End If
	End Sub

	Private Sub OptRendicontazione_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptRendicontazione.CheckedChanged
		If OptRendicontazione.Checked = True Then
			OptRDB.Checked = False
			TxtConnessioneDB.Enabled = False
			CmdCreaFile.Enabled = False : CmdPopolaDaFile.Enabled = False : Button1.Enabled = False
		Else
			OptRDB.Checked = True
			TxtConnessioneDB.Enabled = True
			CmdCreaFile.Enabled = True : CmdPopolaDaFile.Enabled = True : Button1.Enabled = True
		End If
	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		WriteLOG("c:\log_per_errore_OPENae.txt", "devo dichiarare")
		Dim FncAE As New ServiceOPENae.ServiceOPENae
		WriteLOG("c:\log_per_errore_OPENae.txt", "dichiarato")
		FncAE.Url = "http://localhost/WSOPENae/ServiceOPENae.asmx"
		WriteLOG("c:\log_per_errore_OPENae.txt", "riassegnato url")

		'Dim SQL As String
		'Dim rsRuolo As DAO.Recordset
		'Dim drResultRif As DAO.Recordset
		'Dim RsDatiFlusso As DAO.Recordset

		'Dim AnnoRuoloE As String

		'Dim DataDal, DataAl As Date
		'Dim CodImmobile As String

		'Dim ProgUnivoco As Integer

		'Dim ProgEstrazione As Integer
		'Dim MyArray() As String

		''^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^^.^
		'Dim i, y As Integer
		'Dim sSQL, sDataInizioOccup, FileSingle As String
		'Dim oDBManager As New GetConnectionDB
		'Dim myResult As Integer
		'Dim drResult As OleDb.OleDbDataReader
		'Dim drResultRif As OleDb.OleDbDataReader
		'Dim oMyTariffe() As Tariffe
		'Dim oMyDatiFlusso As ObjDatiFlusso

		''If UBound(tariffe, 1) = 0 Then
		''    MsgBox("Attenzione, non sono state codificate le categorie tariffarie attraverso l'apposito pulsante. Impossibile estrarre il flusso", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		''    Exit Sub
		''End If

		'If IsNumeric(TxtAnno.Text) = False Then
		'	MsgBox("Attenzione! E' necessario selezionare un anno a ruolo!", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		'	Exit Sub
		'End If

		''If TxtDal.Text = "" Then
		''    MsgBox("Attenzione! E' necessario inserire una data di partenza!", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		''    TxtDal.Focus()
		''    Exit Sub
		''Else
		''    If IsDate(TxtDal.Text) = False Or CDbl(Mid(TxtDal.Text, 4, 2)) > 12 Then
		''        MsgBox("Attenzione! E' necessario inserire una data di partenza valida!", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		''        TxtDal.Focus()
		''        Exit Sub
		''    End If
		''End If
		'Try
		'	ProgUnivoco = 0
		'	Try
		'		'pulisco la tabella dati flusso dai record non consolidati
		'		sSQL = "DELETE * FROM _AE_DATI_FLUSSO WHERE CONSOLIDATO =0"
		'		oDBManager.RunExecuteQueryACCESS(sSQL, TxtConnessioneDB.Text)
		'		sSQL = "DELETE [_AE_ESTRAZIONI].*, [_AE_ESTRAZIONI].PROG_ESTRAZIONE FROM _AE_ESTRAZIONI WHERE ((([_AE_ESTRAZIONI].PROG_ESTRAZIONE) Not In (SELECT PROG_ESTRAZIONE FROM _AE_DATI_FLUSSO)))"
		'		oDBManager.RunExecuteQueryACCESS(sSQL, TxtConnessioneDB.Text)
		'	Catch ex As Exception
		'		'creo tabella dati flusso
		'		sSQL = "CREATE TABLE _AE_DATI_FLUSSO ("
		'		sSQL += " PROG_ESTRAZIONE LONG, "
		'		sSQL += " PROGRESSIVO DOUBLE, "
		'		sSQL += " ABBINATO LONG, CONSOLIDATO LONG, "
		'		sSQL += " COD_CONTRIB LONG, NSCHEDA TEXT(10), "
		'		sSQL += " CFPIVA TEXT(16), COGNOME TEXT(26), NOME TEXT(25), "
		'		sSQL += " DENOMINAZIONE TEXT(50), COMUNE TEXT(40), PROVINCIA TEXT(2), "
		'		sSQL += " TITOLO_OCCUPAZIONE TEXT(1), OCCUPANTE TEXT(1), "
		'		sSQL += " DI_OCCUPAZIONE TEXT(10), DF_OCCUPAZIONE TEXT(10), "
		'		sSQL += " DI_RIFC TEXT(10), DF_RIFC TEXT(10), "
		'		sSQL += " DESTINAZIONE_USO TEXT(1), TIPO_UNITA TEXT(1), "
		'		sSQL += " SEZIONE TEXT(3), FOGLIO TEXT(5), PARTICELLA TEXT(5), EST_PARTICELLA TEXT(4),"
		'		sSQL += " TIPO_PARTICELLA TEXT(1) , SUBALTERNO TEXT(4),"
		'		sSQL += " UBICAZIONE TEXT(30), CIVICO TEXT(6), INTERNO TEXT(2), SCALA TEXT(1), CADC TEXT(1),"
		'		sSQL += " DENOMINAZIONECOM TEXT(60), CODFISCALE TEXT(16), COMUNELEGALE TEXT(40), PROVLEG TEXT(2), "
		'		sSQL += " COMUNEAMM TEXT(20), PROV TEXT(2), COMUNECAT TEXT(20), CODICE TEXT(5), CODISTAT TEXT(6) "
		'		sSQL += ")"
		'		oDBManager.RunExecuteQueryACCESS(sSQL, TxtConnessioneDB.Text)

		'		'creo tabella dati estrazione
		'		sSQL = "CREATE TABLE _AE_ESTRAZIONI ("
		'		sSQL += " PROG_ESTRAZIONE DOUBLE, "
		'		sSQL += " DATA_CREAZIONE DATETIME, DATA_ESTRAZIONE DATETIME, ANNO_RUOLO TEXT (4))"
		'		oDBManager.RunExecuteQueryACCESS(sSQL, TxtConnessioneDB.Text)
		'	End Try

		'	AnnoRuoloE = TxtAnno.Text

		'	'estrapolo il progressivo estrazione
		'	ProgEstrazione = ProgressivoEstrazione(AnnoRuoloE, oDBManager, TxtConnessioneDB.Text)
		'	oMyTariffe = PrelevaTariffe(TxtAnno.Text, oDBManager, TxtConnessioneDB.Text)
		'	'***  IL DAL LO DEVE PRENDERE DA OCCUPANTI QUELLO DI RUOLO E' UN DAL FITTIZIO ***
		'	sSQL = "SELECT DISTINCT RUOLO_TRRSU.COD_IMMOBILE, RUOLO_TRRSU.COD_OCCUPANTE, RUOLO_TRRSU.OCCUPATO_DAL, RUOLO_TRRSU.BIMESTRI, RUOLO_TRRSU.CATEGORIA, RUOLO_TRRSU.CIVICO, RIDUZIONI.COD_RIDUZIONE, ANAGRAFE.*, STRADARIO.*, IIf(Not [COD FISCALE] Is Null,[COD FISCALE],[PARTITA IVA]) AS CFPIVA, OCCUPANTI.[OCCUPATO DAL]"
		'	sSQL += " FROM (((RUOLO_TRRSU"
		'	sSQL += " INNER JOIN ANAGRAFE ON RUOLO_TRRSU.COD_OCCUPANTE = ANAGRAFE.[COD ANAGRAFICO])"
		'	sSQL += " LEFT JOIN STRADARIO ON RUOLO_TRRSU.COD_STRADA = STRADARIO.[COD STRADA])"
		'	sSQL += " LEFT JOIN RIDUZIONI ON (RUOLO_TRRSU.COD_OCCUPANTE = RIDUZIONI.[COD OCCUPANTE]) AND (RUOLO_TRRSU.ANNO = RIDUZIONI.ANNO) AND (RUOLO_TRRSU.COD_IMMOBILE = RIDUZIONI.[COD IMMOBILE]))"
		'	sSQL += " INNER JOIN OCCUPANTI ON (RUOLO_TRRSU.COD_IMMOBILE = OCCUPANTI.[COD IMMOBILE]) AND (RUOLO_TRRSU.COD_OCCUPANTE = OCCUPANTI.[COD OCCUPANTE])"
		'	sSQL += " WHERE(((RUOLO_TRRSU.ANNO) = " & AnnoRuoloE & ") AND ((OCCUPANTI.[OCCUPATO DAL]) >= #" & TxtDal.Text & "#))"
		'	sSQL += " AND  (([RUOLO_TRRSU].[COD_IMMOBILE] & CStr([RUOLO_TRRSU].[MQ])) IN ("
		'	sSQL += "    SELECT COD_IMMOBILE & CSTR(MAX(MQ))"
		'	sSQL += "    FROM RUOLO_TRRSU"
		'	sSQL += "    INNER JOIN OCCUPANTI ON (RUOLO_TRRSU.COD_OCCUPANTE = OCCUPANTI.[COD OCCUPANTE]) AND (RUOLO_TRRSU.COD_IMMOBILE = OCCUPANTI.[COD IMMOBILE])"
		'	sSQL += "    WHERE (ANNO=" & AnnoRuoloE & ") AND (OCCUPANTI.[OCCUPATO DAL] >= #" & TxtDal.Text & "#)"
		'	sSQL += "    GROUP BY COD_IMMOBILE"
		'	sSQL += " ))"
		'	sSQL += " ORDER BY RUOLO_TRRSU.COD_OCCUPANTE, RUOLO_TRRSU.COD_IMMOBILE, RUOLO_TRRSU.OCCUPATO_DAL"
		'	'********************************************************************************
		'	drResult = oDBManager.GetDataReaderACCESS(sSQL, TxtConnessioneDB.Text, "")
		'	Do While drResult.Read
		'		oMyDatiFlusso = New ObjDatiFlusso
		'		If ControlloCampo(drResult("cod_immobile")) = "" Then				'se non c'è codice immobile non posso valorizzare tutti i dati
		'			oMyDatiFlusso.PROGESTRAZIONE = ProgEstrazione
		'			ProgUnivoco += 1
		'			oMyDatiFlusso.PROGUNIVOCO = ProgUnivoco
		'			oMyDatiFlusso.CONSOLIDATO = 0
		'			oMyDatiFlusso.COD_OCCUPANTE = drResult("cod_occupante")
		'			oMyDatiFlusso.CFPIVA = drResult("cfpiva")
		'			If drResult("sesso") = "G" Then
		'				oMyDatiFlusso.DENOMINAZIONE = Mid(ControlloCampo(drResult("cognome")), 1, 50)
		'				oMyDatiFlusso.CITTA = Mid(ControlloCampo(drResult("CITTA RES")), 1, 40)
		'				oMyDatiFlusso.PROVINCIA = Mid(ControlloCampo(drResult("PROVINCIA RES")), 1, 2)
		'			Else
		'				oMyDatiFlusso.COGNOME = Mid(ControlloCampo(drResult("cognome")), 1, 26)
		'				oMyDatiFlusso.NOME = Mid(ControlloCampo(drResult("nome")), 1, 25)
		'			End If
		'			oMyDatiFlusso.TITOLOOCCUPAZIONE = 4
		'			'controllo se l'occupante è unico occupante (da elenco single o da riduzioni) e attribuisco in base a configurazione manuale
		'			oMyDatiFlusso.OCCUPANTE = ControlloOccupante(oMyTariffe, drResult("categoria"), ControlloCampo(drResult("cod_riduzione")), ControlloCampo(drResult("cfpiva")), FileSingle)

		'			If ControlloCampo(drResult("occupato_dal")) = "" Then
		'				oMyDatiFlusso.DATAINIZIO = +"31/12/" & CShort(AnnoRuoloE) - 1
		'			Else
		'				oMyDatiFlusso.DATAINIZIO = ControlloCampo(drResult("occupato_dal"))
		'			End If
		'			sDataInizioOccup = drResult("occupato_dal")
		'			Call CalcolaFineOccupazione(drResult("Bimestri"), sDataInizioOccup)
		'			If CDate(sDataInizioOccup) > CDate("31/12/" & AnnoRuoloE) Then
		'				sDataInizioOccup = "31/12/" & AnnoRuoloE
		'			End If
		'			oMyDatiFlusso.DATAFINE = sDataInizioOccup
		'			oMyDatiFlusso.DESTINAZIONEUSO = 1
		'			oMyDatiFlusso.ABBINATO = 0
		'			oMyDatiFlusso.TIPOUNITA = "F"
		'			oMyDatiFlusso.SEZIONE = ""
		'			oMyDatiFlusso.FOGLIO = ""
		'			oMyDatiFlusso.NUMERO = ""
		'			oMyDatiFlusso.ESTPARTICELLA = ""
		'			oMyDatiFlusso.TIPOPARTICELLA = ""
		'			oMyDatiFlusso.SUBALTERNO = ""
		'			oMyDatiFlusso.UBICAZIONE = Mid(ControlloCampo(drResult("TIP STRADA")) & " " & ControlloCampo(drResult("STRADA")), 1, 30)
		'			oMyDatiFlusso.CIVICO = Mid(ControlloCampo(drResult("CIVICO")), 1, 6)
		'			oMyDatiFlusso.INTERNO = ""
		'			oMyDatiFlusso.SCALA = ""
		'			oMyDatiFlusso.CADC = "3"
		'			If SetDatiFlusso(oMyDatiFlusso, oDBManager, TxtConnessioneDB.Text) = False Then
		'				Exit Sub
		'			End If
		'		Else

		'			'attraverso il codice scheda reperisco i riferimenti catastali
		'			DataDal = CDate("31/12/" & AnnoRuoloE)
		'			DataAl = CDate("01/01/" & CDbl(AnnoRuoloE) + 1)
		'			CodImmobile = drResult("cod_immobile")

		'			sSQL = "SELECT distinct Catasto_p.CodiceImmobile, Catasto_p.sezione, Catasto_p.Foglio, Catasto_p.Numero, Catasto_p.Sub"					  ', Min(Catasto_p.Dal) AS MinDiDal, Max(IIf(IsNull([CATASTO_P.AL),#" & DataDal & "#,[CATASTO_P.AL)) AS MaxDiAl"
		'			sSQL += " FROM Catasto_p INNER JOIN Catasto_s ON Catasto_p.CodiceCatasto = Catasto_s.CodiceCatasto"
		'			sSQL += " WHERE ((Catasto_p.Dal<=#" & DataDal & "# AND Catasto_p.Al Is Null) OR (Catasto_p.Dal<=#" & DataDal & "# AND Catasto_p.Al>=#" & DataAl & "#))"
		'			sSQL += " AND ((Catasto_s.Dal<=#" & DataDal & "# AND Catasto_s.Al Is Null) OR (Catasto_s.Dal<=#" & DataDal & "# AND Catasto_s.Al>=#" & DataAl & "#))"
		'			sSQL += " AND (Catasto_s.TipoRendita='RE')"
		'			sSQL += " AND (Catasto_p.CodiceImmobile = """ & CodImmobile & """)"
		'			sSQL += " GROUP BY Catasto_p.CodiceImmobile, Catasto_p.Sezione, Catasto_p.Foglio, "
		'			sSQL += " Catasto_p.Numero, Catasto_p.Sub"
		'			drResultRif = oDBManager.GetDataReaderACCESS(sSQL, TxtConnessioneDB.Text, "")
		'			If drResultRif.HasRows Then
		'				Do While drResultRif.Read
		'					oMyDatiFlusso = New ObjDatiFlusso
		'					'trovato riferimenti catastali attraverso codice immobile
		'					oMyDatiFlusso.PROGESTRAZIONE = ProgEstrazione
		'					ProgUnivoco += 1
		'					oMyDatiFlusso.PROGUNIVOCO = ProgUnivoco
		'					oMyDatiFlusso.CONSOLIDATO = "0"
		'					oMyDatiFlusso.COD_OCCUPANTE = drResult("cod_occupante")
		'					oMyDatiFlusso.CFPIVA = drResult("cfpiva")
		'					oMyDatiFlusso.COD_IMMOBILE = drResult("cod_immobile")
		'					If drResult("sesso") = "G" Then
		'						oMyDatiFlusso.DENOMINAZIONE = Mid(ControlloCampo(drResult("cognome")), 1, 50)
		'						oMyDatiFlusso.CITTA = Mid(ControlloCampo(drResult("CITTA RES")), 1, 40)
		'						oMyDatiFlusso.PROVINCIA = Mid(ControlloCampo(drResult("PROVINCIA RES")), 1, 2)
		'					Else
		'						oMyDatiFlusso.COGNOME = Mid(ControlloCampo(drResult("cognome")), 1, 26)
		'						oMyDatiFlusso.NOME = Mid(ControlloCampo(drResult("nome")), 1, 25)
		'					End If

		'					'controllo se l'occupante è tra i proprietari attivi
		'					'UPGRADE_WARNING: Couldn't resolve default property of object OccupanteAttivo(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'					oMyDatiFlusso.TITOLOOCCUPAZIONE = OccupanteAttivo(drResult("cod_occupante"), drResult("cod_immobile"), DataDal, DataAl, oDBManager, TxtConnessioneDB.Text)

		'					'controllo se l'occupante è unico occupante (da elenco single o da riduzioni) e attribuisco in base a configurazione manuale
		'					oMyDatiFlusso.OCCUPANTE = ControlloOccupante(oMyTariffe, drResult("categoria"), ControlloCampo(drResult("cod_riduzione")), ControlloCampo(drResult("cfpiva")), FileSingle)
		'					If ControlloCampo(drResult("occupato_dal")) = "" Then
		'						oMyDatiFlusso.DATAINIZIO = "31/12/" & CShort(AnnoRuoloE) - 1
		'					Else
		'						oMyDatiFlusso.DATAINIZIO = ControlloCampo(drResult("occupato_dal"))
		'					End If

		'					'recupero i dati di occupazione
		'					MyArray = Split(DatiOccupazione(oMyTariffe, ControlloCampo(drResult("categoria")), AnnoRuoloE, drResult("cod_immobile"), drResult("cod_occupante"), ControlloCampo(drResult("occupato_dal")), oDBManager, TxtConnessioneDB.Text), "-")
		'					sDataInizioOccup = drResult("occupato_dal")
		'					Call CalcolaFineOccupazione(drResult("Bimestri"), sDataInizioOccup)
		'					If CDate(sDataInizioOccup) > CDate("31/12/" & AnnoRuoloE) Then
		'						sDataInizioOccup = "31/12/" & AnnoRuoloE
		'					End If
		'					oMyDatiFlusso.DATAFINE = sDataInizioOccup

		'					oMyDatiFlusso.DESTINAZIONEUSO = MyArray(1)
		'					oMyDatiFlusso.ABBINATO = "1"
		'					'memorizzo i dati relativi a inizio e fine riferimento catastale principale
		'					oMyDatiFlusso.TIPOUNITA = "F"
		'					oMyDatiFlusso.SEZIONE = Mid(ControlloCampo(drResultRif("SEZIONE")), 1, 3)
		'					oMyDatiFlusso.FOGLIO = Mid(ControlloCampo(drResultRif("FOGLIO")), 1, 5)
		'					oMyDatiFlusso.NUMERO = Mid(ControlloCampo(drResultRif("numero")), 1, 5)
		'					oMyDatiFlusso.ESTPARTICELLA = ""
		'					oMyDatiFlusso.TIPOPARTICELLA = ""
		'					oMyDatiFlusso.SUBALTERNO = Mid(ControlloCampo(drResultRif("Sub")), 1, 4)
		'					oMyDatiFlusso.UBICAZIONE = Mid(ControlloCampo(drResult("TIP STRADA")) & " " & ControlloCampo(drResult("STRADA")), 1, 30)
		'					oMyDatiFlusso.CIVICO = Mid(ControlloCampo(drResult("CIVICO")), 1, 6)
		'					oMyDatiFlusso.INTERNO = ""
		'					oMyDatiFlusso.SCALA = ""
		'					If ControlloCampo(drResultRif("FOGLIO")) = "" Or ControlloCampo(drResultRif("numero")) = "" Then
		'						oMyDatiFlusso.CADC = "3"
		'					Else
		'						oMyDatiFlusso.CADC = "0"
		'					End If
		'					If SetDatiFlusso(oMyDatiFlusso, oDBManager, TxtConnessioneDB.Text) = False Then
		'						Exit Sub
		'					End If
		'				Loop
		'			Else
		'				oMyDatiFlusso = New ObjDatiFlusso
		'				oMyDatiFlusso.PROGESTRAZIONE = ProgEstrazione.ToString
		'				ProgUnivoco += 1
		'				oMyDatiFlusso.PROGUNIVOCO = ProgUnivoco.ToString
		'				oMyDatiFlusso.CONSOLIDATO = "0"
		'				oMyDatiFlusso.COD_OCCUPANTE = CStr(drResult("cod_occupante"))
		'				oMyDatiFlusso.CFPIVA = CStr(drResult("cfpiva"))
		'				oMyDatiFlusso.COD_IMMOBILE = CStr(drResult("cod_immobile"))
		'				If CStr(drResult("sesso")) = "G" Then
		'					oMyDatiFlusso.DENOMINAZIONE = Mid(ControlloCampo(drResult("cognome")), 1, 50)
		'					oMyDatiFlusso.CITTA = Mid(ControlloCampo(drResult("CITTA RES")), 1, 40)
		'					oMyDatiFlusso.PROVINCIA = Mid(ControlloCampo(drResult("PROVINCIA RES")), 1, 2)
		'				Else
		'					oMyDatiFlusso.COGNOME = Mid(ControlloCampo(drResult("cognome")), 1, 26)
		'					oMyDatiFlusso.NOME = Mid(ControlloCampo(drResult("nome")), 1, 25)
		'				End If
		'				oMyDatiFlusso.TITOLOOCCUPAZIONE = OccupanteAttivo(drResult("cod_occupante"), drResult("cod_immobile"), DataDal, DataAl, oDBManager, TxtConnessioneDB.Text)

		'				'controllo se l'occupante è unico occupante (da elenco single o da riduzioni) e attribuisco in base a configurazione manuale
		'				oMyDatiFlusso.OCCUPANTE = ControlloOccupante(oMyTariffe, drResult("categoria"), ControlloCampo(drResult("cod_riduzione")), ControlloCampo(drResult("cfpiva")), FileSingle)

		'				If ControlloCampo(drResult("occupato_dal")) = "" Then
		'					oMyDatiFlusso.DATAINIZIO = "31/12/" & CShort(AnnoRuoloE) - 1
		'				Else
		'					oMyDatiFlusso.DATAINIZIO = ControlloCampo(drResult("occupato_dal"))
		'				End If
		'				sDataInizioOccup = CStr(drResult("occupato_dal"))
		'				Call CalcolaFineOccupazione(drResult("Bimestri"), sDataInizioOccup)
		'				If CDate(sDataInizioOccup) > CDate("31/12/" & AnnoRuoloE) Then
		'					sDataInizioOccup = "31/12/" & AnnoRuoloE
		'				End If
		'				oMyDatiFlusso.DATAFINE = sDataInizioOccup
		'				MyArray = Split(DatiOccupazione(oMyTariffe, ControlloCampo(drResult("Categoria")), AnnoRuoloE, drResult("cod_immobile"), drResult("cod_occupante"), ControlloCampo(drResult("occupato_dal")), oDBManager, TxtConnessioneDB.Text), "-")
		'				oMyDatiFlusso.DESTINAZIONEUSO = MyArray(1)
		'				oMyDatiFlusso.ABBINATO = "0"
		'				oMyDatiFlusso.TIPOUNITA = "F"
		'				oMyDatiFlusso.SEZIONE = ""
		'				oMyDatiFlusso.FOGLIO = ""
		'				oMyDatiFlusso.NUMERO = ""
		'				oMyDatiFlusso.ESTPARTICELLA = ""
		'				oMyDatiFlusso.TIPOPARTICELLA = ""
		'				oMyDatiFlusso.SUBALTERNO = ""
		'				oMyDatiFlusso.UBICAZIONE = Mid(ControlloCampo(drResult("TIP STRADA")) & " " & ControlloCampo(drResult("STRADA")), 1, 30)
		'				oMyDatiFlusso.CIVICO = Mid(ControlloCampo(drResult("CIVICO")), 1, 6)
		'				oMyDatiFlusso.INTERNO = ""
		'				oMyDatiFlusso.SCALA = ""
		'				oMyDatiFlusso.CADC = "3"
		'				If SetDatiFlusso(oMyDatiFlusso, oDBManager, TxtConnessioneDB.Text) = False Then
		'					Exit Sub
		'				End If
		'			End If
		'			drResultRif.Close()
		'			'controllo su presenza dati catastali
		'		End If				'controllo su presenza codice immobile
		'	Loop
		'	MsgBox("Preparazione dati per estrazione flusso terminata correttamente!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		'Catch err As Exception
		'	Log.Debug("Button1_Click::si è verificato il seguente errore::" + err.Message)
		'	MsgBox(err.Message, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Agenzia Entrate")
		'Finally
		'	drResult.Close()
		'End Try
	End Sub

	Private Sub WriteLOG(ByVal glbFileLog As String, ByVal TextLOG As String)
		Dim DataOra As String

		FileOpen(1, glbFileLog, OpenMode.Append)
		PrintLine(1, "****************************************")
		DataOra = "Data e Ora: " & Now.Now
		PrintLine(1, DataOra & " Operazione: " & TextLOG)
		FileClose(1)

		'Dim StreamWriter As IO.StreamWriter = IO.File.AppendText(glbFileLog)

		'Try
		'    If TextLOG <> "" Then
		'        StreamWriter.WriteLine("****************************************")
		'        StreamWriter.WriteLine("Data e Ora: " & " Operazione: " & TextLOG)
		'    End If
		'    StreamWriter.Flush()
		'    StreamWriter.Close()
		'Catch err As Exception

		'End Try

	End Sub

End Class
