Namespace DLL
  Friend Class DataSetContatti
	Inherits System.ComponentModel.Component

#Region " Component Designer generated code "

	Public Sub New(ByVal Container As System.ComponentModel.IContainer)
	  MyClass.New()

	  'Required for Windows.Forms Class Composition Designer support
	  Container.Add(Me)
	End Sub

	Public Sub New()
	  MyBase.New()

	  'This call is required by the Component Designer.
	  InitializeComponent()

	  'Add any initialization after the InitializeComponent() call

	End Sub

	'Component overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
	  If disposing Then
		If Not (components Is Nothing) Then
		  components.Dispose()
		End If
	  End If
	  MyBase.Dispose(disposing)
	End Sub

	'Required by the Component Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Component Designer
	'It can be modified using the Component Designer.
	'Do not modify it using the code editor.
	Friend WithEvents daContatti As System.Data.SqlClient.SqlDataAdapter
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
	  Me.daContatti = New System.Data.SqlClient.SqlDataAdapter()
	  '
	  'daContatti
	  '
	  Me.daContatti.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "CONTATTI", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("IDRIFERIMENTO", "IDRIFERIMENTO"), New System.Data.Common.DataColumnMapping("TipoRiferimento", "TipoRiferimento"), New System.Data.Common.DataColumnMapping("DatiRiferimento", "DatiRiferimento"), New System.Data.Common.DataColumnMapping("COD_CONTRIBUENTE", "COD_CONTRIBUENTE"), New System.Data.Common.DataColumnMapping("IDDATAANAGRAFICA", "IDDATAANAGRAFICA")})})

	End Sub

#End Region

	Public Property DataSetCompleto() As dsContatti

	  Get
		'Restituisce un Data Set Completo dei Contatti
		Dim ds As dsContatti = New dsContatti()
		daContatti.Fill(ds.CONTATTI)

		Return ds

	  End Get

	  Set(ByVal Value As dsContatti)
		DataSetCompleto = Value
	  End Set

	End Property

	Public WriteOnly Property daContattiPersona() As SqlClient.SqlDataAdapter
	  Set(ByVal Value As SqlClient.SqlDataAdapter)
		daContatti = Value
	  End Set
	End Property


	Public Function ds(ByVal dsTemp As DataSet) As dsContatti

	  ds = dsTemp
	  Return ds

	End Function


  End Class
End Namespace