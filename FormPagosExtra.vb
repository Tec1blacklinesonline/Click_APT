' ============================================================================
' FORMPAGOSEXTRA.VB - NUEVO FORMULARIO PARA PAGOS EXTRA
' ✅ Maneja: Multas, Adiciones, Pagos Atrasados, Sanciones, etc.
' ✅ Integrado con el sistema existente de PDFs y correos
' ============================================================================

Imports System.Data.SQLite
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text
Imports System.IO
Imports System.Threading.Tasks
Imports System.Threading
Imports System.Linq
Imports System
Imports System.ComponentModel


Public Class FormPagosExtra
    Inherits Form
    Private numeroTorre As Integer
    Private apartamentos As List(Of Apartamento)
    Private WithEvents dgvPagosExtra As DataGridView
    Private btnRegistrarExtra As Button
    Private btnCancelar As Button
    Private btnEnvioMasivo As Button
    Private btnExportarPagos As Button
    Private cmbTipoPagoExtra As ComboBox
    Private txtValorExtra As TextBox
    Private txtConceptoExtra As TextBox
    Private panelBotones As Panel
    Private panelFiltros As Panel
    Private lblInfo As Label

    Public Sub New(numeroTorre As Integer)
        Me.numeroTorre = numeroTorre
        InitializeComponent()
        ConfigurarFormulario()
        CargarApartamentos()
    End Sub

    ' Clase para datos de descarga
    Public Class DatosDescargaPDF
        Public Property IdApartamento As Integer
        Public Property NumeroRecibo As String
        Public Property TipoPago As String
        Public Property Apartamento As String
        Public Property NombreArchivo As String
        Public Property RutaDestino As String
    End Class


    ' Clase para resultado de descarga masiva
    Public Class ResultadoDescargaMasiva
        Public Property Exitoso As Boolean
        Public Property Mensaje As String
        Public Property PDFsDescargados As Integer
        Public Property PDFsConError As Integer
        Public Property TotalPDFs As Integer
        Public Property RutaCarpetaDestino As String
        Public Property ErroresDetallados As New List(Of String)
        Public Property ArchivosGenerados As New List(Of String)
    End Class


    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.Text = "Pagos Extra - Torre " & numeroTorre.ToString()
        Me.Size = New Size(1600, 900)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(250, 250, 250)
        Me.WindowState = FormWindowState.Maximized
        Me.ResumeLayout(False)
    End Sub

    Private Sub ConfigurarFormulario()
        ' Header del formulario
        Dim headerPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(44, 62, 80) ' Color morado para diferenciarlo
        }

        Dim lblTitulo As New Label With {
            .Text = "💳 PAGOS EXTRA - TORRE " & numeroTorre.ToString(),
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .ForeColor = Color.White,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Fill
        }
        headerPanel.Controls.Add(lblTitulo)

        ' Panel de filtros y controles superiores
        ConfigurarPanelFiltros()

        ' DataGridView para pagos extra
        dgvPagosExtra = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(44, 62, 80), ' color azul de
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold),
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            },
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .Font = New Font("Segoe UI", 9),
                .SelectionBackColor = Color.FromArgb(44, 62, 80),
                .SelectionForeColor = Color.White
            },
            .RowHeadersVisible = False,
            .AllowUserToResizeRows = False,
            .ColumnHeadersHeight = 40,
            .RowTemplate = New DataGridViewRow With {.Height = 35}
        }

        ConfigurarColumnas()

        ' Eventos específicos para .NET 8
        AddHandler dgvPagosExtra.CellValueChanged, AddressOf dgvPagosExtra_CellValueChanged
        AddHandler dgvPagosExtra.CellClick, AddressOf dgvPagosExtra_CellClick
        AddHandler dgvPagosExtra.CellBeginEdit, AddressOf dgvPagosExtra_CellBeginEdit
        AddHandler dgvPagosExtra.CurrentCellDirtyStateChanged, AddressOf dgvPagosExtra_CurrentCellDirtyStateChanged
        AddHandler dgvPagosExtra.CellContentClick, AddressOf dgvPagosExtra_CellContentClick ' NUEVO para .NET 8
        AddHandler dgvPagosExtra.CellMouseClick, AddressOf dgvPagosExtra_CellMouseClick ' NUEVO para .NET 8

        ' Panel de botones
        ConfigurarPanelBotones()

        ' Agregar controles al formulario
        Me.Controls.Add(dgvPagosExtra)
        Me.Controls.Add(panelBotones)
        Me.Controls.Add(panelFiltros)
        Me.Controls.Add(headerPanel)
    End Sub

    Private Sub ConfigurarPanelFiltros()
        panelFiltros = New Panel With {
            .Dock = DockStyle.Top,
            .Height = 120,
            .BackColor = Color.FromArgb(236, 240, 241)
        }

        ' Título del panel
        Dim lblFiltroTitulo As New Label With {
            .Text = "🎯 CONFIGURAR PAGO EXTRA",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.FromArgb(44, 62, 80),
            .Location = New Point(20, 10),
            .AutoSize = True
        }

        ' Tipo de pago extra
        Dim lblTipoPago As New Label With {
            .Text = "Tipo de Pago:",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .Location = New Point(20, 45),
            .Size = New Size(100, 25)
        }

        cmbTipoPagoExtra = New ComboBox With {
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(130, 42),
            .Size = New Size(180, 30),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cmbTipoPagoExtra.Items.AddRange({"MULTA", "ADICION", "PAGO_ATRASADO", "SANCION", "REPARACION", "SERVICIO_EXTRA", "OTRO"})
        cmbTipoPagoExtra.SelectedIndex = 0

        ' Valor del pago extra
        Dim lblValor As New Label With {
            .Text = "Valor ($):",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .Location = New Point(330, 45),
            .Size = New Size(70, 25)
        }

        txtValorExtra = New TextBox With {
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(410, 42),
            .Size = New Size(120, 30),
            .Text = "0"
        }

        ' Concepto/Descripción
        Dim lblConcepto As New Label With {
            .Text = "Concepto:",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .Location = New Point(550, 45),
            .Size = New Size(80, 25)
        }

        txtConceptoExtra = New TextBox With {
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(640, 42),
            .Size = New Size(300, 30),
            .PlaceholderText = "Descripción del pago extra..."
        }

        ' Botón aplicar a seleccionados
        Dim btnAplicarSeleccionados As New Button With {
            .Text = "➕ APLICAR",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(39, 174, 96),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(220, 35),
            .Location = New Point(960, 40)
        }
        btnAplicarSeleccionados.FlatAppearance.BorderSize = 0
        AddHandler btnAplicarSeleccionados.Click, AddressOf btnAplicarSeleccionados_Click

        ' Instrucciones
        Dim lblInstrucciones As New Label With {
            .Text = "💡 Instrucciones: 1) Seleccione tipo y valor del pago extra, 2) Marque apartamentos en la tabla, 3) Clic 'Aplicar a Seleccionados', 4) Registrar pagos",
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(127, 140, 141),
            .Location = New Point(20, 85),
            .Size = New Size(1000, 25)
        }

        panelFiltros.Controls.AddRange({lblFiltroTitulo, lblTipoPago, cmbTipoPagoExtra, lblValor, txtValorExtra, lblConcepto, txtConceptoExtra, btnAplicarSeleccionados, lblInstrucciones})
    End Sub

    Private Sub ConfigurarColumnas()
        With dgvPagosExtra.Columns

            ' Columna de selección (checkbox) - OPTIMIZADA PARA .NET 8
            Dim chkColumn As New DataGridViewCheckBoxColumn With {
                .Name = "Seleccionar",
                .HeaderText = "☑",
                .Width = 70,
                .ReadOnly = False,
                .ThreeState = False,
                .TrueValue = True,
                .FalseValue = False,
                .IndeterminateValue = False
            }
            .Add(chkColumn)

            .Add("IdApartamento", "ID")
            .Add("Apartamento", "APART-")
            .Add("NombreResidente", "PROPIETARIO")
            .Add("TipoPagoExtra", "TIPO PAGO")
            .Add("ConceptoPago", "CONCEPTO")
            .Add("ValorPagoExtra", "VALOR EXTRA")
            .Add("FechaPago", "FECHA PAGO")
            .Add("Observaciones", "OBSERVACIONES")
            .Add("NumeroRecibo", "No. RECIBO")
            .Add("EstadoPago", "ESTADO")
        End With

        ' Configurar DataGridView para .NET 8
        dgvPagosExtra.EditMode = DataGridViewEditMode.EditOnEnter
        dgvPagosExtra.AllowUserToAddRows = False
        dgvPagosExtra.AllowUserToDeleteRows = False
        dgvPagosExtra.StandardTab = True

        ' Configurar anchos y propiedades
        dgvPagosExtra.Columns("IdApartamento").Width = 70
        dgvPagosExtra.Columns("Apartamento").Width = 120
        dgvPagosExtra.Columns("NombreResidente").Width = 250
        dgvPagosExtra.Columns("TipoPagoExtra").Width = 140
        dgvPagosExtra.Columns("ConceptoPago").Width = 220
        dgvPagosExtra.Columns("ValorPagoExtra").Width = 220
        dgvPagosExtra.Columns("FechaPago").Width = 150
        dgvPagosExtra.Columns("Observaciones").Width = 250
        dgvPagosExtra.Columns("NumeroRecibo").Width = 160
        dgvPagosExtra.Columns("EstadoPago").Width = 130

        ' Campos editables
        dgvPagosExtra.Columns("TipoPagoExtra").ReadOnly = False
        dgvPagosExtra.Columns("ConceptoPago").ReadOnly = False
        dgvPagosExtra.Columns("ValorPagoExtra").ReadOnly = False
        dgvPagosExtra.Columns("FechaPago").ReadOnly = False
        dgvPagosExtra.Columns("Observaciones").ReadOnly = False

        ' Campos readonly
        dgvPagosExtra.Columns("IdApartamento").ReadOnly = True
        dgvPagosExtra.Columns("Apartamento").ReadOnly = True
        dgvPagosExtra.Columns("NombreResidente").ReadOnly = True
        dgvPagosExtra.Columns("NumeroRecibo").ReadOnly = True
        dgvPagosExtra.Columns("EstadoPago").ReadOnly = True

        ' Formato de moneda
        dgvPagosExtra.Columns("ValorPagoExtra").DefaultCellStyle.Format = "C"
        dgvPagosExtra.Columns("ValorPagoExtra").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        ' Colores para campos editables
        dgvPagosExtra.Columns("TipoPagoExtra").DefaultCellStyle.BackColor = Color.LightYellow
        dgvPagosExtra.Columns("ConceptoPago").DefaultCellStyle.BackColor = Color.LightBlue
        dgvPagosExtra.Columns("ValorPagoExtra").DefaultCellStyle.BackColor = Color.LightYellow
        dgvPagosExtra.Columns("FechaPago").DefaultCellStyle.BackColor = Color.LightYellow
        dgvPagosExtra.Columns("Observaciones").DefaultCellStyle.BackColor = Color.LightBlue

        ' Botones de acción
        Dim btnPDFColumn As New DataGridViewButtonColumn With {
            .Name = "BtnPDF",
            .HeaderText = "PDF",
            .Text = "📄",
            .UseColumnTextForButtonValue = True,
            .Width = 60,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(231, 76, 60),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold)
            }
        }

        Dim btnCorreoColumn As New DataGridViewButtonColumn With {
            .Name = "BtnCorreo",
            .HeaderText = "✉",
            .Text = "📧",
            .UseColumnTextForButtonValue = True,
            .Width = 60,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(39, 174, 96),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold)
            }
        }

        dgvPagosExtra.Columns.Add(btnPDFColumn)
        dgvPagosExtra.Columns.Add(btnCorreoColumn)
    End Sub

    Private Sub ConfigurarPanelBotones()
        panelBotones = New Panel With {
        .Dock = DockStyle.Bottom,
        .Height = 90,
        .BackColor = Color.FromArgb(236, 240, 241)
    }

        btnRegistrarExtra = New Button With {
        .Text = "💳 REGISTRAR PAGOS",
        .Font = New Font("Segoe UI", 11, FontStyle.Bold),
        .ForeColor = Color.White,
        .BackColor = Color.FromArgb(231, 76, 60),
        .FlatStyle = FlatStyle.Flat,
        .Size = New Size(220, 40),
        .Location = New Point(20, 10)
    }
        btnRegistrarExtra.FlatAppearance.BorderSize = 0
        AddHandler btnRegistrarExtra.Click, AddressOf btnRegistrarExtra_Click

        btnCancelar = New Button With {
        .Text = "🔄 LIMPIAR",
        .Font = New Font("Segoe UI", 11, FontStyle.Bold),
        .ForeColor = Color.White,
       .BackColor = Color.FromArgb(44, 62, 80),
        .FlatStyle = FlatStyle.Flat,
        .Size = New Size(120, 40),
        .Location = New Point(260, 10)
    }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click

        ' ✅ NUEVO BOTÓN PARA DESCARGA MASIVA DE PDFs
        Dim btnDescargarPDFs As New Button With {
        .Text = "📥 DESCARGAR",
        .Font = New Font("Segoe UI", 11, FontStyle.Bold),
        .ForeColor = Color.White,
        .BackColor = Color.FromArgb(44, 62, 80), ' Color púrpura
        .FlatStyle = FlatStyle.Flat,
        .Size = New Size(180, 40),
        .Location = New Point(400, 10)
    }
        btnDescargarPDFs.FlatAppearance.BorderSize = 0
        AddHandler btnDescargarPDFs.Click, AddressOf btnDescargarPDFs_Click

        btnEnvioMasivo = New Button With {
        .Text = "📧 ENVÍO MASIVO",
        .Font = New Font("Segoe UI", 9, FontStyle.Bold),
        .ForeColor = Color.White,
        .BackColor = Color.FromArgb(44, 62, 80),
        .FlatStyle = FlatStyle.Flat,
        .Size = New Size(160, 40),
        .Location = New Point(600, 10) ' Ajustar posición
    }
        btnEnvioMasivo.FlatAppearance.BorderSize = 0
        AddHandler btnEnvioMasivo.Click, AddressOf btnEnvioMasivo_Click

        btnExportarPagos = New Button With {
        .Text = "📄 EXPORTAR",
        .Font = New Font("Segoe UI", 11, FontStyle.Bold),
        .ForeColor = Color.White,
        .BackColor = Color.FromArgb(44, 62, 80), '44, 62, 80 color azul ozuro
        .FlatStyle = FlatStyle.Flat,
        .Size = New Size(130, 40),
        .Location = New Point(780, 10) ' Ajustar posición
    }
        btnExportarPagos.FlatAppearance.BorderSize = 0
        AddHandler btnExportarPagos.Click, AddressOf btnExportarPagos_Click

        Dim btnVolver As New Button With {
        .Text = "← VOLVER",
        .Font = New Font("Segoe UI", 11, FontStyle.Bold),
        .ForeColor = Color.White,
        .BackColor = Color.FromArgb(44, 62, 80),
        .FlatStyle = FlatStyle.Flat,
        .Size = New Size(120, 40),
        .Location = New Point(930, 10) ' Ajustar posición
    }
        btnVolver.FlatAppearance.BorderSize = 0
        AddHandler btnVolver.Click, AddressOf btnVolver_Click

        lblInfo = New Label With {
        .Text = "💡 Campos AMARILLOS = Editables | AZULES = Descripciones | Seleccione apartamentos para operaciones masivas",
        .Font = New Font("Segoe UI", 10, FontStyle.Italic),
        .ForeColor = Color.FromArgb(127, 140, 141),
        .AutoSize = True,
        .Location = New Point(20, 60)
    }

        panelBotones.Controls.AddRange({btnRegistrarExtra, btnCancelar, btnDescargarPDFs, btnEnvioMasivo, btnExportarPagos, btnVolver, lblInfo})
    End Sub
    Private Sub CargarApartamentos()
        Try
            apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(numeroTorre)
            dgvPagosExtra.Rows.Clear()

            For Each apartamento In apartamentos
                Dim fila As Integer = dgvPagosExtra.Rows.Add()

                ' Llenar datos del apartamento - MEJORADO EL CHECKBOX
                dgvPagosExtra.Rows(fila).Cells("Seleccionar").Value = False ' Inicializar explícitamente
                dgvPagosExtra.Rows(fila).Cells("IdApartamento").Value = apartamento.IdApartamento
                dgvPagosExtra.Rows(fila).Cells("Apartamento").Value = "T" & numeroTorre.ToString() & "-" & apartamento.NumeroApartamento
                dgvPagosExtra.Rows(fila).Cells("NombreResidente").Value = If(String.IsNullOrEmpty(apartamento.NombreResidente), "No registrado", apartamento.NombreResidente)

                ' Campos para el pago extra (vacíos inicialmente)
                dgvPagosExtra.Rows(fila).Cells("TipoPagoExtra").Value = ""
                dgvPagosExtra.Rows(fila).Cells("ConceptoPago").Value = ""
                dgvPagosExtra.Rows(fila).Cells("ValorPagoExtra").Value = 0
                dgvPagosExtra.Rows(fila).Cells("FechaPago").Value = DateTime.Now.ToString("dd/MM/yyyy")
                dgvPagosExtra.Rows(fila).Cells("Observaciones").Value = ""
                dgvPagosExtra.Rows(fila).Cells("NumeroRecibo").Value = ""
                dgvPagosExtra.Rows(fila).Cells("EstadoPago").Value = "PENDIENTE"

                ' Estilo para pendientes
                dgvPagosExtra.Rows(fila).DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 220) ' Amarillo muy claro
            Next

            lblInfo.Text = $"📊 Torre {numeroTorre}: {apartamentos.Count} apartamentos cargados para pagos extra"

        Catch ex As Exception
            MessageBox.Show("Error al cargar apartamentos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ============================================================================
    ' EVENTOS DEL FORMULARIO
    ' ============================================================================

    Private Sub btnAplicarSeleccionados_Click(sender As Object, e As EventArgs)
        Try
            ' Validar campos
            Dim valor As Decimal
            If Not Decimal.TryParse(txtValorExtra.Text, valor) OrElse valor <= 0 Then
                MessageBox.Show("Ingrese un valor válido mayor a $0", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtValorExtra.Focus()
                Return
            End If

            If String.IsNullOrEmpty(txtConceptoExtra.Text.Trim()) Then
                MessageBox.Show("Ingrese una descripción del concepto", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtConceptoExtra.Focus()
                Return
            End If

            ' Aplicar a filas seleccionadas - MEJORADA LA VERIFICACIÓN
            Dim aplicados As Integer = 0
            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                ' Verificar múltiples formas de obtener el valor del checkbox
                Dim seleccionado As Boolean = False
                Try
                    Dim valorCelda = row.Cells("Seleccionar").Value
                    If valorCelda IsNot Nothing Then
                        If TypeOf valorCelda Is Boolean Then
                            seleccionado = CBool(valorCelda)
                        ElseIf valorCelda.ToString().ToLower() = "true" Then
                            seleccionado = True
                        End If
                    End If
                Catch ex As Exception
                    ' Si hay error, asumir no seleccionado
                    seleccionado = False
                End Try

                If seleccionado Then
                    row.Cells("TipoPagoExtra").Value = cmbTipoPagoExtra.SelectedItem.ToString()
                    row.Cells("ConceptoPago").Value = txtConceptoExtra.Text.Trim()
                    row.Cells("ValorPagoExtra").Value = valor
                    row.Cells("FechaPago").Value = DateTime.Now.ToString("dd/MM/yyyy")

                    ' Cambiar color de fila
                    row.DefaultCellStyle.BackColor = Color.FromArgb(220, 255, 220) ' Verde claro
                    aplicados += 1
                End If
            Next

            If aplicados > 0 Then
                MessageBox.Show($"✅ Pago extra aplicado a {aplicados} apartamento(s).{vbCrLf}Ahora puede registrar los pagos.",
                               "Aplicación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblInfo.Text = $"✅ Pago extra configurado para {aplicados} apartamentos. Clic en 'Registrar Pagos Extra'"
            Else
                MessageBox.Show("No hay apartamentos seleccionados. " & vbCrLf & vbCrLf &
                               "💡 Para seleccionar:" & vbCrLf &
                               "1. Haga clic en la casilla de verificación (☑) de cada apartamento" & vbCrLf &
                               "2. La casilla debe aparecer marcada" & vbCrLf &
                               "3. Luego haga clic en 'Aplicar a Seleccionados'",
                               "Aviso - Seleccionar Apartamentos", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error al aplicar pago extra: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRegistrarExtra_Click(sender As Object, e As EventArgs)
        Try
            ' Contar pagos para registrar
            Dim pagosParaRegistrar As Integer = 0
            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                Dim valor As Decimal = ConvertirADecimal(row.Cells("ValorPagoExtra").Value)
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                If valor > 0 AndAlso String.IsNullOrEmpty(numeroRecibo) Then
                    pagosParaRegistrar += 1
                End If
            Next

            If pagosParaRegistrar = 0 Then
                ' ✅ MENSAJE MEJORADO con información de debug
                Dim mensajeDebug As String = "No hay pagos extra para registrar." & vbCrLf & vbCrLf

                ' Verificar qué está pasando
                Dim filasConDatos As Integer = 0
                For Each row As DataGridViewRow In dgvPagosExtra.Rows
                    Dim valor As Decimal = ConvertirADecimal(row.Cells("ValorPagoExtra").Value)
                    Dim tipo As String = ObtenerValorCelda(row.Cells("TipoPagoExtra").Value)
                    If valor > 0 OrElse Not String.IsNullOrEmpty(tipo) Then
                        filasConDatos += 1
                    End If
                Next

                If filasConDatos > 0 Then
                    mensajeDebug += $"Se encontraron {filasConDatos} filas con datos, pero sin recibos." & vbCrLf &
                                   "Verifique que los valores sean mayores a $0 y que haya aplicado los pagos correctamente."
                Else
                    mensajeDebug += "Configure valores mayores a $0 usando el panel superior:" & vbCrLf &
                                   "1. Seleccione tipo de pago y valor" & vbCrLf &
                                   "2. Marque apartamentos (☑)" & vbCrLf &
                                   "3. Clic 'APLICAR A SELECCIONADOS'"
                End If

                MessageBox.Show(mensajeDebug, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim resultado As DialogResult = MessageBox.Show(
                $"¿Confirma el registro de {pagosParaRegistrar} pago(s) extra?",
                "Confirmar Registro",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                RegistrarPagosExtra()
            End If

        Catch ex As Exception
            MessageBox.Show("Error al registrar pagos extra: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RegistrarPagosExtra()
        Dim pagosExitosos As Integer = 0
        Dim pagosConError As New List(Of String)
        Dim recibosGenerados As New List(Of String)

        Try
            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                Try
                    ' Verificar que no esté ya registrado
                    Dim numeroReciboExistente As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                    If Not String.IsNullOrEmpty(numeroReciboExistente) Then
                        Continue For
                    End If

                    Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                    Dim apartamentoNombre As String = ObtenerValorCelda(row.Cells("Apartamento").Value)

                    ' Verificar que haya datos para registrar - MEJORADA LA VALIDACIÓN
                    Dim valorExtra As Decimal = ConvertirADecimal(row.Cells("ValorPagoExtra").Value)
                    Dim tipoPago As String = ObtenerValorCelda(row.Cells("TipoPagoExtra").Value)
                    Dim conceptoPago As String = ObtenerValorCelda(row.Cells("ConceptoPago").Value)

                    ' ✅ DEBUG: Mostrar valores para verificar
                    System.Diagnostics.Debug.WriteLine($"Fila {apartamentoNombre}: Valor={valorExtra}, Tipo={tipoPago}, Concepto={conceptoPago}")

                    If valorExtra <= 0 OrElse String.IsNullOrEmpty(tipoPago) Then
                        System.Diagnostics.Debug.WriteLine($"Saltando fila {apartamentoNombre} - Valor: {valorExtra}, Tipo: '{tipoPago}'")
                        Continue For
                    End If

                    ' ✅ CONTINUAR con el resto del código...

                    ' Crear objeto PagoModel para pago extra
                    Dim nuevoPagoExtra As New PagoModel()
                    nuevoPagoExtra.IdApartamento = idApartamento
                    nuevoPagoExtra.NumeroRecibo = GenerarNumeroRecibo(idApartamento)

                    Try
                        nuevoPagoExtra.FechaPago = DateTime.Parse(row.Cells("FechaPago").Value.ToString())
                    Catch
                        nuevoPagoExtra.FechaPago = DateTime.Now
                    End Try

                    ' Configurar como pago extra
                    nuevoPagoExtra.SaldoAnterior = 0 ' Los pagos extra no afectan saldo anterior
                    nuevoPagoExtra.PagoAdministracion = 0 ' No es pago de administración
                    nuevoPagoExtra.PagoIntereses = 0 ' No son intereses
                    nuevoPagoExtra.CuotaActual = valorExtra ' El valor extra va como cuota actual
                    nuevoPagoExtra.TotalPagado = valorExtra
                    nuevoPagoExtra.SaldoActual = 0 ' Los pagos extra no generan saldo pendiente
                    nuevoPagoExtra.Detalle = $"PAGO EXTRA - {tipoPago}: {conceptoPago}"
                    nuevoPagoExtra.Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value)
                    nuevoPagoExtra.EstadoPago = "REGISTRADO"
                    nuevoPagoExtra.UsuarioRegistro = "Sistema"
                    nuevoPagoExtra.FechaRegistro = DateTime.Now
                    nuevoPagoExtra.TipoPago = tipoPago ' MULTA, ADICION, etc.
                    nuevoPagoExtra.MetodoPago = "EFECTIVO"

                    ' Registrar usando PagosDAL modificado
                    If PagosExtraDAL.RegistrarPagoExtra(nuevoPagoExtra) Then
                        ' Éxito: Actualizar interfaz
                        row.Cells("NumeroRecibo").Value = nuevoPagoExtra.NumeroRecibo
                        row.Cells("EstadoPago").Value = "REGISTRADO"

                        ' Cambiar estilo visual
                        For Each cell As DataGridViewCell In row.Cells
                            If cell.ColumnIndex < dgvPagosExtra.Columns("BtnPDF").Index Then
                                cell.Style.BackColor = Color.FromArgb(200, 255, 200)
                                cell.Style.ForeColor = Color.FromArgb(0, 100, 0)
                                cell.ReadOnly = True
                            End If
                        Next

                        pagosExitosos += 1
                        recibosGenerados.Add(nuevoPagoExtra.NumeroRecibo)
                        lblInfo.Text = $"✅ Pago extra {nuevoPagoExtra.NumeroRecibo} registrado exitosamente"
                        Application.DoEvents()
                    Else
                        pagosConError.Add($"{apartamentoNombre}: Error al guardar en base de datos")
                    End If

                Catch ex As Exception
                    Dim apartamento As String = "N/A"
                    Try
                        apartamento = row.Cells("Apartamento").Value.ToString()
                    Catch
                    End Try
                    pagosConError.Add($"{apartamento}: {ex.Message}")
                    Continue For
                End Try
            Next

            ' Mostrar resumen
            Dim mensaje As String = $"Proceso completado:{vbCrLf}✅ Registrados: {pagosExitosos}"

            If recibosGenerados.Count > 0 Then
                mensaje += $"{vbCrLf}📄 Recibos: {String.Join(", ", recibosGenerados.Take(5))}"
                If recibosGenerados.Count > 5 Then
                    mensaje += $"... y {recibosGenerados.Count - 5} más"
                End If
            End If

            If pagosConError.Count > 0 Then
                mensaje += $"{vbCrLf}❌ Con errores: {pagosConError.Count}"
            End If

            MessageBox.Show(mensaje, "Resultado del Registro", MessageBoxButtons.OK,
                           If(pagosConError.Count = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))

            lblInfo.Text = $"🎉 Proceso completado: {pagosExitosos} pagos extra registrados"

        Catch ex As Exception
            MessageBox.Show($"Error general en registro de pagos extra: {ex.Message}",
                           "Error Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ============================================================================
    ' EVENTOS DE LA TABLA Y BOTONES
    ' ============================================================================

    ''' <summary>
    ''' Evento para manejar cambios inmediatos en checkboxes
    ''' </summary>
    Private Sub dgvPagosExtra_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If dgvPagosExtra.IsCurrentCellDirty Then
            dgvPagosExtra.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>
    ''' NUEVO: Manejo específico para clicks en contenido de celdas (.NET 8)
    ''' </summary>
    Private Sub dgvPagosExtra_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim columnName As String = dgvPagosExtra.Columns(e.ColumnIndex).Name

            If columnName = "Seleccionar" Then
                ' Forzar cambio del checkbox
                Dim currentValue As Boolean = False
                Try
                    If dgvPagosExtra.Rows(e.RowIndex).Cells("Seleccionar").Value IsNot Nothing Then
                        currentValue = CBool(dgvPagosExtra.Rows(e.RowIndex).Cells("Seleccionar").Value)
                    End If
                Catch
                    currentValue = False
                End Try

                ' Cambiar el valor
                dgvPagosExtra.Rows(e.RowIndex).Cells("Seleccionar").Value = Not currentValue
                dgvPagosExtra.InvalidateCell(e.ColumnIndex, e.RowIndex)

                ' Actualizar información
                ActualizarContadorSeleccionados()
            End If
        End If
    End Sub

    ''' <summary>
    ''' NUEVO: Manejo alternativo para clicks del mouse (.NET 8)
    ''' </summary>
    Private Sub dgvPagosExtra_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim columnName As String = dgvPagosExtra.Columns(e.ColumnIndex).Name

            If columnName = "Seleccionar" AndAlso e.Button = MouseButtons.Left Then
                ' Alternativa para manejar clicks en checkboxes
                Dim currentValue As Boolean = False
                Try
                    If dgvPagosExtra.Rows(e.RowIndex).Cells("Seleccionar").Value IsNot Nothing Then
                        currentValue = CBool(dgvPagosExtra.Rows(e.RowIndex).Cells("Seleccionar").Value)
                    End If
                Catch
                    currentValue = False
                End Try

                ' Cambiar el valor inmediatamente
                dgvPagosExtra.Rows(e.RowIndex).Cells("Seleccionar").Value = Not currentValue
                dgvPagosExtra.RefreshEdit()

                ' Actualizar información
                ActualizarContadorSeleccionados()
            End If
        End If
    End Sub

    ''' <summary>
    ''' NUEVO: Actualizar contador de apartamentos seleccionados
    ''' </summary>
    Private Sub ActualizarContadorSeleccionados()
        Try
            Dim seleccionados As Integer = 0
            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                Try
                    Dim valorCelda = row.Cells("Seleccionar").Value
                    If valorCelda IsNot Nothing AndAlso CBool(valorCelda) = True Then
                        seleccionados += 1
                    End If
                Catch
                    ' Ignorar errores de conversión
                End Try
            Next

            lblInfo.Text = $"📊 Torre {numeroTorre}: {seleccionados} apartamentos seleccionados para pago extra"
        Catch ex As Exception
            ' Error silencioso
        End Try
    End Sub

    Private Sub dgvPagosExtra_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        Dim columna As String = dgvPagosExtra.Columns(e.ColumnIndex).Name
        Dim row As DataGridViewRow = dgvPagosExtra.Rows(e.RowIndex)

        ' No permitir edición si ya está registrado
        Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
        If Not String.IsNullOrEmpty(numeroRecibo) Then
            e.Cancel = True
            Return
        End If

        ' Solo permitir edición de campos específicos
        Dim camposEditables As String() = {"TipoPagoExtra", "ConceptoPago", "ValorPagoExtra", "FechaPago", "Observaciones"}
        If Not camposEditables.Contains(columna) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub dgvPagosExtra_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim nombreColumna As String = dgvPagosExtra.Columns(e.ColumnIndex).Name

            If nombreColumna = "BtnPDF" OrElse nombreColumna = "BtnCorreo" Then
                Dim row As DataGridViewRow = dgvPagosExtra.Rows(e.RowIndex)
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

                If String.IsNullOrEmpty(numeroRecibo) Then
                    MessageBox.Show("Debe registrar el pago extra primero para generar recibo.",
                                  "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                If nombreColumna = "BtnPDF" Then
                    GenerarPDFPagoExtra(row)
                ElseIf nombreColumna = "BtnCorreo" Then
                    EnviarCorreoPagoExtra(row)
                End If
            End If
        End If
    End Sub

    Private Sub dgvPagosExtra_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Try
                Dim row As DataGridViewRow = dgvPagosExtra.Rows(e.RowIndex)
                Dim columnaEditada As String = dgvPagosExtra.Columns(e.ColumnIndex).Name

                ' Solo procesar si no está registrado
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                If Not String.IsNullOrEmpty(numeroRecibo) Then
                    Return
                End If

                ' Actualizar información cuando se cambie el valor
                If columnaEditada = "ValorPagoExtra" Then
                    Dim valor As Decimal = ConvertirADecimal(row.Cells("ValorPagoExtra").Value)
                    If valor > 0 Then
                        lblInfo.Text = $"💰 Pago extra de {valor:C} configurado para apartamento {row.Cells("Apartamento").Value}"
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show("Error en validación: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("¿Desea limpiar todos los campos de pagos extra?", "Confirmar",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CargarApartamentos()
            txtValorExtra.Text = "0"
            txtConceptoExtra.Text = ""
            cmbTipoPagoExtra.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnExportarPagos_Click(sender As Object, e As EventArgs)
        ' Implementar exportación similar al FormPagos original
        Try
            Dim pagosRegistrados As New List(Of DataGridViewRow)

            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                If Not String.IsNullOrEmpty(numeroRecibo) Then
                    pagosRegistrados.Add(row)
                End If
            Next

            If pagosRegistrados.Count = 0 Then
                MessageBox.Show("No hay pagos extra registrados para exportar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            MessageBox.Show($"Función de exportación lista para {pagosRegistrados.Count} pagos extra.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error al exportar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnVolver_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ============================================================================
    ' MÉTODOS PARA PDF Y CORREO DE PAGOS EXTRA
    ' ============================================================================
    Private Sub GenerarPDFPagoExtra(row As DataGridViewRow)
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

            ' Obtener pago extra desde la base de datos
            Dim pagoExtra As PagoModel = PagosExtraDAL.ObtenerPagoExtraPorNumeroRecibo(numeroRecibo)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If pagoExtra IsNot Nothing AndAlso apartamento IsNot Nothing Then
                ' Usar ReciboPDF modificado para pagos extra
                Dim rutaPdfGenerado As String = ReciboPDFExtra.GenerarReciboPagoExtra(pagoExtra, apartamento)

                If Not String.IsNullOrEmpty(rutaPdfGenerado) AndAlso File.Exists(rutaPdfGenerado) Then
                    Dim resultado As DialogResult = MessageBox.Show(
                        $"✅ Recibo de pago extra generado exitosamente.{vbCrLf}{vbCrLf}" &
                        $"📁 Ubicación: {rutaPdfGenerado}{vbCrLf}{vbCrLf}" &
                        $"¿Desea abrir el archivo ahora?",
                        "PDF Generado Correctamente",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information)

                    If resultado = DialogResult.Yes Then
                        Try
                            Process.Start(New ProcessStartInfo(rutaPdfGenerado) With {.UseShellExecute = True})
                        Catch ex As Exception
                            MessageBox.Show($"PDF generado correctamente pero no se pudo abrir: {ex.Message}",
                                          "PDF Generado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End Try
                    End If
                Else
                    MessageBox.Show("Error al generar el PDF del recibo de pago extra.", "Error PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al generar PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub EnviarCorreoPagoExtra(row As DataGridViewRow)
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

            Dim pagoExtra As PagoModel = PagosExtraDAL.ObtenerPagoExtraPorNumeroRecibo(numeroRecibo)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If pagoExtra Is Nothing OrElse apartamento Is Nothing Then
                MessageBox.Show("No se encontraron los datos del pago extra.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If String.IsNullOrWhiteSpace(apartamento.Correo) Then
                MessageBox.Show($"El apartamento {apartamento.ObtenerCodigoApartamento()} no tiene correo electrónico registrado.",
                              "Correo no registrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Confirmar envío
            Dim tipoPago As String = ObtenerValorCelda(row.Cells("TipoPagoExtra").Value)
            Dim resultado As DialogResult = MessageBox.Show(
                $"¿Confirma el envío del recibo de {tipoPago} No. {numeroRecibo}?" & vbCrLf & vbCrLf &
                $"📧 Destinatario: {apartamento.Correo}" & vbCrLf &
                $"👤 Propietario: {apartamento.NombreResidente}" & vbCrLf &
                $"🏠 Apartamento: {apartamento.ObtenerCodigoApartamento()}",
                "Confirmar Envío de Correo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                ' Generar PDF temporal
                Dim rutaPdfTemporal As String = ReciboPDFExtra.GenerarReciboPagoExtraTemporal(pagoExtra, apartamento)

                If Not String.IsNullOrEmpty(rutaPdfTemporal) AndAlso File.Exists(rutaPdfTemporal) Then
                    ' Enviar correo usando EmailService adaptado
                    Dim envioExitoso As Boolean = EmailServiceExtra.EnviarReciboPagoExtra(
                        apartamento.Correo,
                        apartamento.NombreResidente,
                        numeroRecibo,
                        tipoPago,
                        rutaPdfTemporal)

                    ' Limpiar archivo temporal
                    Try
                        If File.Exists(rutaPdfTemporal) Then
                            File.Delete(rutaPdfTemporal)
                        End If
                    Catch
                    End Try

                    If envioExitoso Then
                        MessageBox.Show(
                            $"✅ Recibo de {tipoPago} enviado exitosamente.{vbCrLf}{vbCrLf}" &
                            $"📧 Destinatario: {apartamento.Correo}{vbCrLf}" &
                            $"📄 Recibo: {numeroRecibo}",
                            "Correo Enviado",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Error al enviar el correo electrónico.", "Error de Envío", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBox.Show("Error al generar PDF temporal para el correo.", "Error PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al enviar correo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ' ============================================================================
    ' MÉTODOS AUXILIARES
    ' ============================================================================
    Private Function GenerarNumeroRecibo(idApartamento As Integer) As String
        Try
            Return GeneradorConsecutivosMejorado.GenerarNumeroReciboUnico()
        Catch ex As Exception
            MessageBox.Show($"Error al generar número de recibo: {ex.Message}",
                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return DateTime.Now.ToString("yyyyMMddHHmmssfff")
        End Try
    End Function

    Private Function ObtenerValorCelda(valor As Object) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return ""
        End If
        Return valor.ToString().Trim()
    End Function

    Private Function ConvertirADecimal(valor As Object) As Decimal
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return 0D
        End If

        Dim valorString As String = valor.ToString().Trim()
        If String.IsNullOrEmpty(valorString) Then
            Return 0D
        End If

        Dim resultado As Decimal
        If Decimal.TryParse(valorString, resultado) Then
            Return resultado
        End If

        Return 0D
    End Function

    ' ============================================================================
    ' MÉTODO MEJORADO PARA ENVÍO MASIVO EN TU FORMPAGOSEXTRA.VB
    ' Reemplaza tu método btnEnvioMasivo_Click existente
    ' ============================================================================
    Private Async Sub btnEnvioMasivo_Click(sender As Object, e As EventArgs)
        Dim formProgreso As FormProgresoEnvio = Nothing
        Dim recibosParaEnviar As New List(Of DatosEnvioRecibo)
        Dim cerrarFormDespues As Boolean = False

        Try
            If dgvPagosExtra.Rows.Count = 0 Then
                MessageBox.Show("No hay pagos extra registrados para enviar.", "Sin Datos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Recopilar recibos válidos
            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                Try
                    If row.IsNewRow OrElse row.Cells("NumeroRecibo").Value Is Nothing Then Continue For

                    Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                    If String.IsNullOrWhiteSpace(numeroRecibo) Then Continue For

                    Dim idApartamento As Integer
                    If Not Integer.TryParse(ObtenerValorCelda(row.Cells("IdApartamento").Value), idApartamento) Then Continue For

                    Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)
                    If apartamento Is Nothing OrElse String.IsNullOrWhiteSpace(apartamento.Correo) Then Continue For

                    Dim pagoExtra As PagoModel = PagosExtraDAL.ObtenerPagoExtraPorNumeroRecibo(numeroRecibo)
                    If pagoExtra Is Nothing Then Continue For

                    Dim rutaPdfTemporal As String = ""
                    Try
                        rutaPdfTemporal = ReciboPDFExtra.GenerarReciboPagoExtraTemporal(pagoExtra, apartamento)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine("Error generando PDF: " & ex.Message)
                        Continue For
                    End Try

                    If String.IsNullOrWhiteSpace(rutaPdfTemporal) OrElse Not File.Exists(rutaPdfTemporal) Then Continue For

                    recibosParaEnviar.Add(New DatosEnvioRecibo With {
                    .CorreoDestino = apartamento.Correo.Trim(),
                    .NombreDestino = If(String.IsNullOrWhiteSpace(apartamento.NombreResidente), "Propietario", apartamento.NombreResidente),
                    .NumeroRecibo = numeroRecibo,
                    .TipoPago = ObtenerValorCelda(row.Cells("TipoPagoExtra").Value),
                    .RutaPDF = rutaPdfTemporal,
                    .Apartamento = $"T{numeroTorre}-{apartamento.NumeroApartamento}",
                    .IdApartamento = idApartamento
                })

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("Error procesando fila: " & ex.Message)
                    Continue For
                End Try
            Next

            If recibosParaEnviar.Count = 0 Then
                MessageBox.Show("No hay recibos de pagos extra válidos para enviar por correo." & vbCrLf & vbCrLf &
                          "Verifique que:" & vbCrLf &
                          "• Los pagos estén registrados (tengan número de recibo)" & vbCrLf &
                          "• Los apartamentos tengan correo electrónico registrado" & vbCrLf &
                          "• Los recibos se puedan generar correctamente",
                          "Sin Recibos para Enviar", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim confirmacion As DialogResult = MessageBox.Show(
            $"¿Confirma el envío masivo de {recibosParaEnviar.Count} recibos de pagos extra?" & vbCrLf & vbCrLf &
            "📧 Se enviarán correos con los recibos adjuntos" & vbCrLf &
            "⏱️ Este proceso puede tomar varios minutos" & vbCrLf &
            "📦 Los correos se enviarán en lotes para cumplir límites de Gmail",
            "Confirmar Envío Masivo",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2)

            If confirmacion <> DialogResult.Yes Then
                LimpiarPDFsTemporalesSeguro(recibosParaEnviar)
                Return
            End If

            formProgreso = New FormProgresoEnvio()
            formProgreso.Text = $"Enviando {recibosParaEnviar.Count} Recibos - Torre {numeroTorre}"
            formProgreso.Show(Me)
            formProgreso.ActualizarProgreso("🚀 Iniciando envío masivo...", 0)

            btnEnvioMasivo.Enabled = False
            btnEnvioMasivo.Text = "📧 ENVIANDO..."
            btnEnvioMasivo.BackColor = Color.Gray
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()

            Dim progress As System.IProgress(Of ProgressInfo) = New System.Progress(Of ProgressInfo)(
            Sub(info)
                Try
                    If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                        formProgreso.ActualizarProgreso(info.Mensaje, info.Progreso)
                    End If
                    lblInfo.Text = info.Mensaje
                Catch
                    ' Silencioso
                End Try
            End Sub)

            Dim cancellationToken As CancellationToken = If(formProgreso?.CancellationToken, CancellationToken.None)
            Dim resultado As ResultadoEnvioMasivo = Nothing

            Try
                progress.Report(New ProgressInfo With {.Mensaje = "🔄 Preparando envío...", .Progreso = 5})
                Await Task.Delay(200)

                resultado = Await EmailServiceExtraAdaptado.EnviarRecibosMasivosMejorado(
                recibosParaEnviar,
                progress,
                cancellationToken)

            Catch envioEx As Exception
                resultado = New ResultadoEnvioMasivo With {
                .Exitoso = False,
                .Mensaje = "❌ Error crítico durante el envío: " & envioEx.Message,
                .EmailsExitosos = 0,
                .EmailsConError = recibosParaEnviar.Count,
                .TotalRecibos = recibosParaEnviar.Count
            }
                resultado.ErroresDetallados.Add("Error crítico: " & envioEx.Message)
            End Try

            If resultado IsNot Nothing Then
                MostrarResultadoEnvioMasivoDetallado(resultado, recibosParaEnviar.Count)
                If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                    formProgreso.MarcarCompletado(resultado.Exitoso, resultado.Mensaje)
                    cerrarFormDespues = True
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("❌ Error crítico: " & ex.Message & vbCrLf & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                formProgreso.MarcarCompletado(False, "Error crítico: " & ex.Message)
                cerrarFormDespues = True
            End If

        Finally
            btnEnvioMasivo.Enabled = True
            btnEnvioMasivo.Text = "📧 ENVÍO MASIVO"
            btnEnvioMasivo.BackColor = Color.FromArgb(52, 152, 219)
            Me.Cursor = Cursors.Default

            Try
                If recibosParaEnviar IsNot Nothing Then
                    LimpiarPDFsTemporalesSeguro(recibosParaEnviar)
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error limpiando PDFs: " & ex.Message)
            End Try
        End Try

        ' ⏳ Retardo final fuera del Finally (válido para Await)
        If cerrarFormDespues AndAlso formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
            Await Task.Delay(3000)
            Try
                formProgreso.Close()
                formProgreso.Dispose()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error cerrando formulario: " & ex.Message)
            End Try
        End If
    End Sub

    Private Async Sub btnDescargarPDFs_Click(sender As Object, e As EventArgs)
        Dim formProgreso As FormProgresoDescarga = Nothing
        Dim pdfsParaDescargar As New List(Of DatosDescargaPDF)
        Dim cerrarFormDespues As Boolean = False

        Try
            ' 1. Validar que hay datos para descargar
            If dgvPagosExtra.Rows.Count = 0 Then
                MessageBox.Show("No hay pagos extra registrados para descargar.", "Sin Datos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 2. Recopilar PDFs válidos para descarga
            For Each row As DataGridViewRow In dgvPagosExtra.Rows
                Try
                    If row.IsNewRow Then Continue For

                    ' Verificar si está seleccionado
                    Dim seleccionado As Boolean = False
                    Try
                        Dim valorCelda = row.Cells("Seleccionar").Value
                        If valorCelda IsNot Nothing Then
                            If TypeOf valorCelda Is Boolean Then
                                seleccionado = CBool(valorCelda)
                            ElseIf valorCelda.ToString().ToLower() = "true" Then
                                seleccionado = True
                            End If
                        End If
                    Catch
                        seleccionado = False
                    End Try

                    If Not seleccionado Then Continue For

                    ' Verificar que tiene número de recibo (está registrado)
                    Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                    If String.IsNullOrWhiteSpace(numeroRecibo) Then
                        Continue For
                    End If

                    Dim idApartamento As Integer
                    If Not Integer.TryParse(ObtenerValorCelda(row.Cells("IdApartamento").Value), idApartamento) Then
                        Continue For
                    End If

                    Dim tipoPago As String = ObtenerValorCelda(row.Cells("TipoPagoExtra").Value)
                    Dim apartamentoNombre As String = ObtenerValorCelda(row.Cells("Apartamento").Value)

                    ' Generar nombre de archivo usando número de recibo
                    Dim nombreArchivo As String = SanitizarNombreArchivo($"ReciboExtra_{numeroRecibo}.pdf")

                    pdfsParaDescargar.Add(New DatosDescargaPDF With {
                    .IdApartamento = idApartamento,
                    .NumeroRecibo = numeroRecibo,
                    .TipoPago = If(String.IsNullOrEmpty(tipoPago), "EXTRA", tipoPago),
                    .Apartamento = apartamentoNombre,
                    .NombreArchivo = nombreArchivo,
                    .RutaDestino = ""
                })

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"Error procesando fila para descarga: {ex.Message}")
                    Continue For
                End Try
            Next

            ' 3. Validar que hay PDFs para descargar
            If pdfsParaDescargar.Count = 0 Then
                MessageBox.Show("No hay recibos seleccionados para descargar." & vbCrLf & vbCrLf &
                      "Para descargar PDFs:" & vbCrLf &
                      "1. Marque los apartamentos deseados (☑)" & vbCrLf &
                      "2. Asegúrese de que los pagos estén registrados" & vbCrLf &
                      "3. Haga clic en 'DESCARGAR PDFs'",
                      "Sin PDFs para Descargar", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 4. Obtener carpeta de descargas del usuario
            Dim carpetaDescargas As String = ObtenerCarpetaDescargas()
            Dim carpetaDestino As String = Path.Combine(carpetaDescargas, "COOPDIASAM_PagosExtra", $"Torre_{numeroTorre}_{DateTime.Now:yyyyMMdd}")

            ' 5. Confirmar descarga
            Dim confirmacion As DialogResult = MessageBox.Show(
        $"¿Confirma la descarga de {pdfsParaDescargar.Count} recibos de pagos extra?" & vbCrLf & vbCrLf &
        $"📁 Destino: {carpetaDestino}" & vbCrLf &
        $"📄 Los archivos se nombrarán con el número de recibo" & vbCrLf &
        $"⏱️ Este proceso puede tomar varios minutos",
        "Confirmar Descarga Masiva",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question,
        MessageBoxDefaultButton.Button1)

            If confirmacion <> DialogResult.Yes Then
                Return
            End If

            ' 6. Crear directorio de destino
            Try
                If Not Directory.Exists(carpetaDestino) Then
                    Directory.CreateDirectory(carpetaDestino)
                End If

                ' Actualizar rutas de destino
                For Each pdf In pdfsParaDescargar
                    pdf.RutaDestino = Path.Combine(carpetaDestino, pdf.NombreArchivo)
                Next

            Catch ex As Exception
                MessageBox.Show($"Error creando carpeta de destino: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' 7. Mostrar formulario de progreso y procesar
            formProgreso = New FormProgresoDescarga()
            formProgreso.Text = $"Descargando {pdfsParaDescargar.Count} PDFs - Torre {numeroTorre}"
            formProgreso.Show(Me)

            ' Deshabilitar controles durante el proceso
            DeshabilitarControlesDuranteDescarga(True)

            ' Crear progress reporter
            Dim progress As System.IProgress(Of ProgressInfo) = New System.Progress(Of ProgressInfo)(
        Sub(info)
            Try
                If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                    formProgreso.ActualizarProgreso(info.Mensaje, info.Progreso)
                End If
                lblInfo.Text = info.Mensaje
                Application.DoEvents()
            Catch
                ' Error silencioso
            End Try
        End Sub)

            ' 8. Ejecutar descarga masiva
            Dim resultado As ResultadoDescargaMasiva = Await ProcesarDescargaMasivaPDFs(pdfsParaDescargar, progress, formProgreso?.CancellationToken)

            ' 9. Mostrar resultado
            If resultado IsNot Nothing Then
                MostrarResultadoDescargaMasiva(resultado)
                If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                    formProgreso.MarcarCompletado(resultado.Exitoso, resultado.Mensaje)
                End If
            End If

            cerrarFormDespues = True

        Catch ex As Exception
            MessageBox.Show($"Error crítico en descarga masiva: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                formProgreso.MarcarCompletado(False, "Error crítico: " & ex.Message)
            End If
            cerrarFormDespues = True

        Finally
            ' Rehabilitar controles
            DeshabilitarControlesDuranteDescarga(False)

            ' Cerrar formulario de progreso (sin delay aquí)
            If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                Try
                    If Not cerrarFormDespues Then
                        formProgreso.Close()
                        formProgreso.Dispose()
                    End If
                Catch
                    ' Error silencioso
                End Try
            End If
        End Try

        ' Aplicar delay después del bloque Try-Catch-Finally si es necesario
        If cerrarFormDespues AndAlso formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
            Await Task.Delay(3000)
            Try
                formProgreso.Close()
                formProgreso.Dispose()
            Catch
                ' Error silencioso
            End Try
        End If
    End Sub
    Private Async Function ProcesarDescargaMasivaPDFs(pdfs As List(Of DatosDescargaPDF), progress As IProgress(Of ProgressInfo), cancellationToken As CancellationToken) As Task(Of ResultadoDescargaMasiva)
        Dim resultado As New ResultadoDescargaMasiva With {
        .TotalPDFs = pdfs.Count,
        .RutaCarpetaDestino = Path.GetDirectoryName(pdfs.FirstOrDefault()?.RutaDestino)
    }

        Try
            progress?.Report(New ProgressInfo With {.Mensaje = "🚀 Iniciando descarga masiva de PDFs...", .Progreso = 0})
            Await Task.Delay(500, cancellationToken)

            For i As Integer = 0 To pdfs.Count - 1
                If cancellationToken.IsCancellationRequested Then
                    resultado.Mensaje = "❌ Descarga cancelada por el usuario"
                    resultado.Exitoso = False
                    Return resultado
                End If

                Dim pdf As DatosDescargaPDF = pdfs(i)
                Dim progreso As Integer = CInt((i / pdfs.Count) * 100)

                progress?.Report(New ProgressInfo With {
                .Mensaje = $"📄 Generando PDF {i + 1}/{pdfs.Count}: {pdf.NumeroRecibo}...",
                .Progreso = progreso
            })

                Try
                    ' Obtener datos del pago extra
                    Dim pagoExtra As PagoModel = PagosExtraDAL.ObtenerPagoExtraPorNumeroRecibo(pdf.NumeroRecibo)
                    If pagoExtra Is Nothing Then
                        resultado.ErroresDetallados.Add($"No se encontró el pago {pdf.NumeroRecibo}")
                        resultado.PDFsConError += 1
                        Continue For
                    End If

                    Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(pdf.IdApartamento)
                    If apartamento Is Nothing Then
                        resultado.ErroresDetallados.Add($"No se encontró el apartamento {pdf.IdApartamento}")
                        resultado.PDFsConError += 1
                        Continue For
                    End If

                    ' Generar PDF en la ruta específica
                    Dim rutaPDFGenerado As String = ReciboPDFExtra.GenerarReciboPagoExtraEspecifico(pagoExtra, apartamento, pdf.RutaDestino)

                    If Not String.IsNullOrEmpty(rutaPDFGenerado) AndAlso File.Exists(rutaPDFGenerado) Then
                        resultado.PDFsDescargados += 1
                        resultado.ArchivosGenerados.Add(rutaPDFGenerado)

                        progress?.Report(New ProgressInfo With {
                        .Mensaje = $"✅ PDF generado: {Path.GetFileName(rutaPDFGenerado)}",
                        .Progreso = progreso
                    })
                    Else
                        resultado.ErroresDetallados.Add($"Error generando PDF para recibo {pdf.NumeroRecibo}")
                        resultado.PDFsConError += 1
                    End If

                    ' Pequeña pausa para no sobrecargar el sistema
                    Await Task.Delay(100, cancellationToken)

                Catch ex As Exception
                    resultado.ErroresDetallados.Add($"Error en {pdf.NumeroRecibo}: {ex.Message}")
                    resultado.PDFsConError += 1
                    System.Diagnostics.Debug.WriteLine($"Error generando PDF {pdf.NumeroRecibo}: {ex.Message}")
                End Try
            Next

            ' Resultado final
            progress?.Report(New ProgressInfo With {.Mensaje = "🎯 Finalizando descarga masiva...", .Progreso = 95})
            Await Task.Delay(500, cancellationToken)

            resultado.Exitoso = (resultado.PDFsDescargados > 0)
            If resultado.Exitoso Then
                resultado.Mensaje = $"✅ Descarga completada: {resultado.PDFsDescargados}/{resultado.TotalPDFs} PDFs generados correctamente"
            Else
                resultado.Mensaje = $"❌ Error en descarga: No se pudo generar ningún PDF"
            End If

            progress?.Report(New ProgressInfo With {.Mensaje = resultado.Mensaje, .Progreso = 100})

        Catch ex As Exception
            resultado.Exitoso = False
            resultado.Mensaje = $"❌ Error crítico: {ex.Message}"
            resultado.ErroresDetallados.Add($"Error crítico: {ex.Message}")
        End Try

        Return resultado
    End Function


#Region "Métodos de Soporte"

    ''' <summary>
    ''' Obtiene la carpeta de descargas del usuario
    ''' </summary>
    Private Function ObtenerCarpetaDescargas() As String
        Try
            ' Intentar obtener carpeta de descargas estándar
            Dim carpetaDescargas As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")

            If Not Directory.Exists(carpetaDescargas) Then
                ' Fallback a documentos si no existe Downloads
                carpetaDescargas = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            Return carpetaDescargas

        Catch ex As Exception
            ' Fallback final al escritorio
            Return Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End Try
    End Function

    ''' <summary>
    ''' Sanitiza nombres de archivo para evitar caracteres inválidos
    ''' </summary>
    Private Function SanitizarNombreArchivo(nombreArchivo As String) As String
        Try
            Dim caracteresInvalidos As Char() = Path.GetInvalidFileNameChars()
            For Each caracter In caracteresInvalidos
                nombreArchivo = nombreArchivo.Replace(caracter, "_"c)
            Next

            ' Limitar longitud del nombre
            If nombreArchivo.Length > 100 Then
                Dim extension As String = Path.GetExtension(nombreArchivo)
                Dim nombreSinExtension As String = Path.GetFileNameWithoutExtension(nombreArchivo)
                nombreArchivo = nombreSinExtension.Substring(0, 90) & extension
            End If

            Return nombreArchivo

        Catch
            Return "ReciboExtra_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"
        End Try
    End Function

    ''' <summary>
    ''' Habilita/deshabilita controles durante la descarga
    ''' </summary>
    Private Sub DeshabilitarControlesDuranteDescarga(deshabilitar As Boolean)
        Try
            ' Buscar el botón de descarga en el panel de botones
            For Each control As Control In panelBotones.Controls
                If TypeOf control Is Button Then
                    Dim btn As Button = CType(control, Button)
                    If btn.Text.Contains("DESCARGAR") Then
                        btn.Enabled = Not deshabilitar
                        If deshabilitar Then
                            btn.Text = "📥 DESCARGANDO..."
                            btn.BackColor = Color.Gray
                        Else
                            btn.Text = "📥 DESCARGAR PDFs"
                            btn.BackColor = Color.FromArgb(155, 89, 182)
                        End If
                        Exit For
                    End If
                End If
            Next

            ' Otros controles importantes
            btnRegistrarExtra.Enabled = Not deshabilitar
            btnEnvioMasivo.Enabled = Not deshabilitar
            panelFiltros.Enabled = Not deshabilitar

            ' Cursor
            Me.Cursor = If(deshabilitar, Cursors.WaitCursor, Cursors.Default)

            Application.DoEvents()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en DeshabilitarControlesDuranteDescarga: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Muestra el resultado detallado de la descarga masiva
    ''' </summary>
    Private Sub MostrarResultadoDescargaMasiva(resultado As ResultadoDescargaMasiva)
        Try
            Dim icono As MessageBoxIcon = If(resultado.PDFsConError = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning)
            Dim titulo As String = "Resultado de Descarga Masiva - Torre " & numeroTorre.ToString()

            Dim mensaje As String = ""
            mensaje = mensaje & "📁 DESCARGA MASIVA DE PDFs COMPLETADA" & vbCrLf & vbCrLf
            mensaje = mensaje & resultado.Mensaje & vbCrLf & vbCrLf
            mensaje = mensaje & "📊 ESTADÍSTICAS:" & vbCrLf
            mensaje = mensaje & "✅ PDFs generados: " & resultado.PDFsDescargados.ToString() & vbCrLf
            mensaje = mensaje & "❌ Con errores: " & resultado.PDFsConError.ToString() & vbCrLf
            mensaje = mensaje & "📄 Total procesados: " & resultado.TotalPDFs.ToString() & vbCrLf

            If resultado.PDFsDescargados > 0 Then
                mensaje = mensaje & vbCrLf & "📁 UBICACIÓN:" & vbCrLf
                mensaje = mensaje & resultado.RutaCarpetaDestino & vbCrLf
            End If

            ' Mostrar errores si los hay (máximo 5)
            If resultado.ErroresDetallados.Count > 0 Then
                mensaje = mensaje & vbCrLf & "🔍 ERRORES ENCONTRADOS:" & vbCrLf
                For i = 0 To Math.Min(4, resultado.ErroresDetallados.Count - 1)
                    mensaje = mensaje & "• " & resultado.ErroresDetallados(i) & vbCrLf
                Next

                If resultado.ErroresDetallados.Count > 5 Then
                    mensaje = mensaje & "... y " & (resultado.ErroresDetallados.Count - 5).ToString() & " errores más" & vbCrLf
                End If
            End If

            ' Agregar acción recomendada
            If resultado.PDFsDescargados > 0 Then
                mensaje = mensaje & vbCrLf & "🎉 ¡PDFs descargados exitosamente!" & vbCrLf
                mensaje = mensaje & "💡 ¿Desea abrir la carpeta de destino?" & vbCrLf

                Dim resultadoMsg As DialogResult = MessageBox.Show(mensaje, titulo, MessageBoxButtons.YesNo, icono)

                If resultadoMsg = DialogResult.Yes AndAlso Not String.IsNullOrEmpty(resultado.RutaCarpetaDestino) Then
                    Try
                        Process.Start("explorer.exe", resultado.RutaCarpetaDestino)
                    Catch ex As Exception
                        MessageBox.Show($"No se pudo abrir la carpeta: {ex.Message}", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End Try
                End If
            Else
                MessageBox.Show(mensaje, titulo, MessageBoxButtons.OK, icono)
            End If

            ' Actualizar información en el formulario
            If resultado.PDFsDescargados > 0 Then
                lblInfo.Text = $"📁 Descarga completada: {resultado.PDFsDescargados} PDFs guardados en Descargas/COOPDIASAM_PagosExtra"
            Else
                lblInfo.Text = "❌ Error en descarga masiva: Verifique los datos y reintente"
            End If

        Catch ex As Exception
            MessageBox.Show("Error mostrando resultado: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

    ' ============================================================================
    ' MÉTODOS DE SOPORTE PARA EL ENVÍO MASIVO
    ' ============================================================================

    ''' <summary>
    ''' Deshabilita o habilita controles durante el envío masivo
    ''' </summary>
    Private Sub DeshabilitarControlesDuranteEnvio(deshabilitar As Boolean)
        Try
            ' Deshabilitar/habilitar botones principales
            btnEnvioMasivo.Enabled = Not deshabilitar
            btnRegistrarExtra.Enabled = Not deshabilitar
            btnExportarPagos.Enabled = Not deshabilitar
            btnCancelar.Enabled = Not deshabilitar

            ' La tabla se mantiene habilitada para consulta pero no para edición
            If deshabilitar Then
                dgvPagosExtra.ReadOnly = True
            Else
                ' Restaurar configuración original de campos editables
                dgvPagosExtra.ReadOnly = False
                dgvPagosExtra.Columns("TipoPagoExtra").ReadOnly = False
                dgvPagosExtra.Columns("ConceptoPago").ReadOnly = False
                dgvPagosExtra.Columns("ValorPagoExtra").ReadOnly = False
                dgvPagosExtra.Columns("FechaPago").ReadOnly = False
                dgvPagosExtra.Columns("Observaciones").ReadOnly = False
            End If

            ' Panel de filtros
            panelFiltros.Enabled = Not deshabilitar

            ' Cambiar apariencia del botón y cursor
            If deshabilitar Then
                Me.Cursor = Cursors.WaitCursor
                btnEnvioMasivo.Text = "📧 ENVIANDO..."
                btnEnvioMasivo.BackColor = Color.FromArgb(127, 140, 141)
            Else
                Me.Cursor = Cursors.Default
                btnEnvioMasivo.Text = "📧 ENVÍO MASIVO"
                btnEnvioMasivo.BackColor = Color.FromArgb(52, 152, 219)
            End If

            Application.DoEvents()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en DeshabilitarControlesDuranteEnvio: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Muestra el resultado detallado del envío masivo
    ''' </summary>
    Private Sub MostrarResultadoEnvioMasivoDetallado(resultado As ResultadoEnvioMasivo, totalRecibos As Integer)
        Try
            Dim icono As MessageBoxIcon = If(resultado.EmailsConError = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning)
            Dim titulo As String = "Resultado del Envío Masivo - Torre " & numeroTorre.ToString()

            ' Crear mensaje detallado usando concatenación simple
            Dim mensaje As String = ""
            mensaje = mensaje & "🎯 ENVÍO MASIVO DE PAGOS EXTRA COMPLETADO" & vbCrLf & vbCrLf
            mensaje = mensaje & resultado.Mensaje & vbCrLf & vbCrLf
            mensaje = mensaje & "📊 ESTADÍSTICAS DETALLADAS:" & vbCrLf
            mensaje = mensaje & "✅ Exitosos: " & resultado.EmailsExitosos.ToString() & " correos" & vbCrLf
            mensaje = mensaje & "❌ Con errores: " & resultado.EmailsConError.ToString() & " correos" & vbCrLf
            mensaje = mensaje & "📄 Total procesados: " & totalRecibos.ToString() & " recibos" & vbCrLf

            ' Calcular porcentaje de éxito
            Dim porcentajeExito As Double = If(totalRecibos > 0, (resultado.EmailsExitosos / totalRecibos) * 100, 0)
            mensaje = mensaje & "📈 Tasa de éxito: " & porcentajeExito.ToString("F1") & "%" & vbCrLf

            ' Agregar detalles de errores si los hay (máximo 5)
            If resultado.ErroresDetallados.Count > 0 Then
                mensaje = mensaje & vbCrLf & "🔍 ERRORES ENCONTRADOS:" & vbCrLf
                For i = 0 To Math.Min(4, resultado.ErroresDetallados.Count - 1)
                    mensaje = mensaje & "• " & resultado.ErroresDetallados(i) & vbCrLf
                Next

                If resultado.ErroresDetallados.Count > 5 Then
                    mensaje = mensaje & "... y " & (resultado.ErroresDetallados.Count - 5).ToString() & " errores más" & vbCrLf
                End If
            End If

            ' Agregar consejos según el resultado
            mensaje = mensaje & vbCrLf
            If resultado.EmailsConError > 0 Then
                mensaje = mensaje & "💡 RECOMENDACIONES:" & vbCrLf
                mensaje = mensaje & "• Verifique las direcciones de correo electrónico" & vbCrLf
                mensaje = mensaje & "• Los correos pueden llegar a spam/promociones" & vbCrLf
                mensaje = mensaje & "• Puede reenviar individualmente desde cada fila" & vbCrLf
                mensaje = mensaje & "• Contacte a los propietarios para confirmar recepción" & vbCrLf
            Else
                mensaje = mensaje & "🎉 ¡EXCELENTE! Todos los correos se enviaron correctamente." & vbCrLf
                mensaje = mensaje & "💡 Recuerde informar a los propietarios que revisen" & vbCrLf
                mensaje = mensaje & "   su bandeja de entrada y carpeta de spam/promociones." & vbCrLf
            End If

            MessageBox.Show(mensaje, titulo, MessageBoxButtons.OK, icono)

            ' Actualizar información en el formulario
            If resultado.EmailsExitosos > 0 Then
                lblInfo.Text = "🎉 Envío masivo completado: " & resultado.EmailsExitosos.ToString() & "/" & totalRecibos.ToString() & " recibos de pagos extra enviados exitosamente"
            Else
                lblInfo.Text = "❌ Envío masivo falló: Verifique los datos y reintente"
            End If

        Catch ex As Exception
            MessageBox.Show("Error mostrando resultado: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblInfo.Text = "❌ Error mostrando resultado del envío masivo"
        End Try
    End Sub
    ''' <summary>
    ''' Limpia archivos PDF temporales generados para el envío de forma segura
    ''' </summary>
    Private Sub LimpiarPDFsTemporalesSeguro(recibos As List(Of DatosEnvioRecibo))
        Try
            Dim archivosEliminados As Integer = 0
            Dim erroresEliminacion As Integer = 0

            For Each recibo In recibos
                Try
                    If Not String.IsNullOrEmpty(recibo.RutaPDF) AndAlso
                       File.Exists(recibo.RutaPDF) AndAlso
                       recibo.RutaPDF.Contains("COOPDIASAM_Temp_PDFs_Extra") Then

                        ' Intentar eliminar el archivo con reintentos
                        For intento = 1 To 3
                            Try
                                File.Delete(recibo.RutaPDF)
                                archivosEliminados += 1
                                Exit For
                            Catch ex As IOException
                                If intento = 3 Then
                                    erroresEliminacion += 1
                                    System.Diagnostics.Debug.WriteLine($"No se pudo eliminar PDF temporal tras 3 intentos: {recibo.RutaPDF}")
                                Else
                                    System.Threading.Thread.Sleep(200) ' Esperar antes de reintentar
                                End If
                            End Try
                        Next
                    End If
                Catch fileEx As Exception
                    erroresEliminacion += 1
                    System.Diagnostics.Debug.WriteLine($"Error eliminando PDF temporal: {fileEx.Message}")
                End Try
            Next

            ' Intentar limpiar carpeta temporal si está vacía
            Try
                Dim carpetaTemporal As String = Path.Combine(Path.GetTempPath(), "COOPDIASAM_Temp_PDFs_Extra")
                If Directory.Exists(carpetaTemporal) Then
                    Dim archivosRestantes = Directory.GetFiles(carpetaTemporal)
                    If archivosRestantes.Length = 0 Then
                        Directory.Delete(carpetaTemporal)
                        System.Diagnostics.Debug.WriteLine("Carpeta temporal eliminada correctamente")
                    ElseIf archivosRestantes.Length <= 3 Then
                        ' Si quedan pocos archivos, intentar eliminarlos
                        For Each archivo In archivosRestantes
                            Try
                                File.Delete(archivo)
                            Catch
                                ' Error silencioso
                            End Try
                        Next

                        ' Intentar eliminar carpeta nuevamente
                        Try
                            Directory.Delete(carpetaTemporal)
                        Catch
                            ' Error silencioso
                        End Try
                    End If
                End If
            Catch dirEx As Exception
                System.Diagnostics.Debug.WriteLine($"Error limpiando carpeta temporal: {dirEx.Message}")
            End Try

            ' Log de limpieza
            If archivosEliminados > 0 OrElse erroresEliminacion > 0 Then
                System.Diagnostics.Debug.WriteLine($"Limpieza de PDFs temporales: {archivosEliminados} eliminados, {erroresEliminacion} errores")
            End If

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error general en limpieza de PDFs temporales: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Valida que un recibo esté listo para envío masivo
    ''' </summary>
    Private Function ValidarReciboParaEnvio(row As DataGridViewRow) As String
        Try
            Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
            If String.IsNullOrEmpty(numeroRecibo) Then
                Return "Sin número de recibo - debe registrar el pago primero"
            End If

            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If apartamento Is Nothing Then
                Return "Apartamento no encontrado en la base de datos"
            End If

            If String.IsNullOrWhiteSpace(apartamento.Correo) Then
                Return "Apartamento sin correo electrónico registrado"
            End If

            ' Validar formato de correo básico
            If Not apartamento.Correo.Contains("@") OrElse Not apartamento.Correo.Contains(".") Then
                Return "Formato de correo electrónico inválido"
            End If

            Dim tipoPago As String = ObtenerValorCelda(row.Cells("TipoPagoExtra").Value)
            If String.IsNullOrEmpty(tipoPago) Then
                Return "Sin tipo de pago especificado"
            End If

            Dim valor As Decimal = ConvertirADecimal(row.Cells("ValorPagoExtra").Value)
            If valor <= 0 Then
                Return "Valor de pago inválido"
            End If

            Return String.Empty ' Sin errores

        Catch ex As Exception
            Return $"Error de validación: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' Obtiene estadísticas rápidas para mostrar antes del envío
    ''' </summary>
    Private Function ObtenerEstadisticasEnvio(recibos As List(Of DatosEnvioRecibo)) As String
        Try
            If recibos Is Nothing OrElse recibos.Count = 0 Then
                Return "Sin recibos para analizar"
            End If

            ' Calcular tiempo estimado (8 correos por lote + pausas)
            Dim lotes As Integer = Math.Ceiling(recibos.Count / 8.0)
            Dim tiempoEstimadoMinutos As Integer = (lotes * 2) + Math.Ceiling(recibos.Count * 0.1)

            ' Crear estadísticas usando concatenación simple
            Dim estadisticas As String = ""
            estadisticas = estadisticas & "📊 ANÁLISIS PRE-ENVÍO:" & vbCrLf
            estadisticas = estadisticas & "📄 Total recibos: " & recibos.Count.ToString() & vbCrLf

            ' Contar apartamentos únicos de forma simple
            Dim apartamentosUnicos As New List(Of Integer)
            For Each recibo In recibos
                If Not apartamentosUnicos.Contains(recibo.IdApartamento) Then
                    apartamentosUnicos.Add(recibo.IdApartamento)
                End If
            Next
            estadisticas = estadisticas & "🏠 Apartamentos únicos: " & apartamentosUnicos.Count.ToString() & vbCrLf

            ' Agrupar tipos de pago de forma simple
            Dim tiposPagoDict As New Dictionary(Of String, Integer)
            For Each recibo In recibos
                If tiposPagoDict.ContainsKey(recibo.TipoPago) Then
                    tiposPagoDict(recibo.TipoPago) += 1
                Else
                    tiposPagoDict(recibo.TipoPago) = 1
                End If
            Next

            Dim distribucionTexto As String = ""
            For Each kvp In tiposPagoDict
                If distribucionTexto <> "" Then
                    distribucionTexto = distribucionTexto & ", "
                End If
                distribucionTexto = distribucionTexto & kvp.Key & ": " & kvp.Value.ToString()
            Next

            estadisticas = estadisticas & "💳 Distribución: " & distribucionTexto & vbCrLf
            estadisticas = estadisticas & "⏱️ Tiempo estimado: " & tiempoEstimadoMinutos.ToString() & " minutos" & vbCrLf
            estadisticas = estadisticas & "📦 Lotes a procesar: " & lotes.ToString()

            Return estadisticas

        Catch ex As Exception
            Return "Error calculando estadísticas: " & ex.Message
        End Try
    End Function
End Class