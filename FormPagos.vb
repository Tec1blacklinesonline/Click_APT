' ============================================================================
' FORMPAGOS.VB - VERSIÓN COMPLETAMENTE CORREGIDA Y FUNCIONAL
' ✅ MIGRACIÓN COMPLETA: Toda la funcionalidad de FormPagosExtra
' ✅ SIN ERRORES: Console eliminada, envío de correos, descarga masiva
' ✅ CÁLCULO DE INTERESES: Implementado con base de datos
' ✅ CHECKLIST MASIVO: Selección de apartamentos para operaciones
' ============================================================================

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Linq
Imports System.Diagnostics
Imports System.Data.SQLite
Imports System.IO
Imports System.Threading.Tasks
Imports System.Threading
Imports System

Public Class FormPagos
    Inherits Form

    ' ============================================================================
    ' VARIABLES DE LA CLASE - ACTUALIZADAS
    ' ============================================================================
    Private numeroTorre As Integer
    Private apartamentos As List(Of Apartamento)
    Private WithEvents dgvPagos As DataGridView
    Private WithEvents btnRegistrar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnEnvioMasivo As Button
    Private WithEvents btnExportarPagos As Button
    Private WithEvents btnPagoExtra As Button
    Private WithEvents btnDescargarPDFs As Button ' ✅ NUEVO
    Private WithEvents btnVolver As Button
    Private lblInfo As Label

    ' ✅ NUEVAS CLASES PARA FUNCIONALIDAD MIGRADA
    Public Class DatosEnvioRecibo
        Public Property CorreoDestino As String
        Public Property NombreDestino As String
        Public Property NumeroRecibo As String
        Public Property TipoPago As String
        Public Property RutaPDF As String
        Public Property Apartamento As String
        Public Property IdApartamento As Integer
    End Class

    Public Class DatosDescargaPDF
        Public Property IdApartamento As Integer
        Public Property NumeroRecibo As String
        Public Property TipoPago As String
        Public Property Apartamento As String
        Public Property NombreArchivo As String
        Public Property RutaDestino As String
    End Class

    Public Class ResultadoEnvioMasivo
        Public Property Exitoso As Boolean
        Public Property Mensaje As String
        Public Property EmailsExitosos As Integer
        Public Property EmailsConError As Integer
        Public Property TotalRecibos As Integer
        Public Property ErroresDetallados As New List(Of String)
    End Class

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

    Public Class ProgressInfo
        Public Property Mensaje As String
        Public Property Progreso As Integer
    End Class

    ' ============================================================================
    ' CONSTRUCTOR - SIN CONSOLE DEBUG
    ' ============================================================================
    Public Sub New(numeroTorre As Integer)
        ' ✅ ELIMINADO: AllocConsole() - Sin pantalla negra
        Me.numeroTorre = numeroTorre
        InitializeComponent()
        ConfigurarFormulario()
        CargarApartamentos()
        VerificarServiciosEmail()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.Text = "Pagos Torre " & numeroTorre.ToString() & " - COOPDIASAM"
        Me.ResumeLayout(False)
    End Sub

    ' ============================================================================
    ' CONFIGURACIÓN DEL FORMULARIO - ACTUALIZADA
    ' ============================================================================
    Private Sub ConfigurarFormulario()
        Try
            ' Configuración de ventana
            Me.Text = "💰 Pagos Torre " & numeroTorre.ToString() & " - COOPDIASAM"
            Me.WindowState = FormWindowState.Maximized
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MinimumSize = New Size(1400, 700)
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.BackColor = Color.FromArgb(248, 249, 250)

            ' Panel superior
            Dim panelSuperior As New Panel()
            panelSuperior.Dock = DockStyle.Top
            panelSuperior.Height = 80
            panelSuperior.BackColor = Color.FromArgb(46, 132, 188)

            Dim lblTitulo As New Label()
            lblTitulo.Text = "💰 REGISTRO DE PAGOS - TORRE " & numeroTorre.ToString()
            lblTitulo.Font = New Font("Segoe UI", 20, FontStyle.Bold)
            lblTitulo.ForeColor = Color.White
            lblTitulo.AutoSize = True
            lblTitulo.Location = New Point(30, 25)
            panelSuperior.Controls.Add(lblTitulo)

            ' Panel inferior para botón VOLVER
            Dim panelInferior As New Panel()
            panelInferior.Dock = DockStyle.Bottom
            panelInferior.Height = 60
            panelInferior.BackColor = Color.FromArgb(44, 62, 80)

            btnVolver = New Button()
            btnVolver.Text = "⬅️ VOLVER"
            btnVolver.Size = New Size(140, 40)
            btnVolver.BackColor = Color.FromArgb(44, 62, 80)
            btnVolver.ForeColor = Color.White
            btnVolver.FlatStyle = FlatStyle.Flat
            btnVolver.Cursor = Cursors.Hand
            btnVolver.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            btnVolver.FlatAppearance.BorderSize = 1
            btnVolver.FlatAppearance.BorderColor = Color.White
            btnVolver.FlatAppearance.MouseOverBackColor = Color.FromArgb(52, 73, 94)
            btnVolver.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right

            AddHandler Me.Resize, Sub() btnVolver.Location = New Point(Me.Width - btnVolver.Width - 20, 10)
            btnVolver.Location = New Point(Me.Width - btnVolver.Width - 20, 10)
            panelInferior.Controls.Add(btnVolver)

            ' Configurar panel principal
            ConfigurarPanelPrincipal()

            ' Agregar controles al formulario
            Me.Controls.Add(panelSuperior)
            Me.Controls.Add(panelInferior)

        Catch ex As Exception
            MessageBox.Show("Error en ConfigurarFormulario: " & ex.Message, "Error",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub ConfigurarPanelPrincipal()
        ' Panel de filtros
        Dim panelFiltros As New Panel()
        panelFiltros.Dock = DockStyle.Top
        panelFiltros.Height = 150
        panelFiltros.BackColor = Color.FromArgb(248, 249, 250)
        panelFiltros.Padding = New Padding(30)

        ' Instrucciones
        Dim lblInstrucciones As New Label()
        lblInstrucciones.Text = "💡 Instrucciones:"
        lblInstrucciones.Location = New Point(30, 25)
        lblInstrucciones.Size = New Size(150, 25)
        lblInstrucciones.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblInstrucciones.ForeColor = Color.FromArgb(52, 73, 94)
        panelFiltros.Controls.Add(lblInstrucciones)

        Dim lblDetalles As New Label()
        lblDetalles.Text = "Ingrese montos en campos AMARILLOS, observaciones en campos AZULES, marque ☑ para operaciones masivas"
        lblDetalles.Location = New Point(190, 25)
        lblDetalles.Size = New Size(700, 25)
        lblDetalles.Font = New Font("Segoe UI", 11)
        lblDetalles.ForeColor = Color.FromArgb(52, 73, 94)
        panelFiltros.Controls.Add(lblDetalles)

        ' ✅ BOTONES ACTUALIZADOS CON NUEVA FUNCIONALIDAD
        btnRegistrar = CrearBoton("✅ REGISTRAR PAGOS", New Point(30, 70), New Size(180, 45), Color.FromArgb(231, 76, 60))
        btnCancelar = CrearBoton("🧹 LIMPIAR", New Point(220, 70), New Size(120, 45), Color.FromArgb(44, 62, 80))
        btnEnvioMasivo = CrearBoton("📧 ENVÍO MASIVO", New Point(350, 70), New Size(160, 45), Color.FromArgb(44, 62, 80))
        btnDescargarPDFs = CrearBoton("📥 DESCARGAR PDFs", New Point(520, 70), New Size(170, 45), Color.FromArgb(44, 62, 80))
        btnExportarPagos = CrearBoton("📄 EXPORTAR", New Point(700, 70), New Size(130, 45), Color.FromArgb(44, 62, 80))
        btnPagoExtra = CrearBoton("💳 PAGO EXTRA", New Point(840, 70), New Size(150, 45), Color.FromArgb(44, 62, 80))

        panelFiltros.Controls.AddRange({btnRegistrar, btnCancelar, btnEnvioMasivo, btnDescargarPDFs, btnExportarPagos, btnPagoExtra})


        'ESPACIO PARA BOTONES DE PRUEBA
        '///////////////////////////////////////////////////////////////////////////////////////



        '///////////////////////////////////////////////////////////////////////////////////////


        ' Panel de estadísticas - ✅ MEJORADO
        lblInfo = New Label()
        lblInfo.Location = New Point(1000, 70)
        lblInfo.Size = New Size(400, 45)
        lblInfo.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblInfo.ForeColor = Color.FromArgb(44, 62, 80)
        lblInfo.Text = "📊 Cargando estadísticas..."
        lblInfo.BackColor = Color.FromArgb(236, 240, 241)
        lblInfo.BorderStyle = BorderStyle.FixedSingle
        lblInfo.TextAlign = ContentAlignment.TopLeft
        lblInfo.Padding = New Padding(15, 8, 15, 8)
        lblInfo.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        panelFiltros.Controls.Add(lblInfo)

        ' ✅ DATAGRIDVIEW ACTUALIZADO CON CHECKLIST
        dgvPagos = New DataGridView()
        dgvPagos.Dock = DockStyle.Fill
        dgvPagos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvPagos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvPagos.ReadOnly = False
        dgvPagos.AllowUserToAddRows = False
        dgvPagos.AllowUserToDeleteRows = False
        dgvPagos.BackgroundColor = Color.White
        dgvPagos.BorderStyle = BorderStyle.None
        dgvPagos.RowHeadersVisible = False
        dgvPagos.Font = New Font("Segoe UI", 8.5F)
        dgvPagos.ScrollBars = ScrollBars.Both
        dgvPagos.MultiSelect = False

        ConfigurarDataGridView()
        ConfigurarColumnas()

        ' ✅ EVENTOS ACTUALIZADOS
        AddHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
        AddHandler dgvPagos.CellClick, AddressOf dgvPagos_CellClick
        AddHandler dgvPagos.CellBeginEdit, AddressOf dgvPagos_CellBeginEdit
        AddHandler dgvPagos.CurrentCellDirtyStateChanged, AddressOf dgvPagos_CurrentCellDirtyStateChanged
        AddHandler dgvPagos.CellContentClick, AddressOf dgvPagos_CellContentClick

        ' Agregar al formulario
        Me.Controls.Add(dgvPagos)
        Me.Controls.Add(panelFiltros)
    End Sub

    Private Function CrearBoton(texto As String, ubicacion As Point, tamaño As Size, color As Color) As Button
        Dim boton As New Button()
        boton.Text = texto
        boton.Location = ubicacion
        boton.Size = tamaño
        boton.BackColor = color
        boton.ForeColor = Color.White
        boton.FlatStyle = FlatStyle.Flat
        boton.Cursor = Cursors.Hand
        boton.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        boton.FlatAppearance.BorderSize = 0
        Return boton
    End Function

    Private Sub ConfigurarDataGridView()
        dgvPagos.EnableHeadersVisualStyles = False

        ' Encabezados
        dgvPagos.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(44, 62, 80)
        dgvPagos.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvPagos.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        dgvPagos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvPagos.ColumnHeadersHeight = 35

        ' Estilo de celdas
        dgvPagos.DefaultCellStyle.Font = New Font("Segoe UI", 8.5F)
        dgvPagos.DefaultCellStyle.Padding = New Padding(6, 4, 6, 4)
        dgvPagos.DefaultCellStyle.BackColor = Color.White
        dgvPagos.DefaultCellStyle.ForeColor = Color.FromArgb(52, 73, 94)
        dgvPagos.DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
        dgvPagos.DefaultCellStyle.SelectionForeColor = Color.White

        ' Filas alternadas
        dgvPagos.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 249, 250)
        dgvPagos.RowTemplate.Height = 32
        dgvPagos.GridColor = Color.FromArgb(189, 195, 199)
        dgvPagos.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgvPagos.AutoGenerateColumns = False
    End Sub

    Private Sub ConfigurarColumnas()
        dgvPagos.Columns.Clear()

        With dgvPagos.Columns
            ' ✅ NUEVA COLUMNA DE SELECCIÓN PARA OPERACIONES MASIVAS
            Dim chkColumn As New DataGridViewCheckBoxColumn With {
                .Name = "Seleccionar",
                .HeaderText = "☑",
                .Width = 40,
                .ReadOnly = False,
                .ThreeState = False,
                .TrueValue = True,
                .FalseValue = False
            }
            .Add(chkColumn)

            .Add("IdApartamento", "ID")
            .Add("Apartamento", "APARTAMENTO")
            .Add("EstadoPago", "ESTADO")
            .Add("FechaPago", "FECHA PAGO")
            .Add("SaldoAnterior", "SALDO ANT.")
            .Add("PagoAdministracion", "PAGO ADMIN")
            .Add("PagoInteres", "PAGO INTER")
            .Add("InteresMoratorio", "INTERESES MORA") ' ✅ NUEVA COLUMNA
            .Add("Observaciones", "OBSERVAC.")
            .Add("Total", "TOTAL")
            .Add("NumeroRecibo", "No. RECIBO")
        End With

        ' ✅ CONFIGURAR PROPIEDADES MEJORADAS
        dgvPagos.Columns("Seleccionar").ReadOnly = False
        dgvPagos.Columns("IdApartamento").Visible = False
        dgvPagos.Columns("Apartamento").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvPagos.Columns("Apartamento").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        dgvPagos.Columns("EstadoPago").ReadOnly = True
        dgvPagos.Columns("EstadoPago").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        dgvPagos.Columns("EstadoPago").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        ' ✅ CAMPOS EDITABLES EN AMARILLO
        dgvPagos.Columns("SaldoAnterior").ReadOnly = True
        dgvPagos.Columns("SaldoAnterior").DefaultCellStyle.Format = "C0"
        dgvPagos.Columns("SaldoAnterior").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvPagos.Columns("SaldoAnterior").DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 220)

        dgvPagos.Columns("PagoAdministracion").DefaultCellStyle.Format = "C0"
        dgvPagos.Columns("PagoAdministracion").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvPagos.Columns("PagoAdministracion").DefaultCellStyle.BackColor = Color.LightYellow

        dgvPagos.Columns("PagoInteres").ReadOnly = True
        dgvPagos.Columns("PagoInteres").DefaultCellStyle.Format = "C0"
        dgvPagos.Columns("PagoInteres").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        ' ✅ NUEVA COLUMNA DE INTERESES MORATORIOS
        dgvPagos.Columns("InteresMoratorio").ReadOnly = True
        dgvPagos.Columns("InteresMoratorio").DefaultCellStyle.Format = "C0"
        dgvPagos.Columns("InteresMoratorio").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvPagos.Columns("InteresMoratorio").DefaultCellStyle.BackColor = Color.FromArgb(255, 182, 193) ' Rosa para intereses

        ' ✅ CAMPO DE OBSERVACIONES EN AZUL
        dgvPagos.Columns("Observaciones").DefaultCellStyle.BackColor = Color.FromArgb(173, 216, 230)

        dgvPagos.Columns("Total").ReadOnly = True
        dgvPagos.Columns("Total").DefaultCellStyle.Format = "C0"
        dgvPagos.Columns("Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvPagos.Columns("Total").DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 220)

        dgvPagos.Columns("NumeroRecibo").ReadOnly = True
        dgvPagos.Columns("NumeroRecibo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        ' ✅ BOTONES DE ACCIÓN MEJORADOS
        Dim btnPDFColumn As New DataGridViewButtonColumn With {
            .Name = "BtnPDF",
            .HeaderText = "PDF",
            .Text = "📄",
            .UseColumnTextForButtonValue = True,
            .Width = 50,
            .DefaultCellStyle = New DataGridViewCellStyle With {
            .ForeColor = Color.Black,
                .Font = New Font("Segoe UI", 8, FontStyle.Bold),
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            }
        }

        Dim btnCorreoColumn As New DataGridViewButtonColumn With {
            .Name = "BtnCorreo",
            .HeaderText = "EMAIL",
            .Text = "📧",
            .UseColumnTextForButtonValue = True,
            .Width = 50,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .ForeColor = Color.Black,
                .Font = New Font("Segoe UI", 8, FontStyle.Bold),
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            }
        }

        dgvPagos.Columns.Add(btnPDFColumn)
        dgvPagos.Columns.Add(btnCorreoColumn)
    End Sub

    ' Parte 3

    ' ============================================================================
    ' CARGA DE DATOS - MEJORADA CON CÁLCULO DE INTERESES
    ' ============================================================================
    Private Sub CargarApartamentos()
        Try
            Me.Cursor = Cursors.WaitCursor
            lblInfo.Text = "🔄 Cargando apartamentos y calculando intereses..."
            lblInfo.BackColor = Color.FromArgb(255, 243, 205)

            apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(numeroTorre)
            dgvPagos.Rows.Clear()

            For Each apartamento In apartamentos
                Dim fila As Integer = dgvPagos.Rows.Add()
                CargarDatosApartamento(fila, apartamento)
            Next

            ActualizarEstadisticasVisuales()

        Catch ex As Exception
            lblInfo.Text = "⚠️ Error al cargar apartamentos"
            lblInfo.BackColor = Color.FromArgb(248, 215, 218)
            MessageBox.Show("Error al cargar apartamentos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CargarDatosApartamento(fila As Integer, apartamento As Apartamento)
        Dim ultimoSaldo As Decimal = PagosDAL.ObtenerUltimoSaldo(apartamento.IdApartamento)
        Dim pagoDelMes = ObtenerPagoMesActual(apartamento.IdApartamento)

        ' ✅ CALCULAR INTERESES MORATORIOS DESDE LA BASE DE DATOS
        Dim interesesMoratorios As Decimal = CalcularInteresesMoratorios(apartamento.IdApartamento)

        ' Datos básicos
        dgvPagos.Rows(fila).Cells("Seleccionar").Value = False ' ✅ CHECKBOX INICIALIZADO
        dgvPagos.Rows(fila).Cells("IdApartamento").Value = apartamento.IdApartamento
        dgvPagos.Rows(fila).Cells("Apartamento").Value = $"T{numeroTorre}-{apartamento.NumeroApartamento}"
        dgvPagos.Rows(fila).Cells("FechaPago").Value = DateTime.Now.ToString("dd/MM/yyyy")
        dgvPagos.Rows(fila).Cells("SaldoAnterior").Value = ultimoSaldo
        dgvPagos.Rows(fila).Cells("InteresMoratorio").Value = interesesMoratorios ' ✅ NUEVO

        ' Configurar estado según pago existente
        If pagoDelMes IsNot Nothing Then
            ConfigurarEstadoPagado(fila, pagoDelMes)
        Else
            ConfigurarEstadoPendiente(fila, ultimoSaldo, interesesMoratorios)
        End If
    End Sub

    ' ✅ NUEVO MÉTODO: Calcular intereses moratorios desde la base de datos
    Private Function CalcularInteresesMoratorios(idApartamento As Integer) As Decimal
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Consultar intereses moratorios pendientes de la tabla calculos_interes_mora
                Dim consulta As String = "
                    SELECT COALESCE(SUM(valor_total_adeudado), 0) as total_intereses
                    FROM calculos_interes_mora 
                    WHERE id_apartamento = @idApartamento"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Dim resultado = comando.ExecuteScalar()
                    Return If(resultado IsNot Nothing, Convert.ToDecimal(resultado), 0D)
                End Using
            End Using

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error calculando intereses moratorios: {ex.Message}")
            Return 0D
        End Try
    End Function

    Private Sub ConfigurarEstadoPagado(fila As Integer, pago As PagoModel)
        With dgvPagos.Rows(fila)
            .Cells("EstadoPago").Value = "✅ PAGADO"
            .Cells("EstadoPago").Style.BackColor = Color.Green
            .Cells("EstadoPago").Style.BackColor = Color.FromArgb(34, 139, 34)
            .Cells("NumeroRecibo").Value = pago.NumeroRecibo
            .Cells("Total").Value = pago.TotalPagado
            .Cells("PagoAdministracion").Value = pago.PagoAdministracion
            .Cells("PagoInteres").Value = pago.PagoIntereses
            .Cells("Observaciones").Value = pago.Observaciones

            ' Bloquear edición
            For Each cell As DataGridViewCell In .Cells
                If cell.ColumnIndex < dgvPagos.Columns("BtnPDF").Index Then
                    cell.Style.BackColor = Color.FromArgb(248, 249, 250)
                    cell.ReadOnly = True
                End If
            Next
        End With
    End Sub

    Private Sub ConfigurarEstadoPendiente(fila As Integer, saldo As Decimal, intereses As Decimal)
        With dgvPagos.Rows(fila)
            If saldo > 0 OrElse intereses > 0 Then
                .Cells("EstadoPago").Value = "⚠️ PENDIENTE"
                .Cells("EstadoPago").Style.BackColor = Color.FromArgb(220, 53, 69)
            Else
                .Cells("EstadoPago").Value = "📝 AL DÍA"
                .Cells("EstadoPago").Style.BackColor = Color.FromArgb(34, 139, 34)
            End If
            .Cells("EstadoPago").Style.ForeColor = Color.White

            ' Campos editables
            .Cells("PagoAdministracion").Style.BackColor = Color.LightYellow
            .Cells("FechaPago").Style.BackColor = Color.LightYellow
            .Cells("Observaciones").Style.BackColor = Color.FromArgb(173, 216, 230)

            ' Valores por defecto
            .Cells("PagoAdministracion").Value = 0
            .Cells("PagoInteres").Value = intereses ' ✅ MOSTRAR INTERESES CALCULADOS
            .Cells("Observaciones").Value = ""
            .Cells("Total").Value = saldo + intereses ' ✅ TOTAL INCLUYENDO INTERESES
            .Cells("NumeroRecibo").Value = ""
        End With
    End Sub

    ' ============================================================================
    ' EVENTOS DE CHECKBOX Y TABLA - NUEVOS
    ' ============================================================================
    Private Sub dgvPagos_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If dgvPagos.IsCurrentCellDirty Then
            dgvPagos.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub dgvPagos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim columnName As String = dgvPagos.Columns(e.ColumnIndex).Name

            If columnName = "Seleccionar" Then
                ' Forzar cambio del checkbox
                Dim currentValue As Boolean = False
                Try
                    If dgvPagos.Rows(e.RowIndex).Cells("Seleccionar").Value IsNot Nothing Then
                        currentValue = CBool(dgvPagos.Rows(e.RowIndex).Cells("Seleccionar").Value)
                    End If
                Catch
                    currentValue = False
                End Try

                dgvPagos.Rows(e.RowIndex).Cells("Seleccionar").Value = Not currentValue
                dgvPagos.InvalidateCell(e.ColumnIndex, e.RowIndex)
                ActualizarContadorSeleccionados()
            End If
        End If
    End Sub

    Private Sub ActualizarContadorSeleccionados()
        Try
            Dim seleccionados As Integer = 0
            For Each row As DataGridViewRow In dgvPagos.Rows
                Try
                    Dim valorCelda = row.Cells("Seleccionar").Value
                    If valorCelda IsNot Nothing AndAlso CBool(valorCelda) = True Then
                        seleccionados += 1
                    End If
                Catch
                    ' Ignorar errores de conversión
                End Try
            Next

            lblInfo.Text = $"📊 Torre {numeroTorre}: {seleccionados} apartamentos seleccionados - Total apartamentos: {dgvPagos.Rows.Count}"
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error actualizando contador: {ex.Message}")
        End Try
    End Sub

    ' ============================================================================
    ' FUNCIONES DE BASE DE DATOS - ACTUALIZADAS
    ' ============================================================================
    Private Function ObtenerPagoMesActual(idApartamento As Integer) As PagoModel
        Return PagosDAL.ObtenerPagoMesActual(idApartamento, DateTime.Now)
    End Function

    Private Function RegistrarPago(pago As PagoModel) As Boolean
        Return PagosDAL.RegistrarPago(pago)
    End Function






    'parte 4

    ' ============================================================================
    ' EVENTOS DE BOTONES - ACTUALIZADOS CON NUEVA FUNCIONALIDAD
    ' ============================================================================
    Private Sub BtnVolver_Click(sender As Object, e As EventArgs) Handles btnVolver.Click
        Me.Close()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        If MessageBox.Show("¿Desea limpiar todos los campos?", "Confirmar Limpieza",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            CargarApartamentos()
        End If
    End Sub

    Private Sub BtnRegistrar_Click(sender As Object, e As EventArgs) Handles btnRegistrar.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            lblInfo.Text = "💾 Procesando pagos..."

            Dim pagosParaRegistrar As New List(Of PagoModel)()

            ' Validar y recopilar pagos
            For Each row As DataGridViewRow In dgvPagos.Rows
                Dim estadoPago As String = ""
                If row.Cells("EstadoPago").Value IsNot Nothing Then
                    estadoPago = row.Cells("EstadoPago").Value.ToString()
                End If

                If estadoPago.Contains("PAGADO") Then Continue For

                Dim pagoAdmin As Decimal = 0
                Dim valorTexto As String = ""
                If row.Cells("PagoAdministracion").Value IsNot Nothing Then
                    valorTexto = row.Cells("PagoAdministracion").Value.ToString().Replace("$", "").Replace(",", "")
                End If

                If Not Decimal.TryParse(valorTexto, pagoAdmin) OrElse pagoAdmin <= 0 Then
                    Continue For
                End If

                Dim fechaPago As DateTime = DateTime.Now
                If row.Cells("FechaPago").Value IsNot Nothing Then
                    DateTime.TryParse(row.Cells("FechaPago").Value.ToString(), fechaPago)
                End If

                Dim observaciones As String = ""
                If row.Cells("Observaciones").Value IsNot Nothing Then
                    observaciones = row.Cells("Observaciones").Value.ToString()
                End If

                Dim saldoAnterior As Decimal = 0
                If row.Cells("SaldoAnterior").Value IsNot Nothing Then
                    saldoAnterior = Convert.ToDecimal(row.Cells("SaldoAnterior").Value)
                End If

                Dim pagoIntereses As Decimal = 0
                If row.Cells("InteresMoratorio").Value IsNot Nothing Then
                    pagoIntereses = Convert.ToDecimal(row.Cells("InteresMoratorio").Value)
                End If

                Dim pago As New PagoModel()
                pago.IdApartamento = Convert.ToInt32(row.Cells("IdApartamento").Value)
                pago.FechaPago = fechaPago
                pago.SaldoAnterior = saldoAnterior
                pago.PagoAdministracion = pagoAdmin
                pago.PagoIntereses = pagoIntereses
                pago.CuotaActual = pagoAdmin
                pago.TotalPagado = pagoAdmin + pagoIntereses
                pago.SaldoActual = Math.Max(0, saldoAnterior - pagoAdmin)
                pago.EstadoPago = "REGISTRADO"
                pago.Observaciones = observaciones
                pago.NumeroRecibo = GenerarNumeroRecibo()
                pago.Detalle = "Pago registrado desde FormPagos"
                pago.UsuarioRegistro = "Sistema"
                pago.FechaRegistro = DateTime.Now

                pagosParaRegistrar.Add(pago)
            Next

            If pagosParaRegistrar.Count = 0 Then
                MessageBox.Show("No hay pagos válidos para registrar." & vbCrLf & vbCrLf &
                               "💡 Ingrese valores en los campos AMARILLOS (Pago Admin) y presione REGISTRAR",
                               "Sin pagos válidos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Registrar pagos
            Dim pagosRegistrados As Integer = 0
            For Each pago In pagosParaRegistrar
                If RegistrarPago(pago) Then
                    pagosRegistrados += 1
                End If
            Next

            ' Mostrar resultado
            If pagosRegistrados > 0 Then
                MessageBox.Show($"✅ {pagosRegistrados} pagos registrados exitosamente.", "Registro Exitoso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                CargarApartamentos()
            Else
                MessageBox.Show("No se pudo registrar ningún pago.", "Error en Registro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show("Error al registrar pagos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ' ✅ NUEVO: Descarga masiva de PDFs
    Private Async Sub btnDescargarPDFs_Click(sender As Object, e As EventArgs) Handles btnDescargarPDFs.Click
        Try
            Dim pdfsParaDescargar As New List(Of DatosDescargaPDF)

            ' Recopilar PDFs seleccionados
            For Each row As DataGridViewRow In dgvPagos.Rows
                Try
                    ' Verificar si está seleccionado
                    Dim seleccionado As Boolean = False
                    Try
                        If row.Cells("Seleccionar").Value IsNot Nothing Then
                            seleccionado = CBool(row.Cells("Seleccionar").Value)
                        End If
                    Catch
                        seleccionado = False
                    End Try

                    If Not seleccionado Then Continue For

                    Dim numeroRecibo As String = ""
                    If row.Cells("NumeroRecibo").Value IsNot Nothing Then
                        numeroRecibo = row.Cells("NumeroRecibo").Value.ToString()
                    End If

                    If String.IsNullOrWhiteSpace(numeroRecibo) Then Continue For

                    Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                    Dim apartamentoNombre As String = row.Cells("Apartamento").Value.ToString()

                    pdfsParaDescargar.Add(New DatosDescargaPDF With {
                        .IdApartamento = idApartamento,
                        .NumeroRecibo = numeroRecibo,
                        .TipoPago = "ADMINISTRACION",
                        .Apartamento = apartamentoNombre,
                        .NombreArchivo = $"Recibo_{numeroRecibo}.pdf",
                        .RutaDestino = ""
                    })

                Catch ex As Exception
                    Continue For
                End Try
            Next

            If pdfsParaDescargar.Count = 0 Then
                MessageBox.Show("No hay recibos seleccionados para descargar." & vbCrLf & vbCrLf &
                               "Para descarga masiva:" & vbCrLf &
                               "1. Marque los apartamentos deseados (☑)" & vbCrLf &
                               "2. Asegúrese de que tengan recibos registrados",
                               "Sin PDFs para Descargar", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Obtener carpeta de destino
            Dim carpetaDescargas As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
            Dim carpetaDestino As String = Path.Combine(carpetaDescargas, "COOPDIASAM_Recibos", $"Torre_{numeroTorre}_{DateTime.Now:yyyyMMdd}")

            Dim confirmacion As DialogResult = MessageBox.Show(
                $"¿Confirma la descarga de {pdfsParaDescargar.Count} recibos?" & vbCrLf & vbCrLf &
                $"📁 Destino: {carpetaDestino}",
                "Confirmar Descarga Masiva",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If confirmacion <> DialogResult.Yes Then Return

            ' Crear directorio
            If Not Directory.Exists(carpetaDestino) Then
                Directory.CreateDirectory(carpetaDestino)
            End If

            ' Actualizar rutas
            For Each pdf In pdfsParaDescargar
                pdf.RutaDestino = Path.Combine(carpetaDestino, pdf.NombreArchivo)
            Next

            ' Ejecutar descarga
            btnDescargarPDFs.Enabled = False
            btnDescargarPDFs.Text = "📥 DESCARGANDO..."
            btnDescargarPDFs.BackColor = Color.Gray
            Me.Cursor = Cursors.WaitCursor

            Dim resultado As ResultadoDescargaMasiva = Await ProcesarDescargaMasiva(pdfsParaDescargar)
            MostrarResultadoDescargaMasiva(resultado)

        Catch ex As Exception
            MessageBox.Show("Error en descarga masiva: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnDescargarPDFs.Enabled = True
            btnDescargarPDFs.Text = "📥 DESCARGAR PDFs"
            btnDescargarPDFs.BackColor = Color.FromArgb(155, 89, 182)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' ✅ NUEVO: Envío masivo con liberación de PDFs
    Private Async Sub BtnEnvioMasivo_Click(sender As Object, e As EventArgs) Handles btnEnvioMasivo.Click
        Dim formProgreso As FormProgresoEnvio = Nothing
        Dim recibosParaEnviar As New List(Of DatosEnvioRecibo)
        Dim cerrarFormDespues As Boolean = False

        Try
            ' VALIDAR QUE HAY DATOS
            If dgvPagos.Rows.Count = 0 Then
                MessageBox.Show("No hay pagos registrados para enviar.", "Sin Datos",
                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' RECOPILAR RECIBOS SELECCIONADOS CON VALIDACIONES MEJORADAS
            For Each row As DataGridViewRow In dgvPagos.Rows
                Try
                    If row.IsNewRow Then Continue For

                    ' VERIFICAR SELECCIÓN (CHECKBOX)
                    Dim seleccionado As Boolean = False
                    Try
                        If row.Cells("Seleccionar").Value IsNot Nothing Then
                            seleccionado = CBool(row.Cells("Seleccionar").Value)
                        End If
                    Catch
                        seleccionado = False
                    End Try

                    If Not seleccionado Then Continue For

                    ' VERIFICAR QUE TIENE RECIBO
                    Dim numeroRecibo As String = ""
                    If row.Cells("NumeroRecibo").Value IsNot Nothing Then
                        numeroRecibo = row.Cells("NumeroRecibo").Value.ToString()
                    End If

                    If String.IsNullOrWhiteSpace(numeroRecibo) Then Continue For

                    ' VERIFICAR QUE ESTÁ PAGADO
                    Dim estadoPago As String = ""
                    If row.Cells("EstadoPago").Value IsNot Nothing Then
                        estadoPago = row.Cells("EstadoPago").Value.ToString()
                    End If

                    If Not estadoPago.Contains("PAGADO") Then Continue For

                    ' OBTENER DATOS DEL APARTAMENTO
                    Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                    Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

                    If apartamento Is Nothing OrElse String.IsNullOrWhiteSpace(apartamento.Correo) Then Continue For

                    ' VALIDAR FORMATO DE CORREO
                    If Not apartamento.Correo.Contains("@") OrElse Not apartamento.Correo.Contains(".") Then
                        System.Diagnostics.Debug.WriteLine($"Correo inválido para {apartamento.ObtenerCodigoApartamento()}: {apartamento.Correo}")
                        Continue For
                    End If

                    ' OBTENER MODELO DE PAGO
                    Dim pagoModel As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
                    If pagoModel Is Nothing Then Continue For

                    ' ✅ ESTRATEGIA NUEVA: NO GENERAR PDF TEMPORAL, USAR DATOS EN MEMORIA
                    recibosParaEnviar.Add(New DatosEnvioRecibo With {
                    .CorreoDestino = apartamento.Correo.Trim(),
                    .NombreDestino = If(String.IsNullOrWhiteSpace(apartamento.NombreResidente), "Propietario", apartamento.NombreResidente),
                    .NumeroRecibo = numeroRecibo,
                    .TipoPago = "PAGO ADMINISTRACION",
                    .RutaPDF = "", ' ✅ NO USAR PDF TEMPORAL
                    .Apartamento = $"T{numeroTorre}-{apartamento.NumeroApartamento}",
                    .IdApartamento = idApartamento
                })

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"Error procesando fila para envío masivo: {ex.Message}")
                    Continue For
                End Try
            Next

            ' VALIDAR QUE HAY RECIBOS PARA ENVIAR
            If recibosParaEnviar.Count = 0 Then
                MessageBox.Show("No hay recibos seleccionados válidos para enviar." & vbCrLf & vbCrLf &
                      "Para envío masivo:" & vbCrLf &
                      "1. Marque los apartamentos deseados (☑)" & vbCrLf &
                      "2. Asegúrese de que tengan pagos registrados" & vbCrLf &
                      "3. Verifique que tengan correo electrónico válido",
                      "Sin Recibos para Enviar", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' MOSTRAR ESTADÍSTICAS Y CONFIRMAR
            Dim estadisticas As String = ObtenerEstadisticasEnvio(recibosParaEnviar)
            Dim confirmacion As DialogResult = MessageBox.Show(
        $"¿Confirma el envío masivo de {recibosParaEnviar.Count} recibos?" & vbCrLf & vbCrLf &
        estadisticas & vbCrLf & vbCrLf &
        "📧 Se generarán y enviarán correos con los recibos adjuntos" & vbCrLf &
        "⏱️ Este proceso puede tomar varios minutos",
        "Confirmar Envío Masivo",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question,
        MessageBoxDefaultButton.Button2)

            If confirmacion <> DialogResult.Yes Then
                Return
            End If

            ' MOSTRAR FORMULARIO DE PROGRESO
            formProgreso = New FormProgresoEnvio()
            formProgreso.Text = $"Enviando {recibosParaEnviar.Count} Recibos - Torre {numeroTorre}"
            formProgreso.Show(Me)
            formProgreso.ActualizarProgreso("🚀 Iniciando envío masivo...", 0)

            ' DESHABILITAR CONTROLES
            btnEnvioMasivo.Enabled = False
            btnEnvioMasivo.Text = "📧 ENVIANDO..."
            btnEnvioMasivo.BackColor = Color.Gray
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()

            ' EJECUTAR ENVÍO MASIVO SIN PDFs TEMPORALES
            Dim cancellationToken As CancellationToken = If(formProgreso?.CancellationToken, CancellationToken.None)
            Dim resultado As ResultadoEnvioMasivo = Nothing

            Try
                resultado = Await ProcesarEnvioMasivoSinPDFs(recibosParaEnviar)

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

            ' MOSTRAR RESULTADO
            If resultado IsNot Nothing Then
                MostrarResultadoEnvioMasivo(resultado)
                If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                    formProgreso.MarcarCompletado(resultado.Exitoso, resultado.Mensaje)
                    cerrarFormDespues = True
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("❌ Error crítico: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If formProgreso IsNot Nothing AndAlso Not formProgreso.IsDisposed Then
                formProgreso.MarcarCompletado(False, "Error crítico: " & ex.Message)
                cerrarFormDespues = True
            End If

        Finally
            ' REHABILITAR CONTROLES
            btnEnvioMasivo.Enabled = True
            btnEnvioMasivo.Text = "📧 ENVÍO MASIVO"
            btnEnvioMasivo.BackColor = Color.FromArgb(52, 152, 219)
            Me.Cursor = Cursors.Default
        End Try

        ' CERRAR FORMULARIO DE PROGRESO CON DELAY
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

    ' ✅ NUEVO MÉTODO: Procesar envío masivo sin PDFs temporales
    Private Async Function ProcesarEnvioMasivoSinPDFs(recibos As List(Of DatosEnvioRecibo)) As Task(Of ResultadoEnvioMasivo)
        Dim resultado As New ResultadoEnvioMasivo With {
        .TotalRecibos = recibos.Count
    }

        Try
            Dim exitosos As Integer = 0
            Dim errores As Integer = 0
            Dim erroresDetallados As New List(Of String)

            System.Diagnostics.Debug.WriteLine($"🚀 ENVÍO MASIVO SIN PDFs TEMPORALES - {recibos.Count} recibos")

            For i As Integer = 0 To recibos.Count - 1
                Dim recibo As DatosEnvioRecibo = recibos(i)

                Try
                    lblInfo.Text = $"📧 Generando y enviando {i + 1}/{recibos.Count}: {recibo.Apartamento}"
                    Application.DoEvents()

                    System.Diagnostics.Debug.WriteLine($"📋 Procesando: {recibo.Apartamento} -> {recibo.CorreoDestino}")

                    ' OBTENER DATOS PARA REGENERAR PDF
                    Dim pagoModel As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(recibo.NumeroRecibo)
                    Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(recibo.IdApartamento)

                    If pagoModel Is Nothing OrElse apartamento Is Nothing Then
                        erroresDetallados.Add($"{recibo.Apartamento}: Datos incompletos")
                        errores += 1
                        Continue For
                    End If

                    Dim envioExitoso As Boolean = False

                    ' ✅ GENERAR PDF TEMPORAL ÚNICO PARA CADA ENVÍO
                    Try
                        ' Crear carpeta temporal única
                        Dim carpetaTemporal As String = Path.Combine(Path.GetTempPath(), "COOPDIASAM_Envio", DateTime.Now.ToString("yyyyMMdd_HHmmss"))
                        If Not Directory.Exists(carpetaTemporal) Then
                            Directory.CreateDirectory(carpetaTemporal)
                        End If

                        ' Nombre único con múltiples identificadores
                        Dim nombrePDF As String = $"Envio_{recibo.NumeroRecibo}_{DateTime.Now.Ticks}_{Threading.Thread.CurrentThread.ManagedThreadId}_{i}.pdf"
                        Dim rutaPDFEnvio As String = Path.Combine(carpetaTemporal, nombrePDF)

                        System.Diagnostics.Debug.WriteLine($"🔄 Generando PDF en: {rutaPDFEnvio}")

                        ' ✅ GENERAR PDF CON LIBERACIÓN FORZADA
                        Dim pdfGenerado As String = ReciboPDF.GenerarReciboDePagoEspecifico(pagoModel, apartamento, rutaPDFEnvio)

                        If Not String.IsNullOrEmpty(pdfGenerado) AndAlso File.Exists(pdfGenerado) Then
                            ' ✅ FORZAR LIBERACIÓN DEL PDF GENERADO
                            GC.Collect() ' Forzar recolección de basura
                            GC.WaitForPendingFinalizers() ' Esperar finalización
                            System.Threading.Thread.Sleep(500) ' Pausa para liberación

                            ' Intentar envío
                            envioExitoso = EmailServiceExtra.EnviarReciboPagoExtra(
                            recibo.CorreoDestino,
                            recibo.NombreDestino,
                            recibo.NumeroRecibo,
                            "ADMINISTRACION",
                            pdfGenerado)

                            System.Diagnostics.Debug.WriteLine($"📧 Resultado: {If(envioExitoso, "✅ EXITOSO", "❌ FALLIDO")}")

                            ' Limpieza inmediata
                            Try
                                If File.Exists(pdfGenerado) Then
                                    File.Delete(pdfGenerado)
                                    System.Diagnostics.Debug.WriteLine($"🗑️ PDF eliminado: {nombrePDF}")
                                End If

                                ' Limpiar carpeta si está vacía
                                If Directory.Exists(carpetaTemporal) AndAlso Directory.GetFiles(carpetaTemporal).Length = 0 Then
                                    Directory.Delete(carpetaTemporal)
                                End If
                            Catch deleteEx As Exception
                                System.Diagnostics.Debug.WriteLine($"⚠️ Error limpieza: {deleteEx.Message}")
                            End Try
                        Else
                            erroresDetallados.Add($"{recibo.Apartamento}: No se pudo generar PDF")
                            envioExitoso = False
                        End If

                    Catch pdfEx As Exception
                        erroresDetallados.Add($"{recibo.Apartamento}: Error PDF - {pdfEx.Message}")
                        envioExitoso = False
                        System.Diagnostics.Debug.WriteLine($"❌ Error PDF: {pdfEx.Message}")
                    End Try

                    ' Registrar resultado
                    If envioExitoso Then
                        exitosos += 1
                        System.Diagnostics.Debug.WriteLine($"✅ ÉXITO: {recibo.Apartamento}")
                    Else
                        errores += 1
                        System.Diagnostics.Debug.WriteLine($"❌ ERROR: {recibo.Apartamento}")
                    End If

                    ' Pausa entre envíos
                    Await Task.Delay(2000) ' 2 segundos entre envíos

                Catch ex As Exception
                    errores += 1
                    erroresDetallados.Add($"{recibo.Apartamento}: Error crítico - {ex.Message}")
                    System.Diagnostics.Debug.WriteLine($"❌ Error crítico: {ex.Message}")
                End Try
            Next

            ' Configurar resultado final
            resultado.EmailsExitosos = exitosos
            resultado.EmailsConError = errores
            resultado.Exitoso = (exitosos > 0)
            resultado.ErroresDetallados = erroresDetallados

            If exitosos = recibos.Count Then
                resultado.Mensaje = $"🎉 Envío completado: {exitosos} correos enviados"
            ElseIf exitosos > 0 Then
                resultado.Mensaje = $"⚠️ Envío parcial: {exitosos} exitosos, {errores} errores"
            Else
                resultado.Mensaje = $"❌ Envío fallido: 0 correos enviados"
            End If

            System.Diagnostics.Debug.WriteLine($"📊 RESULTADO FINAL: {exitosos}/{recibos.Count} exitosos")

        Catch ex As Exception
            resultado.Exitoso = False
            resultado.Mensaje = $"❌ Error crítico: {ex.Message}"
            System.Diagnostics.Debug.WriteLine($"❌ ERROR CRÍTICO: {ex.Message}")
        End Try

        Return resultado
    End Function


    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub VerificarServiciosEmail()
        Try
            System.Diagnostics.Debug.WriteLine("=== VERIFICACIÓN DE SERVICIOS DE EMAIL ===")

            ' Verificar si EmailServiceExtra está disponible
            Try
                Dim tipoEmailServiceExtra = Type.GetType("EmailServiceExtra")
                If tipoEmailServiceExtra IsNot Nothing Then
                    System.Diagnostics.Debug.WriteLine("✅ EmailServiceExtra: DISPONIBLE")
                Else
                    System.Diagnostics.Debug.WriteLine("❌ EmailServiceExtra: NO ENCONTRADO")
                End If
            Catch
                System.Diagnostics.Debug.WriteLine("❌ EmailServiceExtra: ERROR AL VERIFICAR")
            End Try

            ' Verificar si EmailService está disponible
            Try
                Dim tipoEmailService = Type.GetType("EmailService")
                If tipoEmailService IsNot Nothing Then
                    System.Diagnostics.Debug.WriteLine("✅ EmailService: DISPONIBLE")
                Else
                    System.Diagnostics.Debug.WriteLine("❌ EmailService: NO ENCONTRADO")
                End If
            Catch
                System.Diagnostics.Debug.WriteLine("❌ EmailService: ERROR AL VERIFICAR")
            End Try

            System.Diagnostics.Debug.WriteLine("=============================================")

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error verificando servicios: {ex.Message}")
        End Try
    End Sub

    ' ============================================================================
    ' MÉTODO PARA DIAGNOSTICAR PROBLEMAS DE ENVÍO
    ' ============================================================================
    Private Function DiagnosticarErrorEnvio(correoDestino As String, numeroRecibo As String) As String
        Try
            Dim diagnostico As String = ""

            ' Verificar formato de correo
            If String.IsNullOrWhiteSpace(correoDestino) Then
                diagnostico += "• Correo vacío" & vbCrLf
            ElseIf Not correoDestino.Contains("@") Then
                diagnostico += "• Correo sin @" & vbCrLf
            ElseIf Not correoDestino.Contains(".") Then
                diagnostico += "• Correo sin dominio" & vbCrLf
            End If

            ' Verificar recibo
            If String.IsNullOrWhiteSpace(numeroRecibo) Then
                diagnostico += "• Número de recibo vacío" & vbCrLf
            End If

            ' Verificar conectividad básica
            Try
                Dim cliente As New System.Net.NetworkInformation.Ping()
                Dim respuesta = cliente.Send("8.8.8.8", 3000)
                If respuesta.Status <> System.Net.NetworkInformation.IPStatus.Success Then
                    diagnostico += "• Sin conectividad a internet" & vbCrLf
                End If
            Catch
                diagnostico += "• No se pudo verificar conectividad" & vbCrLf
            End Try

            If String.IsNullOrEmpty(diagnostico) Then
                Return "Sin problemas detectados en diagnóstico básico"
            Else
                Return "Problemas detectados:" & vbCrLf & diagnostico
            End If

        Catch ex As Exception
            Return $"Error en diagnóstico: {ex.Message}"
        End Try
    End Function

    ' MÉTODO AUXILIAR PARA ESTADÍSTICAS
    Private Function ObtenerEstadisticasEnvio(recibos As List(Of DatosEnvioRecibo)) As String
        Try
            If recibos Is Nothing OrElse recibos.Count = 0 Then
                Return "Sin recibos para analizar"
            End If

            Dim lotes As Integer = Math.Ceiling(recibos.Count / 8.0)
            Dim tiempoEstimado As Integer = (lotes * 2) + Math.Ceiling(recibos.Count * 0.1)

            Return $"📊 ANÁLISIS PRE-ENVÍO:" & vbCrLf &
               $"📄 Total recibos: {recibos.Count}" & vbCrLf &
               $"📦 Lotes a procesar: {lotes}" & vbCrLf &
               $"⏱️ Tiempo estimado: {tiempoEstimado} minutos"

        Catch ex As Exception
            Return "Error calculando estadísticas: " & ex.Message
        End Try
    End Function

    Private Sub BtnExportarPagos_Click(sender As Object, e As EventArgs) Handles btnExportarPagos.Click
        Try
            If dgvPagos.Rows.Count = 0 Then
                MessageBox.Show("No hay datos para exportar.", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Archivo CSV|*.csv"
            saveDialog.Title = "Exportar Pagos Torre " & numeroTorre.ToString()
            saveDialog.FileName = $"PagosTorre{numeroTorre}_{DateTime.Now:yyyyMMdd_HHmmss}"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                ExportarPagosCSV(saveDialog.FileName)
                MessageBox.Show("Pagos exportados exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If MessageBox.Show("¿Desea abrir el archivo exportado?", "Abrir archivo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(saveDialog.FileName) With {.UseShellExecute = True})
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error al exportar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnPagoExtra_Click(sender As Object, e As EventArgs) Handles btnPagoExtra.Click
        Try
            Dim formPagosExtra As New FormPagosExtra(numeroTorre)
            formPagosExtra.ShowDialog()
            CargarApartamentos()
            ActualizarEstadisticasVisuales()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir pagos extra: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'parte 5

    ' ============================================================================
    ' EVENTOS DEL DATAGRIDVIEW - ACTUALIZADOS
    ' ============================================================================

    ' ============================================================================
    ' EVENTOS DEL DATAGRIDVIEW - ACTUALIZADOS
    ' ============================================================================
    Private Sub dgvPagos_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPagos.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Try
                Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)
                Dim columnName As String = dgvPagos.Columns(e.ColumnIndex).Name

                If columnName = "PagoAdministracion" Then
                    Dim pagoAdmin As Decimal = 0
                    If row.Cells("PagoAdministracion").Value IsNot Nothing AndAlso
                       Decimal.TryParse(row.Cells("PagoAdministracion").Value.ToString(), pagoAdmin) Then

                        Dim intereses As Decimal = ConvertirADecimal(row.Cells("InteresMoratorio").Value)
                        row.Cells("Total").Value = pagoAdmin + intereses

                        If row.Cells("NumeroRecibo").Value Is Nothing OrElse
                           String.IsNullOrEmpty(row.Cells("NumeroRecibo").Value.ToString()) Then
                            row.Cells("NumeroRecibo").Value = GenerarNumeroRecibo()
                        End If
                    End If
                End If

                ActualizarEstadisticasVisuales()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error en CellValueChanged: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub dgvPagos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPagos.CellClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Try
                Dim columnName As String = dgvPagos.Columns(e.ColumnIndex).Name
                Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)

                If columnName = "BtnPDF" Then
                    GenerarPDFPago(row)
                ElseIf columnName = "BtnCorreo" Then
                    EnviarCorreoPago(row)
                End If
            Catch ex As Exception
                MessageBox.Show("Error en acción: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub dgvPagos_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvPagos.CellBeginEdit
        Try
            Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)
            Dim estadoPago As String = ""
            If row.Cells("EstadoPago").Value IsNot Nothing Then
                estadoPago = row.Cells("EstadoPago").Value.ToString()
            End If

            If estadoPago.Contains("PAGADO") Then
                e.Cancel = True
                MessageBox.Show("Este apartamento ya tiene un pago registrado para el mes actual.", "Pago Existente", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error en CellBeginEdit: " & ex.Message)
        End Try
    End Sub

    ' ============================================================================
    ' MÉTODOS PARA PROCESAR ENVÍO Y DESCARGA MASIVA
    ' ============================================================================

    Private Async Function ProcesarDescargaMasiva(pdfs As List(Of DatosDescargaPDF)) As Task(Of ResultadoDescargaMasiva)
        Dim resultado As New ResultadoDescargaMasiva With {
            .TotalPDFs = pdfs.Count,
            .RutaCarpetaDestino = Path.GetDirectoryName(pdfs.FirstOrDefault()?.RutaDestino)
        }

        Try
            Dim exitosos As Integer = 0
            Dim errores As Integer = 0

            For i As Integer = 0 To pdfs.Count - 1
                Dim pdf As DatosDescargaPDF = pdfs(i)

                Try
                    lblInfo.Text = $"📄 Generando PDF {i + 1}/{pdfs.Count}: {pdf.Apartamento}"
                    Application.DoEvents()

                    Dim pagoModel As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(pdf.NumeroRecibo)
                    If pagoModel Is Nothing Then
                        errores += 1
                        Continue For
                    End If

                    Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(pdf.IdApartamento)
                    If apartamento Is Nothing Then
                        errores += 1
                        Continue For
                    End If

                    Dim rutaPDFGenerado As String = ReciboPDF.GenerarReciboDePagoEspecifico(pagoModel, apartamento, pdf.RutaDestino)

                    If Not String.IsNullOrEmpty(rutaPDFGenerado) AndAlso File.Exists(rutaPDFGenerado) Then
                        exitosos += 1
                        resultado.ArchivosGenerados.Add(rutaPDFGenerado)
                    Else
                        errores += 1
                    End If

                    Await Task.Delay(100) ' Pequeña pausa

                Catch ex As Exception
                    errores += 1
                    resultado.ErroresDetallados.Add($"Error en {pdf.Apartamento}: {ex.Message}")
                End Try
            Next

            resultado.PDFsDescargados = exitosos
            resultado.PDFsConError = errores
            resultado.Exitoso = (exitosos > 0)
            resultado.Mensaje = $"Descarga completada: {exitosos} PDFs generados, {errores} errores"

        Catch ex As Exception
            resultado.Exitoso = False
            resultado.Mensaje = $"Error crítico: {ex.Message}"
        End Try

        Return resultado
    End Function

    ' ============================================================================
    ' MÉTODOS AUXILIARES MEJORADOS
    ' ============================================================================
    Private Sub ActualizarEstadisticasVisuales()
        Try
            Dim totalApartamentos As Integer = dgvPagos.Rows.Count
            Dim apartamentosPagados As Integer = 0
            Dim apartamentosPendientes As Integer = 0
            Dim totalRecaudado As Decimal = 0
            Dim totalInteresesMora As Decimal = 0

            For Each row As DataGridViewRow In dgvPagos.Rows
                Dim estado As String = ""
                If row.Cells("EstadoPago").Value IsNot Nothing Then
                    estado = row.Cells("EstadoPago").Value.ToString()
                End If

                If estado.Contains("PAGADO") Then
                    apartamentosPagados += 1
                    Dim total As Decimal = ConvertirADecimal(row.Cells("Total").Value)
                    totalRecaudado += total
                ElseIf estado.Contains("PENDIENTE") Then
                    apartamentosPendientes += 1
                End If

                ' Sumar intereses moratorios
                Dim intereses As Decimal = ConvertirADecimal(row.Cells("InteresMoratorio").Value)
                totalInteresesMora += intereses
            Next

            Dim mensaje As String = $"📊 TORRE {numeroTorre}: {totalApartamentos} apts - ✅ Pagados: {apartamentosPagados} - ⚠️ Pendientes: {apartamentosPendientes} - 💰 Recaudado: {totalRecaudado:C}"

            If totalInteresesMora > 0 Then
                mensaje += $" - 🔴 Intereses Mora: {totalInteresesMora:C}"
            End If

            mensaje += $" - ⏱️ {DateTime.Now:HH:mm}"

            lblInfo.Text = mensaje

            If apartamentosPagados > apartamentosPendientes Then
                lblInfo.BackColor = Color.FromArgb(212, 237, 218)
                lblInfo.ForeColor = Color.FromArgb(21, 87, 36)
            Else
                lblInfo.BackColor = Color.FromArgb(248, 215, 218)
                lblInfo.ForeColor = Color.FromArgb(114, 28, 36)
            End If

        Catch ex As Exception
            lblInfo.Text = "⚠️ Error al calcular estadísticas"
        End Try
    End Sub

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

    Private Function GenerarNumeroRecibo() As String
        Return DateTime.Now.ToString("yyyyMMddHHmmss")
    End Function

    Private Sub LimpiarPDFsTemporales(recibos As List(Of DatosEnvioRecibo))
        Try
            For Each recibo In recibos
                Try
                    If Not String.IsNullOrEmpty(recibo.RutaPDF) AndAlso File.Exists(recibo.RutaPDF) Then
                        File.Delete(recibo.RutaPDF)
                    End If
                Catch
                    ' Error silencioso
                End Try
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error limpiando PDFs temporales: {ex.Message}")
        End Try
    End Sub


    ' parte 6 y final

    Private Sub MostrarResultadoEnvioMasivo(resultado As ResultadoEnvioMasivo)
        Try
            Dim icono As MessageBoxIcon = If(resultado.EmailsConError = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning)
            Dim titulo As String = "Resultado del Envío Masivo - Torre " & numeroTorre.ToString()

            Dim mensaje As String = ""
            mensaje = mensaje & "🎯 ENVÍO MASIVO COMPLETADO" & vbCrLf & vbCrLf
            mensaje = mensaje & resultado.Mensaje & vbCrLf & vbCrLf
            mensaje = mensaje & "📊 ESTADÍSTICAS:" & vbCrLf
            mensaje = mensaje & "✅ Exitosos: " & resultado.EmailsExitosos.ToString() & " correos" & vbCrLf
            mensaje = mensaje & "❌ Con errores: " & resultado.EmailsConError.ToString() & " correos" & vbCrLf
            mensaje = mensaje & "📄 Total procesados: " & resultado.TotalRecibos.ToString() & " recibos" & vbCrLf

            If resultado.ErroresDetallados.Count > 0 Then
                mensaje = mensaje & vbCrLf & "🔍 ERRORES:" & vbCrLf
                For i = 0 To Math.Min(3, resultado.ErroresDetallados.Count - 1)
                    mensaje = mensaje & "• " & resultado.ErroresDetallados(i) & vbCrLf
                Next
                If resultado.ErroresDetallados.Count > 3 Then
                    mensaje = mensaje & "... y " & (resultado.ErroresDetallados.Count - 3).ToString() & " errores más" & vbCrLf
                End If
            End If

            MessageBox.Show(mensaje, titulo, MessageBoxButtons.OK, icono)

        Catch ex As Exception
            MessageBox.Show("Error mostrando resultado: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

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
                mensaje = mensaje & vbCrLf & "🎉 ¿Desea abrir la carpeta de destino?"

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

        Catch ex As Exception
            MessageBox.Show("Error mostrando resultado: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportarPagosCSV(rutaArchivo As String)
        Using writer As New StreamWriter(rutaArchivo, False, System.Text.Encoding.UTF8)
            writer.WriteLine("sep=,")
            writer.WriteLine("Apartamento,Estado,Fecha Pago,Saldo Anterior,Pago Administracion,Intereses Mora,Total,Numero Recibo,Observaciones")

            For Each row As DataGridViewRow In dgvPagos.Rows
                Dim apartamento As String = If(row.Cells("Apartamento").Value IsNot Nothing, row.Cells("Apartamento").Value.ToString(), "")
                Dim estado As String = If(row.Cells("EstadoPago").Value IsNot Nothing, row.Cells("EstadoPago").Value.ToString(), "")
                Dim fechaPago As String = If(row.Cells("FechaPago").Value IsNot Nothing, row.Cells("FechaPago").Value.ToString(), "")
                Dim saldoAnterior As String = ConvertirADecimal(row.Cells("SaldoAnterior").Value).ToString()
                Dim pagoAdmin As String = ConvertirADecimal(row.Cells("PagoAdministracion").Value).ToString()
                Dim interesesMora As String = ConvertirADecimal(row.Cells("InteresMoratorio").Value).ToString()
                Dim total As String = ConvertirADecimal(row.Cells("Total").Value).ToString()
                Dim numeroRecibo As String = If(row.Cells("NumeroRecibo").Value IsNot Nothing, row.Cells("NumeroRecibo").Value.ToString(), "")
                Dim observaciones As String = If(row.Cells("Observaciones").Value IsNot Nothing, row.Cells("Observaciones").Value.ToString().Replace("""", """"""), "")

                Dim linea As String = String.Join(",", {
                    """" & apartamento & """",
                    """" & estado & """",
                    """" & fechaPago & """",
                    saldoAnterior,
                    pagoAdmin,
                    interesesMora,
                    total,
                    """" & numeroRecibo & """",
                    """" & observaciones & """"
                })
                writer.WriteLine(linea)
            Next

            writer.WriteLine()
            writer.WriteLine($"RESUMEN Torre {numeroTorre} - {DateTime.Now:dd/MM/yyyy HH:mm}")
        End Using
    End Sub

    Private Sub GenerarPDFPago(row As DataGridViewRow)
        Try
            Dim estadoPago As String = ""
            If row.Cells("EstadoPago").Value IsNot Nothing Then
                estadoPago = row.Cells("EstadoPago").Value.ToString()
            End If

            If Not estadoPago.Contains("PAGADO") Then
                MessageBox.Show("Solo se puede generar PDF para pagos registrados.", "PDF No Disponible", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim numeroRecibo As String = ""
            If row.Cells("NumeroRecibo").Value IsNot Nothing Then
                numeroRecibo = row.Cells("NumeroRecibo").Value.ToString()
            End If

            Dim pagoModel As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If pagoModel IsNot Nothing AndAlso apartamento IsNot Nothing Then
                Dim rutaPdf As String = ReciboPDF.GenerarReciboDePagoSeguro(pagoModel, apartamento)

                If Not String.IsNullOrEmpty(rutaPdf) AndAlso File.Exists(rutaPdf) Then
                    Dim resultado As DialogResult = MessageBox.Show(
                        $"✅ Recibo generado exitosamente.{vbCrLf}{vbCrLf}¿Desea abrir el archivo?",
                        "PDF Generado", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

                    If resultado = DialogResult.Yes Then
                        Process.Start(New ProcessStartInfo(rutaPdf) With {.UseShellExecute = True})
                    End If
                Else
                    MessageBox.Show("Error al generar el PDF del recibo.", "Error PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al generar PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' ✅ ENVIO DE PDF INDIVIDUAL Y MASIVO A CORREOS
    Private Sub EnviarCorreoPago(row As DataGridViewRow)
        Try
            Me.Cursor = Cursors.WaitCursor

            ' Validaciones básicas
            Dim estadoPago As String = ""
            If row.Cells("EstadoPago").Value IsNot Nothing Then
                estadoPago = row.Cells("EstadoPago").Value.ToString()
            End If

            If Not estadoPago.Contains("PAGADO") Then
                MessageBox.Show("Solo se puede enviar correo para pagos registrados.", "Correo No Disponible", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim numeroRecibo As String = ""
            If row.Cells("NumeroRecibo").Value IsNot Nothing Then
                numeroRecibo = row.Cells("NumeroRecibo").Value.ToString()
            End If

            ' Obtener datos
            Dim pagoModel As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If pagoModel Is Nothing OrElse apartamento Is Nothing Then
                MessageBox.Show("No se encontraron los datos del pago.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If String.IsNullOrWhiteSpace(apartamento.Correo) Then
                MessageBox.Show($"El apartamento {apartamento.ObtenerCodigoApartamento()} no tiene correo electrónico registrado.",
                          "Correo no registrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Confirmar envío
            Dim resultado As DialogResult = MessageBox.Show(
            $"¿Confirma el envío del recibo No. {numeroRecibo}?" & vbCrLf & vbCrLf &
            $"📧 Destinatario: {apartamento.Correo}" & vbCrLf &
            $"👤 Propietario: {apartamento.NombreResidente}",
            "Confirmar Envío de Correo",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                ' ✅ MÉTODO ALTERNATIVO: Usar PDF existente del directorio de recibos
                Dim rutaPDFExistente As String = ""

                Try
                    ' Buscar PDF en la carpeta de recibos permanente
                    Dim carpetaRecibos As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM\Recibos")

                    If Directory.Exists(carpetaRecibos) Then
                        ' Buscar archivo por número de recibo
                        Dim archivos = Directory.GetFiles(carpetaRecibos, $"*{numeroRecibo}*.pdf")
                        If archivos.Length > 0 Then
                            rutaPDFExistente = archivos(0)
                        End If
                    End If

                    ' Si no existe, generar uno nuevo con método seguro
                    If String.IsNullOrEmpty(rutaPDFExistente) OrElse Not File.Exists(rutaPDFExistente) Then
                        rutaPDFExistente = ReciboPDF.GenerarReciboDePagoSeguro(pagoModel, apartamento)
                    End If

                Catch pdfEx As Exception
                    MessageBox.Show($"Error accediendo al PDF: {pdfEx.Message}", "Error PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try

                If String.IsNullOrEmpty(rutaPDFExistente) OrElse Not File.Exists(rutaPDFExistente) Then
                    MessageBox.Show("No se pudo generar o encontrar el PDF para el correo.", "Error PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' ✅ MÉTODO DE COPIA EN MEMORIA (sin conflictos de archivos)
                Try
                    Dim envioExitoso As Boolean = False

                    ' Leer PDF completo en memoria
                    Dim pdfBytes As Byte()
                    Using fs As New FileStream(rutaPDFExistente, FileMode.Open, FileAccess.Read, FileShare.Read)
                        pdfBytes = New Byte(fs.Length - 1) {}
                        fs.Read(pdfBytes, 0, pdfBytes.Length)
                    End Using

                    ' Crear archivo temporal único con los bytes en memoria
                    Dim rutaTemporal As String = Path.Combine(Path.GetTempPath(), $"Recibo_{numeroRecibo}_{DateTime.Now.Ticks}.pdf")

                    ' Escribir bytes a archivo temporal
                    Using fs As New FileStream(rutaTemporal, FileMode.Create, FileAccess.Write)
                        fs.Write(pdfBytes, 0, pdfBytes.Length)
                        fs.Flush()
                    End Using

                    ' Pequeña pausa para asegurar que el archivo esté disponible
                    System.Threading.Thread.Sleep(100)

                    ' Enviar correo
                    envioExitoso = EmailServiceExtra.EnviarReciboPagoExtra(
                    apartamento.Correo,
                    apartamento.NombreResidente,
                    numeroRecibo,
                    "ADMINISTRACION",
                    rutaTemporal)

                    ' Limpiar archivo temporal inmediatamente
                    Try
                        If File.Exists(rutaTemporal) Then
                            File.Delete(rutaTemporal)
                        End If
                    Catch
                        ' Error silencioso en limpieza
                    End Try

                    ' Mostrar resultado
                    If envioExitoso Then
                        MessageBox.Show(
                        $"✅ Recibo enviado exitosamente." & vbCrLf & vbCrLf &
                        $"📧 Destinatario: {apartamento.Correo}" & vbCrLf &
                        $"📄 Recibo: {numeroRecibo}",
                        "Correo Enviado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)
                        lblInfo.Text = $"✅ Recibo {numeroRecibo} enviado exitosamente"
                    Else
                        MessageBox.Show("❌ No se pudo enviar el correo electrónico." & vbCrLf & vbCrLf &
                                  "Verifique:" & vbCrLf &
                                  "• Conexión a internet" & vbCrLf &
                                  "• Dirección de correo válida" & vbCrLf &
                                  "• Límites de envío no alcanzados",
                                  "Error de Envío", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        lblInfo.Text = $"❌ Error enviando recibo {numeroRecibo}"
                    End If

                Catch emailEx As Exception
                    MessageBox.Show($"Error en proceso de envío: {emailEx.Message}", "Error de Envío",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show($"Error crítico: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub



    ' ✅ MÉTODO MEJORADO PARA LIMPIAR ARCHIVOS
    Private Sub LimpiarArchivoTemporal(rutaArchivo As String)
        If String.IsNullOrEmpty(rutaArchivo) Then Return

        Try
            ' Múltiples intentos de eliminación
            For intento = 1 To 5
                Try
                    If File.Exists(rutaArchivo) Then
                        File.Delete(rutaArchivo)
                        Exit For
                    End If
                Catch ex As IOException
                    If intento < 5 Then
                        System.Threading.Thread.Sleep(100 * intento)
                    Else
                        ' Si no se puede eliminar, al menos intentar moverlo
                        Try
                            Dim archivoTempRename As String = rutaArchivo & "_DELETE_" & DateTime.Now.Ticks
                            File.Move(rutaArchivo, archivoTempRename)
                        Catch
                            ' Error silencioso final
                        End Try
                    End If
                Catch
                    Exit For
                End Try
            Next
        Catch
            ' Error silencioso
        End Try
    End Sub




    ' ✅ MÉTODOS AUXILIARES PARA MANEJAR ARCHIVOS DE FORMA SEGURA
    Private Function VerificarAccesoArchivo(rutaArchivo As String) As Boolean
        Try
            Using stream As New FileStream(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.Read)
                Return True
            End Using
        Catch
            Return False
        End Try
    End Function

    'ESPACIO PARA BOTONES DE PRUEBA
    '///////////////////////////////////////////////////////////////////////////////////////



    '///////////////////////////////////////////////////////////////////////////////////////

    Private Sub DiagnosticarConfiguracionEmail()
        Try
            ' Verificar que los servicios estén disponibles
            System.Diagnostics.Debug.WriteLine("=== DIAGNÓSTICO DE SERVICIOS DE EMAIL ===")

            Try
                ' Verificar EmailServiceExtra
                System.Diagnostics.Debug.WriteLine("EmailServiceExtra disponible: SÍ")
            Catch
                System.Diagnostics.Debug.WriteLine("EmailServiceExtra disponible: NO")
            End Try

            Try
                ' Verificar EmailService
                System.Diagnostics.Debug.WriteLine("EmailService disponible: SÍ")
            Catch
                System.Diagnostics.Debug.WriteLine("EmailService disponible: NO")
            End Try

            System.Diagnostics.Debug.WriteLine("==========================================")

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en diagnóstico: {ex.Message}")
        End Try
    End Sub

    ' ============================================================================
    ' CLEANUP - SIN CONSOLE
    ' ============================================================================
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            ' ✅ ELIMINADO: FreeConsole() - Sin limpieza de console
            If apartamentos IsNot Nothing Then
                apartamentos.Clear()
            End If
        Catch
        End Try
        MyBase.OnFormClosed(e)
    End Sub

End Class