' ============================================================================
' FORM DESCARGAR RECIBOS - PERMITE A LOS USUARIOS DESCARGAR RECIBOS GENERADOS
' ✅ Búsqueda por apartamento, fechas, números de recibo
' ✅ Re-generación de PDFs si no existen los archivos originales
' ✅ VERSIÓN COMPLETA CORREGIDA - ESTILO FORMHISTORIAL APLICADO
' ============================================================================

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.SQLite
Imports System.IO
Imports System.Diagnostics

Public Class FormDescargarRecibos
    Inherits Form

    Private WithEvents dgvRecibos As DataGridView
    Private cmbTorre As ComboBox
    Private cmbApartamento As ComboBox
    Private dtpFechaInicio As DateTimePicker
    Private dtpFechaFin As DateTimePicker
    Private txtNumeroRecibo As TextBox
    Private WithEvents btnBuscar As Button
    Private WithEvents btnLimpiar As Button
    Private WithEvents btnDescargarSeleccionado As Button
    Private WithEvents btnDescargarTodos As Button
    Private WithEvents btnRegenerarPDF As Button
    Private WithEvents btnVolver As Button
    Private lblResultados As Label
    Private panelFiltros As Panel

    Public Sub New()
        InitializeComponent()
        ConfigurarFormulario()
        CargarTorres()
        CargarRecibosRecientes()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.Text = "Descargar Recibos - COOPDIASAM"
        Me.ResumeLayout(False)
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configuración de ventana completa - MISMO ESTILO QUE FORMHISTORIAL
            Me.Text = "Descargar Recibos - COOPDIASAM"
            Me.WindowState = FormWindowState.Maximized
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MinimumSize = New Size(1400, 700)
            Me.MaximizeBox = True
            Me.MinimizeBox = True
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' Panel superior con el MISMO diseño que FormHistorial
            Dim panelSuperior As New Panel()
            panelSuperior.Dock = DockStyle.Top
            panelSuperior.Height = 80
            panelSuperior.BackColor = Color.FromArgb(46, 132, 188)  ' MISMO COLOR

            Dim lblTitulo As New Label()
            lblTitulo.Text = "📥 DESCARGAR RECIBOS"
            lblTitulo.Font = New Font("Segoe UI", 20, FontStyle.Bold)  ' MISMA FUENTE
            lblTitulo.ForeColor = Color.White
            lblTitulo.AutoSize = True
            lblTitulo.Location = New Point(30, 25)
            panelSuperior.Controls.Add(lblTitulo)

            ' Panel inferior para el botón VOLVER - CENTRADO ABAJO (IGUAL QUE FORMHISTORIAL)
            Dim panelInferior As New Panel()
            panelInferior.Dock = DockStyle.Bottom
            panelInferior.Height = 60
            panelInferior.BackColor = Color.FromArgb(44, 62, 80)  ' MISMO COLOR DE LAS COLUMNAS

            btnVolver = New Button()
            btnVolver.Text = "⬅️ VOLVER"
            btnVolver.Size = New Size(140, 40)
            btnVolver.BackColor = Color.FromArgb(44, 62, 80)  ' MISMO COLOR DE LAS COLUMNAS
            btnVolver.ForeColor = Color.White
            btnVolver.FlatStyle = FlatStyle.Flat
            btnVolver.Cursor = Cursors.Hand
            btnVolver.Font = New Font("Segoe UI", 10, FontStyle.Bold)  ' FUENTE MÁS PEQUEÑA
            btnVolver.FlatAppearance.BorderSize = 1
            btnVolver.FlatAppearance.BorderColor = Color.White
            btnVolver.FlatAppearance.MouseOverBackColor = Color.FromArgb(52, 73, 94)
            btnVolver.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right  ' ANCLADO A LA DERECHA

            ' EVENTO PARA MANTENERLO EN LA DERECHA
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
        ' Panel de filtros con mejor organización y altura COMPACTA
        panelFiltros = New Panel()
        panelFiltros.Dock = DockStyle.Top
        panelFiltros.Height = 130  ' ALTURA COMPACTA
        panelFiltros.BackColor = Color.FromArgb(248, 249, 250)  ' MISMO COLOR
        panelFiltros.Padding = New Padding(20)

        ' === PRIMERA FILA DE CONTROLES ===

        ' letrero de Torre
        Dim lblTorre As New Label()
        lblTorre.Text = "Torre:"
        lblTorre.Location = New Point(20, 15)
        lblTorre.Size = New Size(60, 25)
        lblTorre.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblTorre.ForeColor = Color.FromArgb(52, 73, 94)
        lblTorre.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblTorre)

        cmbTorre = New ComboBox()
        cmbTorre.Location = New Point(80, 13)
        cmbTorre.Size = New Size(80, 30)
        cmbTorre.DropDownStyle = ComboBoxStyle.DropDownList
        cmbTorre.Font = New Font("Segoe UI", 10)
        AddHandler cmbTorre.SelectedIndexChanged, AddressOf cmbTorre_SelectedIndexChanged
        panelFiltros.Controls.Add(cmbTorre)

        ' letrero de Apartamento
        Dim lblApartamento As New Label()
        lblApartamento.Text = "Apartamento:"
        lblApartamento.Location = New Point(180, 15)
        lblApartamento.Size = New Size(110, 25)
        lblApartamento.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblApartamento.ForeColor = Color.FromArgb(52, 73, 94)
        lblApartamento.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblApartamento)

        cmbApartamento = New ComboBox()
        cmbApartamento.Location = New Point(290, 13)
        cmbApartamento.Size = New Size(110, 30)
        cmbApartamento.DropDownStyle = ComboBoxStyle.DropDownList
        cmbApartamento.Font = New Font("Segoe UI", 10)
        panelFiltros.Controls.Add(cmbApartamento)

        ' Fechas
        Dim lblFechaInicio As New Label()
        lblFechaInicio.Text = "📅 Desde:"
        lblFechaInicio.Location = New Point(420, 15)
        lblFechaInicio.Size = New Size(90, 25)
        lblFechaInicio.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblFechaInicio.ForeColor = Color.FromArgb(52, 73, 94)
        lblFechaInicio.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblFechaInicio)

        dtpFechaInicio = New DateTimePicker()
        dtpFechaInicio.Location = New Point(510, 13)
        dtpFechaInicio.Size = New Size(120, 30)
        dtpFechaInicio.Format = DateTimePickerFormat.Short
        dtpFechaInicio.Font = New Font("Segoe UI", 10)
        dtpFechaInicio.Value = DateTime.Now.AddMonths(-3)
        panelFiltros.Controls.Add(dtpFechaInicio)

        Dim lblFechaFin As New Label()
        lblFechaFin.Text = "📅 Hasta:"
        lblFechaFin.Location = New Point(650, 15)
        lblFechaFin.Size = New Size(90, 25)
        lblFechaFin.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblFechaFin.ForeColor = Color.FromArgb(52, 73, 94)
        lblFechaFin.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblFechaFin)

        dtpFechaFin = New DateTimePicker()
        dtpFechaFin.Location = New Point(740, 13)
        dtpFechaFin.Size = New Size(120, 30)
        dtpFechaFin.Format = DateTimePickerFormat.Short
        dtpFechaFin.Font = New Font("Segoe UI", 10)
        dtpFechaFin.Value = DateTime.Now
        panelFiltros.Controls.Add(dtpFechaFin)

        ' === SEGUNDA FILA - BÚSQUEDA Y BOTONES ===
        Dim lblNumeroRecibo As New Label()
        lblNumeroRecibo.Text = "No. Recibo:"
        lblNumeroRecibo.Location = New Point(20, 50)
        lblNumeroRecibo.Size = New Size(90, 25)
        lblNumeroRecibo.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblNumeroRecibo.ForeColor = Color.FromArgb(52, 73, 94)
        lblNumeroRecibo.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblNumeroRecibo)

        txtNumeroRecibo = New TextBox()
        txtNumeroRecibo.Location = New Point(110, 48)
        txtNumeroRecibo.Size = New Size(150, 30)
        txtNumeroRecibo.Font = New Font("Segoe UI", 11)
        txtNumeroRecibo.BorderStyle = BorderStyle.FixedSingle
        panelFiltros.Controls.Add(txtNumeroRecibo)

        ' Botones con estilo FormHistorial
        btnBuscar = New Button()
        btnBuscar.Text = "🔍 BUSCAR"
        btnBuscar.Location = New Point(280, 47)
        btnBuscar.Size = New Size(110, 35)
        btnBuscar.BackColor = Color.FromArgb(52, 152, 219)
        btnBuscar.ForeColor = Color.White
        btnBuscar.FlatStyle = FlatStyle.Flat
        btnBuscar.Cursor = Cursors.Hand
        btnBuscar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnBuscar.FlatAppearance.BorderSize = 0
        btnBuscar.FlatAppearance.MouseOverBackColor = Color.FromArgb(41, 128, 185)
        panelFiltros.Controls.Add(btnBuscar)

        btnLimpiar = New Button()
        btnLimpiar.Text = "🧹 LIMPIAR"
        btnLimpiar.Location = New Point(400, 47)
        btnLimpiar.Size = New Size(110, 35)
        btnLimpiar.BackColor = Color.FromArgb(231, 76, 60)
        btnLimpiar.ForeColor = Color.White
        btnLimpiar.FlatStyle = FlatStyle.Flat
        btnLimpiar.Cursor = Cursors.Hand
        btnLimpiar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnLimpiar.FlatAppearance.BorderSize = 0
        btnLimpiar.FlatAppearance.MouseOverBackColor = Color.FromArgb(192, 57, 43)
        panelFiltros.Controls.Add(btnLimpiar)

        ' Botones de descarga compactos
        btnDescargarSeleccionado = New Button()
        btnDescargarSeleccionado.Text = "📥 SELECCIONADO"
        btnDescargarSeleccionado.Location = New Point(520, 47)
        btnDescargarSeleccionado.Size = New Size(150, 35)
        btnDescargarSeleccionado.BackColor = Color.FromArgb(39, 174, 96)
        btnDescargarSeleccionado.ForeColor = Color.White
        btnDescargarSeleccionado.FlatStyle = FlatStyle.Flat
        btnDescargarSeleccionado.Cursor = Cursors.Hand
        btnDescargarSeleccionado.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        btnDescargarSeleccionado.FlatAppearance.BorderSize = 0
        btnDescargarSeleccionado.FlatAppearance.MouseOverBackColor = Color.FromArgb(34, 153, 84)
        panelFiltros.Controls.Add(btnDescargarSeleccionado)

        btnDescargarTodos = New Button()
        btnDescargarTodos.Text = "📦 TODOS"
        btnDescargarTodos.Location = New Point(680, 47)
        btnDescargarTodos.Size = New Size(90, 35)
        btnDescargarTodos.BackColor = Color.FromArgb(142, 68, 173)
        btnDescargarTodos.ForeColor = Color.White
        btnDescargarTodos.FlatStyle = FlatStyle.Flat
        btnDescargarTodos.Cursor = Cursors.Hand
        btnDescargarTodos.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        btnDescargarTodos.FlatAppearance.BorderSize = 0
        btnDescargarTodos.FlatAppearance.MouseOverBackColor = Color.FromArgb(125, 60, 152)
        panelFiltros.Controls.Add(btnDescargarTodos)

        btnRegenerarPDF = New Button()
        btnRegenerarPDF.Text = "🔄 PDF"
        btnRegenerarPDF.Location = New Point(780, 47)
        btnRegenerarPDF.Size = New Size(90, 35)
        btnRegenerarPDF.BackColor = Color.FromArgb(230, 126, 34)
        btnRegenerarPDF.ForeColor = Color.White
        btnRegenerarPDF.FlatStyle = FlatStyle.Flat
        btnRegenerarPDF.Cursor = Cursors.Hand
        btnRegenerarPDF.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        btnRegenerarPDF.FlatAppearance.BorderSize = 0
        btnRegenerarPDF.FlatAppearance.MouseOverBackColor = Color.FromArgb(211, 84, 0)
        panelFiltros.Controls.Add(btnRegenerarPDF)

        ' === PANEL DE ESTADÍSTICAS ===
        lblResultados = New Label()
        lblResultados.Location = New Point(20, 90)
        lblResultados.Size = New Size(800, 35)
        lblResultados.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblResultados.ForeColor = Color.FromArgb(44, 62, 80)
        lblResultados.Text = "📊 Mostrando recibos recientes..."
        lblResultados.BackColor = Color.FromArgb(236, 240, 241)
        lblResultados.BorderStyle = BorderStyle.FixedSingle
        lblResultados.TextAlign = ContentAlignment.MiddleLeft
        lblResultados.Padding = New Padding(10, 5, 10, 5)
        lblResultados.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        panelFiltros.Controls.Add(lblResultados)

        ' DataGridView principal
        dgvRecibos = New DataGridView()
        dgvRecibos.Dock = DockStyle.Fill
        dgvRecibos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvRecibos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvRecibos.MultiSelect = True
        dgvRecibos.AllowUserToAddRows = False
        dgvRecibos.AllowUserToDeleteRows = False
        dgvRecibos.BackgroundColor = Color.White
        dgvRecibos.BorderStyle = BorderStyle.None
        dgvRecibos.RowHeadersVisible = False
        dgvRecibos.Font = New Font("Segoe UI", 10.0F)
        dgvRecibos.ScrollBars = ScrollBars.Both
        dgvRecibos.AllowUserToResizeColumns = True
        ConfigurarDataGridView(dgvRecibos)
        ConfigurarColumnas()

        ' Agregar al formulario
        Me.Controls.Add(dgvRecibos)
        Me.Controls.Add(panelFiltros)
    End Sub

    Private Sub ConfigurarDataGridView(dgv As DataGridView)
        ' Configuración visual moderna y profesional - IGUAL QUE FORMHISTORIAL
        dgv.EnableHeadersVisualStyles = False

        ' Encabezados con diseño moderno - LETRAS MÁS PEQUEÑAS
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(44, 62, 80)  ' MISMO COLOR
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)  ' FUENTE MÁS PEQUEÑA
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(3, 8, 3, 8)
        dgv.ColumnHeadersHeight = 40

        ' Estilo de celdas mejorado - LETRAS MÁS PEQUEÑAS
        dgv.DefaultCellStyle.Font = New Font("Segoe UI", 8.5F)  ' FUENTE MÁS PEQUEÑA
        dgv.DefaultCellStyle.Padding = New Padding(6, 4, 6, 4)
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv.DefaultCellStyle.ForeColor = Color.FromArgb(52, 73, 94)
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
        dgv.DefaultCellStyle.SelectionForeColor = Color.White

        ' Filas alternadas con color suave
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 249, 250)
        dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.FromArgb(52, 73, 94)

        ' Configuración de filas y bordes - FILAS MÁS PEQUEÑAS
        dgv.RowTemplate.Height = 32  ' FILAS MÁS PEQUEÑAS
        dgv.GridColor = Color.FromArgb(189, 195, 199)
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        dgv.AutoGenerateColumns = False

        ' Eventos
        AddHandler dgv.CellDoubleClick, AddressOf dgvRecibos_CellDoubleClick
        AddHandler dgv.CellClick, AddressOf dgvRecibos_CellClick
    End Sub

    Private Sub ConfigurarColumnas()
        dgvRecibos.Columns.Clear()

        With dgvRecibos.Columns
            .Add("IdPago", "ID")
            .Add("NumeroRecibo", "N° RECIBO")
            .Add("FechaPago", "FECHA")
            .Add("Apartamento", "APARTAMENTO")
            .Add("Propietario", "PROPIETARIO")
            .Add("TipoPago", "TIPO")
            .Add("TotalPagado", "VALOR")
            .Add("EstadoArchivo", "ARCHIVO")
            .Add("RutaArchivo", "RUTA")
        End With

        ' Configuración con estilo FormHistorial
        dgvRecibos.Columns("IdPago").HeaderText = "ID"
        dgvRecibos.Columns("IdPago").FillWeight = 7
        dgvRecibos.Columns("IdPago").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvRecibos.Columns("IdPago").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        dgvRecibos.Columns("NumeroRecibo").HeaderText = "N° RECIBO"
        dgvRecibos.Columns("NumeroRecibo").FillWeight = 10
        dgvRecibos.Columns("NumeroRecibo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvRecibos.Columns("NumeroRecibo").DefaultCellStyle.Font = New Font("Consolas", 8, FontStyle.Bold)

        dgvRecibos.Columns("FechaPago").HeaderText = "FECHA"
        dgvRecibos.Columns("FechaPago").FillWeight = 10
        dgvRecibos.Columns("FechaPago").DefaultCellStyle.Format = "dd/MM/yyyy"
        dgvRecibos.Columns("FechaPago").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvRecibos.Columns("FechaPago").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        dgvRecibos.Columns("Apartamento").HeaderText = "APARTAMENTO"
        dgvRecibos.Columns("Apartamento").FillWeight = 12
        dgvRecibos.Columns("Apartamento").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvRecibos.Columns("Apartamento").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        dgvRecibos.Columns("Propietario").HeaderText = "PROPIETARIO"
        dgvRecibos.Columns("Propietario").FillWeight = 20
        dgvRecibos.Columns("Propietario").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        dgvRecibos.Columns("TipoPago").HeaderText = "TIPO"
        dgvRecibos.Columns("TipoPago").FillWeight = 13
        dgvRecibos.Columns("TipoPago").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvRecibos.Columns("TipoPago").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        dgvRecibos.Columns("TotalPagado").HeaderText = "VALOR"
        dgvRecibos.Columns("TotalPagado").FillWeight = 13
        dgvRecibos.Columns("TotalPagado").DefaultCellStyle.Format = "C0"
        dgvRecibos.Columns("TotalPagado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvRecibos.Columns("TotalPagado").DefaultCellStyle.Font = New Font("Consolas", 8, FontStyle.Bold)
        dgvRecibos.Columns("TotalPagado").DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 220)
        dgvRecibos.Columns("TotalPagado").DefaultCellStyle.ForeColor = Color.FromArgb(133, 100, 4)

        dgvRecibos.Columns("EstadoArchivo").HeaderText = "ESTADO"
        dgvRecibos.Columns("EstadoArchivo").FillWeight = 15
        dgvRecibos.Columns("EstadoArchivo").Width = 50
        dgvRecibos.Columns("EstadoArchivo").MinimumWidth = 80
        dgvRecibos.Columns("EstadoArchivo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvRecibos.Columns("EstadoArchivo").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        dgvRecibos.Columns("RutaArchivo").Visible = False

        ' Botones de acción
        Dim btnDescargarColumn As New DataGridViewButtonColumn With {
            .Name = "BtnDescargar",
            .HeaderText = "⬇️",
            .Text = "📥",
            .UseColumnTextForButtonValue = True,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            .Width = 70,
            .MinimumWidth = 40,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(52, 152, 219),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 8, FontStyle.Bold),
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            }
        }

        Dim btnRegenerarColumn As New DataGridViewButtonColumn With {
            .Name = "BtnRegenerar",
            .HeaderText = "🔄",
            .Text = "🔄",
            .UseColumnTextForButtonValue = True,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            .Width = 70,
            .MinimumWidth = 40,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(52, 152, 219),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 8, FontStyle.Bold),
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            }
        }
        dgvRecibos.Columns.Add(btnDescargarColumn)
        dgvRecibos.Columns.Add(btnRegenerarColumn)
    End Sub

    ' ============================================================================
    ' MÉTODOS DE CARGA DE DATOS
    ' ============================================================================

    Private Sub CargarTorres()
        Try
            cmbTorre.Items.Clear()
            cmbTorre.Items.Add("TODAS")

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT DISTINCT id_torre FROM Apartamentos ORDER BY id_torre"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            cmbTorre.Items.Add("Torre " & reader("id_torre").ToString())
                        End While
                    End Using
                End Using
            End Using

            If cmbTorre.Items.Count > 0 Then
                cmbTorre.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show("Error al cargar torres: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarApartamentos(numeroTorre As Integer)
        Try
            cmbApartamento.Items.Clear()
            cmbApartamento.Items.Add("TODOS")

            If numeroTorre > 0 Then
                Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                    conexion.Open()
                    Dim consulta As String = "SELECT numero_apartamento FROM Apartamentos WHERE id_torre = @torre ORDER BY numero_apartamento"
                    Using comando As New SQLiteCommand(consulta, conexion)
                        comando.Parameters.AddWithValue("@torre", numeroTorre)
                        Using reader As SQLiteDataReader = comando.ExecuteReader()
                            While reader.Read()
                                cmbApartamento.Items.Add($"T{numeroTorre}-{reader("numero_apartamento")}")
                            End While
                        End Using
                    End Using
                End Using
            End If

            If cmbApartamento.Items.Count > 0 Then
                cmbApartamento.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show("Error al cargar apartamentos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarRecibosRecientes()
        Try
            Me.Cursor = Cursors.WaitCursor
            lblResultados.Text = "🔄 Cargando recibos..."

            dgvRecibos.Rows.Clear()

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        p.id_pago,
                        p.numero_recibo,
                        p.fecha_pago,
                        a.id_torre,
                        a.numero_apartamento,
                        a.nombre_residente,
                        p.vr_pagado_administracion,
                        p.total_pagado,
                        CASE 
                            WHEN p.detalle LIKE 'PAGO EXTRA%' THEN 'PAGO EXTRA'
                            ELSE 'ADMINISTRACION'
                        END as tipo_pago
                    FROM pagos p
                    INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos
                    WHERE p.fecha_pago >= date('now', '-90 days')
                    ORDER BY p.fecha_pago DESC, p.id_pago DESC
                    LIMIT 100"

                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim fila As Integer = dgvRecibos.Rows.Add()

                            dgvRecibos.Rows(fila).Cells("IdPago").Value = reader("id_pago")
                            dgvRecibos.Rows(fila).Cells("NumeroRecibo").Value = reader("numero_recibo")
                            dgvRecibos.Rows(fila).Cells("FechaPago").Value = Convert.ToDateTime(reader("fecha_pago"))
                            dgvRecibos.Rows(fila).Cells("Apartamento").Value = $"T{reader("id_torre")}-{reader("numero_apartamento")}"
                            dgvRecibos.Rows(fila).Cells("Propietario").Value = If(IsDBNull(reader("nombre_residente")), "No registrado", reader("nombre_residente").ToString())
                            dgvRecibos.Rows(fila).Cells("TipoPago").Value = reader("tipo_pago")

                            ' Mostrar valor correcto según el tipo de pago
                            Dim tipoPago As String = reader("tipo_pago").ToString()
                            If tipoPago = "PAGO EXTRA" Then
                                dgvRecibos.Rows(fila).Cells("TotalPagado").Value = Convert.ToDecimal(reader("total_pagado"))
                            Else
                                dgvRecibos.Rows(fila).Cells("TotalPagado").Value = Convert.ToDecimal(reader("vr_pagado_administracion"))
                            End If

                            ' Verificar si existe el archivo PDF
                            Dim estadoArchivo As String = VerificarEstadoArchivoPDF(reader("numero_recibo").ToString(), reader("tipo_pago").ToString())
                            dgvRecibos.Rows(fila).Cells("EstadoArchivo").Value = estadoArchivo

                            ' Aplicar colores según estado del archivo
                            AplicarColoresPorEstado(dgvRecibos.Rows(fila), estadoArchivo, tipoPago)
                        End While
                    End Using
                End Using
            End Using

            lblResultados.Text = $"📊 Se encontraron {dgvRecibos.Rows.Count} recibos en los últimos 3 meses - ⏱️ Actualizado: {DateTime.Now:HH:mm}"
            lblResultados.BackColor = Color.FromArgb(212, 237, 218)
            lblResultados.ForeColor = Color.FromArgb(21, 87, 36)

        Catch ex As Exception
            lblResultados.Text = "⚠️ Error al cargar recibos"
            lblResultados.BackColor = Color.FromArgb(248, 215, 218)
            lblResultados.ForeColor = Color.FromArgb(114, 28, 36)
            MessageBox.Show("Error al cargar recibos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub AplicarColoresPorEstado(row As DataGridViewRow, estadoArchivo As String, tipoPago As String)
        ' Colorear según estado del archivo
        If estadoArchivo.Contains("✅") Then
            row.Cells("EstadoArchivo").Style.BackColor = Color.FromArgb(40, 167, 69)
            row.Cells("EstadoArchivo").Style.ForeColor = Color.White
        ElseIf estadoArchivo.Contains("❌") Then
            row.Cells("EstadoArchivo").Style.BackColor = Color.FromArgb(220, 53, 69)
            row.Cells("EstadoArchivo").Style.ForeColor = Color.White
        Else
            row.Cells("EstadoArchivo").Style.BackColor = Color.FromArgb(255, 193, 7)
            row.Cells("EstadoArchivo").Style.ForeColor = Color.Black
        End If

        ' Colorear según tipo de pago
        If tipoPago = "PAGO EXTRA" Then
            row.Cells("TipoPago").Style.BackColor = Color.FromArgb(142, 68, 173)
            row.Cells("TipoPago").Style.ForeColor = Color.White
            row.Cells("TipoPago").Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        Else
            row.Cells("TipoPago").Style.BackColor = Color.FromArgb(52, 152, 219)
            row.Cells("TipoPago").Style.ForeColor = Color.White
            row.Cells("TipoPago").Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        End If
    End Sub

    ' ============================================================================
    ' MÉTODOS DE VERIFICACIÓN Y DESCARGA
    ' ============================================================================

    Private Function VerificarEstadoArchivoPDF(numeroRecibo As String, tipoPago As String) As String
        Try
            ' Buscar en las rutas típicas de recibos
            Dim rutasRecibos As New List(Of String)

            ' Ruta principal de recibos
            Dim rutaPrincipal As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM", "Recibos")
            rutasRecibos.Add(rutaPrincipal)

            ' Ruta para pagos extra
            If tipoPago = "PAGO EXTRA" Then
                rutasRecibos.Add(Path.Combine(rutaPrincipal, "PagosExtra"))
            End If

            For Each rutaBase In rutasRecibos
                If Directory.Exists(rutaBase) Then
                    Dim archivos As String() = Directory.GetFiles(rutaBase, $"*{numeroRecibo}*.pdf", SearchOption.AllDirectories)
                    If archivos.Length > 0 Then
                        Return "✅ Disponible"
                    End If
                End If
            Next

            Return "❌ No encontrado"

        Catch ex As Exception
            Return "⚠️ Error"
        End Try
    End Function

    Private Function BuscarRutaArchivoPDF(numeroRecibo As String, tipoPago As String) As String
        Try
            Dim rutasRecibos As New List(Of String)

            Dim rutaPrincipal As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM", "Recibos")
            rutasRecibos.Add(rutaPrincipal)

            If tipoPago = "PAGO EXTRA" Then
                rutasRecibos.Add(Path.Combine(rutaPrincipal, "PagosExtra"))
            End If

            For Each rutaBase In rutasRecibos
                If Directory.Exists(rutaBase) Then
                    Dim archivos As String() = Directory.GetFiles(rutaBase, $"*{numeroRecibo}*.pdf", SearchOption.AllDirectories)
                    If archivos.Length > 0 Then
                        Return archivos(0)
                    End If
                End If
            Next

            Return String.Empty

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ' ============================================================================
    ' EVENTOS DEL FORMULARIO
    ' ============================================================================

    Private Sub cmbTorre_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbTorre.SelectedItem IsNot Nothing Then
            Dim seleccion As String = cmbTorre.SelectedItem.ToString()
            If seleccion = "TODAS" Then
                CargarApartamentos(0)
            Else
                Dim numeroTorre As Integer = Convert.ToInt32(seleccion.Replace("Torre ", ""))
                CargarApartamentos(numeroTorre)
            End If
        End If
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            lblResultados.Text = "🔄 Buscando recibos..."

            dgvRecibos.Rows.Clear()

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        p.id_pago,
                        p.numero_recibo,
                        p.fecha_pago,
                        a.id_torre,
                        a.numero_apartamento,
                        a.nombre_residente,
                        p.vr_pagado_administracion,
                        p.total_pagado,
                        CASE 
                            WHEN p.detalle LIKE 'PAGO EXTRA%' THEN 'PAGO EXTRA'
                            ELSE 'ADMINISTRACION'
                        END as tipo_pago
                    FROM pagos p
                    INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos
                    WHERE 1=1"

                Dim parametros As New List(Of SQLiteParameter)

                ' Filtro por torre
                If cmbTorre.SelectedItem IsNot Nothing AndAlso cmbTorre.SelectedItem.ToString() <> "TODAS" Then
                    Dim numeroTorre As Integer = Convert.ToInt32(cmbTorre.SelectedItem.ToString().Replace("Torre ", ""))
                    consulta += " AND a.id_torre = @torre"
                    parametros.Add(New SQLiteParameter("@torre", numeroTorre))
                End If

                ' Filtro por apartamento
                If cmbApartamento.SelectedItem IsNot Nothing AndAlso cmbApartamento.SelectedItem.ToString() <> "TODOS" Then
                    Dim apartamento As String = cmbApartamento.SelectedItem.ToString()
                    Dim numeroApartamento As String = apartamento.Split("-"c)(1)
                    consulta += " AND a.numero_apartamento = @apartamento"
                    parametros.Add(New SQLiteParameter("@apartamento", numeroApartamento))
                End If

                ' Filtro por fechas
                consulta += " AND p.fecha_pago BETWEEN @fechaInicio AND @fechaFin"
                parametros.Add(New SQLiteParameter("@fechaInicio", dtpFechaInicio.Value.ToString("yyyy-MM-dd")))
                parametros.Add(New SQLiteParameter("@fechaFin", dtpFechaFin.Value.ToString("yyyy-MM-dd")))

                ' Filtro por número de recibo
                If Not String.IsNullOrEmpty(txtNumeroRecibo.Text.Trim()) Then
                    consulta += " AND p.numero_recibo LIKE @numeroRecibo"
                    parametros.Add(New SQLiteParameter("@numeroRecibo", "%" & txtNumeroRecibo.Text.Trim() & "%"))
                End If

                consulta += " ORDER BY p.fecha_pago DESC, p.id_pago DESC LIMIT 500"

                Using comando As New SQLiteCommand(consulta, conexion)
                    For Each param In parametros
                        comando.Parameters.Add(param)
                    Next

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim fila As Integer = dgvRecibos.Rows.Add()

                            dgvRecibos.Rows(fila).Cells("IdPago").Value = reader("id_pago")
                            dgvRecibos.Rows(fila).Cells("NumeroRecibo").Value = reader("numero_recibo")
                            dgvRecibos.Rows(fila).Cells("FechaPago").Value = Convert.ToDateTime(reader("fecha_pago"))
                            dgvRecibos.Rows(fila).Cells("Apartamento").Value = $"T{reader("id_torre")}-{reader("numero_apartamento")}"
                            dgvRecibos.Rows(fila).Cells("Propietario").Value = If(IsDBNull(reader("nombre_residente")), "No registrado", reader("nombre_residente").ToString())
                            dgvRecibos.Rows(fila).Cells("TipoPago").Value = reader("tipo_pago")

                            Dim tipoPago As String = reader("tipo_pago").ToString()
                            If tipoPago = "PAGO EXTRA" Then
                                dgvRecibos.Rows(fila).Cells("TotalPagado").Value = Convert.ToDecimal(reader("total_pagado"))
                            Else
                                dgvRecibos.Rows(fila).Cells("TotalPagado").Value = Convert.ToDecimal(reader("vr_pagado_administracion"))
                            End If

                            Dim estadoArchivo As String = VerificarEstadoArchivoPDF(reader("numero_recibo").ToString(), reader("tipo_pago").ToString())
                            dgvRecibos.Rows(fila).Cells("EstadoArchivo").Value = estadoArchivo

                            AplicarColoresPorEstado(dgvRecibos.Rows(fila), estadoArchivo, tipoPago)
                        End While
                    End Using
                End Using
            End Using

            lblResultados.Text = $"🔍 Búsqueda completada: {dgvRecibos.Rows.Count} recibos encontrados - ⏱️ {DateTime.Now:HH:mm}"
            lblResultados.BackColor = Color.FromArgb(212, 237, 218)
            lblResultados.ForeColor = Color.FromArgb(21, 87, 36)

        Catch ex As Exception
            lblResultados.Text = "⚠️ Error en la búsqueda"
            lblResultados.BackColor = Color.FromArgb(248, 215, 218)
            lblResultados.ForeColor = Color.FromArgb(114, 28, 36)
            MessageBox.Show("Error en la búsqueda: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Try
            cmbTorre.SelectedIndex = 0
            cmbApartamento.SelectedIndex = 0
            dtpFechaInicio.Value = DateTime.Now.AddMonths(-3)
            dtpFechaFin.Value = DateTime.Now
            txtNumeroRecibo.Text = ""
            CargarRecibosRecientes()
        Catch ex As Exception
            MessageBox.Show("Error al limpiar filtros: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvRecibos_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim columnName As String = dgvRecibos.Columns(e.ColumnIndex).Name
            Dim row As DataGridViewRow = dgvRecibos.Rows(e.RowIndex)

            If columnName = "BtnDescargar" Then
                DescargarReciboIndividual(row)
            ElseIf columnName = "BtnRegenerar" Then
                RegenerarPDFIndividual(row)
            End If
        End If
    End Sub

    Private Sub dgvRecibos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgvRecibos.Rows(e.RowIndex)
            DescargarReciboIndividual(row)
        End If
    End Sub

    Private Sub BtnVolver_Click(sender As Object, e As EventArgs) Handles btnVolver.Click
        Try
            Me.Close()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    ' ============================================================================
    ' MÉTODOS DE DESCARGA Y REGENERACIÓN
    ' ============================================================================

    Private Sub DescargarReciboIndividual(row As DataGridViewRow)
        Try
            Dim numeroRecibo As String = row.Cells("NumeroRecibo").Value.ToString()
            Dim tipoPago As String = row.Cells("TipoPago").Value.ToString()
            Dim apartamento As String = row.Cells("Apartamento").Value.ToString()

            ' Buscar archivo existente
            Dim rutaArchivo As String = BuscarRutaArchivoPDF(numeroRecibo, tipoPago)

            If String.IsNullOrEmpty(rutaArchivo) Then
                Dim resultado As DialogResult = MessageBox.Show(
                    $"El archivo PDF del recibo {numeroRecibo} no se encontró.{vbCrLf}{vbCrLf}" &
                    "¿Desea regenerar el PDF?",
                    "Archivo no encontrado",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If resultado = DialogResult.Yes Then
                    RegenerarPDFIndividual(row)
                End If
                Return
            End If

            ' Mostrar diálogo para seleccionar ubicación de descarga
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "Archivos PDF|*.pdf"
                saveDialog.Title = "Guardar Recibo"
                saveDialog.FileName = $"Recibo_{numeroRecibo}_{apartamento}_{DateTime.Now:yyyyMMdd}.pdf"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    File.Copy(rutaArchivo, saveDialog.FileName, True)

                    Dim resultadoAbrir As DialogResult = MessageBox.Show(
                        $"✅ Recibo descargado exitosamente en:{vbCrLf}{saveDialog.FileName}{vbCrLf}{vbCrLf}" &
                        "¿Desea abrir el archivo ahora?",
                        "Descarga Exitosa",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information)

                    If resultadoAbrir = DialogResult.Yes Then
                        Process.Start(New ProcessStartInfo(saveDialog.FileName) With {.UseShellExecute = True})
                    End If
                End If
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al descargar recibo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RegenerarPDFIndividual(row As DataGridViewRow)
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim idPago As Integer = Convert.ToInt32(row.Cells("IdPago").Value)
            Dim numeroRecibo As String = row.Cells("NumeroRecibo").Value.ToString()
            Dim tipoPago As String = row.Cells("TipoPago").Value.ToString()

            ' Obtener datos del pago desde la base de datos
            Dim pago As PagoModel = Nothing
            If tipoPago = "PAGO EXTRA" Then
                pago = PagosExtraDAL.ObtenerPagoExtraPorNumeroRecibo(numeroRecibo)
            Else
                pago = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
            End If

            If pago Is Nothing Then
                MessageBox.Show("No se encontraron los datos del pago en la base de datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Obtener datos del apartamento
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(pago.IdApartamento)
            If apartamento Is Nothing Then
                MessageBox.Show("No se encontraron los datos del apartamento.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Regenerar PDF según el tipo
            Dim rutaPDFGenerado As String = ""
            If tipoPago = "PAGO EXTRA" Then
                rutaPDFGenerado = ReciboPDFExtra.GenerarReciboPagoExtra(pago, apartamento)
            Else
                rutaPDFGenerado = ReciboPDF.GenerarReciboDePagoSeguro(pago, apartamento)
            End If

            If Not String.IsNullOrEmpty(rutaPDFGenerado) AndAlso File.Exists(rutaPDFGenerado) Then
                ' Actualizar estado en la tabla
                row.Cells("EstadoArchivo").Value = "✅ Regenerado"
                AplicarColoresPorEstado(row, "✅ Regenerado", tipoPago)

                Dim resultado As DialogResult = MessageBox.Show(
                    $"✅ PDF regenerado exitosamente.{vbCrLf}{vbCrLf}" &
                    $"📁 Ubicación: {rutaPDFGenerado}{vbCrLf}{vbCrLf}" &
                    "¿Desea abrir el archivo ahora?",
                    "PDF Regenerado",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information)

                If resultado = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(rutaPDFGenerado) With {.UseShellExecute = True})
                End If
            Else
                MessageBox.Show("Error al regenerar el PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al regenerar PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnDescargarSeleccionado_Click(sender As Object, e As EventArgs) Handles btnDescargarSeleccionado.Click
        Try
            If dgvRecibos.SelectedRows.Count = 0 Then
                MessageBox.Show("Seleccione al menos un recibo para descargar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Seleccionar carpeta de destino
            Using folderDialog As New FolderBrowserDialog()
                folderDialog.Description = "Seleccione la carpeta donde guardar los recibos"
                folderDialog.ShowNewFolderButton = True

                If folderDialog.ShowDialog() = DialogResult.OK Then
                    ProcesarDescargaMasiva(dgvRecibos.SelectedRows.Cast(Of DataGridViewRow).ToList(), folderDialog.SelectedPath)
                End If
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al descargar recibos seleccionados: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDescargarTodos_Click(sender As Object, e As EventArgs) Handles btnDescargarTodos.Click
        Try
            If dgvRecibos.Rows.Count = 0 Then
                MessageBox.Show("No hay recibos para descargar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim resultado As DialogResult = MessageBox.Show(
                $"¿Confirma la descarga de {dgvRecibos.Rows.Count} recibos?{vbCrLf}{vbCrLf}" &
                "Este proceso puede tomar varios minutos.",
                "Confirmar Descarga Masiva",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                Using folderDialog As New FolderBrowserDialog()
                    folderDialog.Description = "Seleccione la carpeta donde guardar todos los recibos"
                    folderDialog.ShowNewFolderButton = True

                    If folderDialog.ShowDialog() = DialogResult.OK Then
                        ProcesarDescargaMasiva(dgvRecibos.Rows.Cast(Of DataGridViewRow).ToList(), folderDialog.SelectedPath)
                    End If
                End Using
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al descargar todos los recibos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRegenerarPDF_Click(sender As Object, e As EventArgs) Handles btnRegenerarPDF.Click
        Try
            If dgvRecibos.SelectedRows.Count = 0 Then
                MessageBox.Show("Seleccione al menos un recibo para regenerar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim resultado As DialogResult = MessageBox.Show(
                $"¿Confirma la regeneración de {dgvRecibos.SelectedRows.Count} PDF(s)?{vbCrLf}{vbCrLf}" &
                "Los archivos existentes serán reemplazados.",
                "Confirmar Regeneración",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                For Each row As DataGridViewRow In dgvRecibos.SelectedRows
                    RegenerarPDFIndividual(row)
                    System.Threading.Thread.Sleep(500)
                Next
                MessageBox.Show("Regeneración completada.", "Proceso Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al regenerar PDFs: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ProcesarDescargaMasiva(filas As List(Of DataGridViewRow), carpetaDestino As String)
        Dim progreso As New ProgressBar With {
            .Minimum = 0,
            .Maximum = filas.Count,
            .Value = 0,
            .Dock = DockStyle.Bottom,
            .Height = 25
        }

        Dim lblProgreso As New Label With {
            .Text = "Preparando descarga masiva...",
            .Dock = DockStyle.Bottom,
            .Height = 25,
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.LightBlue
        }

        Me.Controls.Add(progreso)
        Me.Controls.Add(lblProgreso)
        Me.Cursor = Cursors.WaitCursor

        Dim exitosos As Integer = 0
        Dim fallidos As Integer = 0
        Dim regenerados As Integer = 0
        Dim errores As New List(Of String)

        Try
            For i As Integer = 0 To filas.Count - 1
                Dim row As DataGridViewRow = filas(i)

                Try
                    lblProgreso.Text = $"Procesando recibo {i + 1} de {filas.Count}..."
                    progreso.Value = i + 1
                    Application.DoEvents()

                    Dim numeroRecibo As String = row.Cells("NumeroRecibo").Value.ToString()
                    Dim tipoPago As String = row.Cells("TipoPago").Value.ToString()
                    Dim apartamento As String = row.Cells("Apartamento").Value.ToString()

                    ' Buscar archivo existente
                    Dim rutaArchivo As String = BuscarRutaArchivoPDF(numeroRecibo, tipoPago)

                    ' Si no existe, intentar regenerar
                    If String.IsNullOrEmpty(rutaArchivo) Then
                        Dim idPago As Integer = Convert.ToInt32(row.Cells("IdPago").Value)
                        rutaArchivo = RegenerarPDFParaDescarga(idPago, numeroRecibo, tipoPago)
                        If Not String.IsNullOrEmpty(rutaArchivo) Then
                            regenerados += 1
                        End If
                    End If

                    If Not String.IsNullOrEmpty(rutaArchivo) AndAlso File.Exists(rutaArchivo) Then
                        ' Copiar archivo a destino
                        Dim nombreDestino As String = $"Recibo_{numeroRecibo}_{apartamento}_{DateTime.Now:yyyyMMdd}.pdf"
                        Dim rutaDestino As String = Path.Combine(carpetaDestino, nombreDestino)

                        File.Copy(rutaArchivo, rutaDestino, True)
                        exitosos += 1
                    Else
                        fallidos += 1
                        errores.Add($"Recibo {numeroRecibo}: No se pudo generar o encontrar el PDF")
                    End If

                    System.Threading.Thread.Sleep(100)

                Catch ex As Exception
                    fallidos += 1
                    errores.Add($"Recibo {i + 1}: {ex.Message}")
                End Try
            Next

            ' Mostrar resultado
            Dim mensaje As String = $"✅ Descarga masiva completada:{vbCrLf}{vbCrLf}" &
                              $"📁 Carpeta: {carpetaDestino}{vbCrLf}" &
                              $"✅ Exitosos: {exitosos}{vbCrLf}" &
                              $"🔄 Regenerados: {regenerados}{vbCrLf}" &
                              $"❌ Fallidos: {fallidos}"

            If errores.Count > 0 AndAlso errores.Count <= 5 Then
                mensaje &= $"{vbCrLf}{vbCrLf}❌ Errores:{vbCrLf}{String.Join(vbCrLf, errores)}"
            ElseIf errores.Count > 5 Then
                mensaje &= $"{vbCrLf}{vbCrLf}❌ Se produjeron {errores.Count} errores."
            End If

            MessageBox.Show(mensaje, "Resultado Descarga Masiva", MessageBoxButtons.OK,
                          If(fallidos = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))

            ' Preguntar si desea abrir la carpeta
            If exitosos > 0 Then
                Dim abrirCarpeta As DialogResult = MessageBox.Show(
                    "¿Desea abrir la carpeta con los recibos descargados?",
                    "Abrir Carpeta",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If abrirCarpeta = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(carpetaDestino) With {.UseShellExecute = True})
                End If
            End If

        Finally
            Me.Controls.Remove(progreso)
            Me.Controls.Remove(lblProgreso)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function RegenerarPDFParaDescarga(idPago As Integer, numeroRecibo As String, tipoPago As String) As String
        Try
            Dim pago As PagoModel = Nothing
            If tipoPago = "PAGO EXTRA" Then
                pago = PagosExtraDAL.ObtenerPagoExtraPorNumeroRecibo(numeroRecibo)
            Else
                pago = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
            End If

            If pago Is Nothing Then Return String.Empty

            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(pago.IdApartamento)
            If apartamento Is Nothing Then Return String.Empty

            If tipoPago = "PAGO EXTRA" Then
                Return ReciboPDFExtra.GenerarReciboPagoExtra(pago, apartamento)
            Else
                Return ReciboPDF.GenerarReciboDePagoSeguro(pago, apartamento)
            End If

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ' ============================================================================
    ' MÉTODOS AUXILIARES
    ' ============================================================================

    Private Sub FormDescargarRecibos_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Try
            ' Reposicionar el botón VOLVER centrado en el panel inferior (IGUAL QUE FORMHISTORIAL)
            If btnVolver IsNot Nothing Then
                btnVolver.Location = New Point((Me.Width - btnVolver.Width) \ 2, 10)
            End If

            ' Redimensionar panel de estadísticas para usar todo el ancho disponible
            If lblResultados IsNot Nothing Then
                Dim anchoDisponible As Integer = Me.Width - 40
                If anchoDisponible > 400 Then
                    lblResultados.Width = anchoDisponible
                End If
            End If

        Catch ex As Exception
            ' Ignorar errores de redimensionamiento
        End Try
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            ' Limpieza de recursos si es necesario
        Catch ex As Exception
            ' Ignorar errores de limpieza
        End Try

        MyBase.OnFormClosed(e)
    End Sub

End Class