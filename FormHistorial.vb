' ============================================================================
' FORMULARIO DE HISTORIAL COMPLETAMENTE CORREGIDO - TODOS LOS PROBLEMAS SOLUCIONADOS
' ============================================================================

Imports System.Data.SQLite
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Linq
Imports System.Diagnostics
Imports System.Drawing.Printing

Public Class FormHistorial
    Private dtPagos As DataTable
    Private apartamentoSeleccionado As Apartamento

    ' Controles del formulario
    Private WithEvents dgvPagos As DataGridView
    Private WithEvents cmbApartamento As ComboBox
    Private WithEvents dtpDesde As DateTimePicker
    Private WithEvents dtpHasta As DateTimePicker
    Private WithEvents btnBuscarPagos As Button
    Private WithEvents btnExportarPagos As Button
    Private WithEvents btnExportarPDF As Button
    Private lblTotalesPagos As Label
    Private WithEvents btnVolver As Button

    ' Variables para exportación PDF
    Private WithEvents printDocument As PrintDocument
    Private datosPagoParaImprimir As List(Of Dictionary(Of String, Object))
    Private apartamentoSeleccionadoPDF As String
    Private fechaImpresionPDF As String
    Private paginaActual As Integer = 1
    Private totalPaginas As Integer = 1
    Private filasPorPagina As Integer = 30

    Private Sub FormHistorial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ConfigurarFormulario()
            CargarDatos()
        Catch ex As Exception
            MessageBox.Show("Error al cargar el formulario: " & ex.Message, "Error",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configuración de ventana completa
            Me.Text = "Historial - COOPDIASAM"
            Me.WindowState = FormWindowState.Maximized
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MinimumSize = New Size(1400, 700)
            Me.MaximizeBox = True
            Me.MinimizeBox = True
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' Panel superior con mejor diseño
            Dim panelSuperior As New Panel()
            panelSuperior.Dock = DockStyle.Top
            panelSuperior.Height = 80
            panelSuperior.BackColor = Color.FromArgb(46, 132, 188)

            Dim lblTitulo As New Label()
            lblTitulo.Text = "📊 HISTORIAL"
            lblTitulo.Font = New Font("Segoe UI", 20, FontStyle.Bold)
            lblTitulo.ForeColor = Color.White
            lblTitulo.AutoSize = True
            lblTitulo.Location = New Point(30, 25)
            panelSuperior.Controls.Add(lblTitulo)

            ' Panel inferior para el botón VOLVER - CENTRADO ABAJO
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
            btnVolver.Anchor = AnchorStyles.None  ' CENTRADO

            ' EVENTO PARA MANTENERLO CENTRADO
            AddHandler Me.Resize, Sub() btnVolver.Location = New Point((Me.Width - btnVolver.Width) \ 2, 10)
            btnVolver.Location = New Point((Me.Width - btnVolver.Width) \ 2, 10)
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
        ' Panel de filtros con mejor organización y altura aumentada
        Dim panelFiltros As New Panel()
        panelFiltros.Dock = DockStyle.Top
        panelFiltros.Height = 150  ' AUMENTADO PARA ACOMODAR TODO
        panelFiltros.BackColor = Color.FromArgb(248, 249, 250)
        panelFiltros.Padding = New Padding(30)

        ' === PRIMERA FILA DE CONTROLES ===
        ' Apartamento
        Dim lblApartamento As New Label()
        lblApartamento.Text = "Apartamento:"
        lblApartamento.Location = New Point(30, 25)
        lblApartamento.Size = New Size(120, 25)
        lblApartamento.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblApartamento.ForeColor = Color.FromArgb(52, 73, 94)
        lblApartamento.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblApartamento)

        cmbApartamento = New ComboBox()
        cmbApartamento.Location = New Point(170, 23)
        cmbApartamento.Size = New Size(250, 30)
        cmbApartamento.DropDownStyle = ComboBoxStyle.DropDownList
        cmbApartamento.Font = New Font("Segoe UI", 11)
        cmbApartamento.BackColor = Color.White
        cmbApartamento.FlatStyle = FlatStyle.Flat
        panelFiltros.Controls.Add(cmbApartamento)

        ' DESDE - CORREGIDO Y VISIBLE
        Dim lblDesde As New Label()
        lblDesde.Text = "📅 Desde:"
        lblDesde.Location = New Point(430, 25)
        lblDesde.Size = New Size(110, 25)
        lblDesde.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblDesde.ForeColor = Color.FromArgb(52, 73, 94)
        lblDesde.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblDesde)

        dtpDesde = New DateTimePicker()
        dtpDesde.Location = New Point(540, 23)
        dtpDesde.Size = New Size(130, 30)
        dtpDesde.Format = DateTimePickerFormat.Short
        dtpDesde.Font = New Font("Segoe UI", 11)
        dtpDesde.Value = DateTime.Now.AddMonths(-6)
        panelFiltros.Controls.Add(dtpDesde)

        ' HASTA - CORREGIDO Y VISIBLE
        Dim lblHasta As New Label()
        lblHasta.Text = "📅 Hasta:"
        lblHasta.Location = New Point(680, 25)
        lblHasta.Size = New Size(110, 25)
        lblHasta.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblHasta.ForeColor = Color.FromArgb(52, 73, 94)
        lblHasta.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblHasta)

        dtpHasta = New DateTimePicker()
        dtpHasta.Location = New Point(790, 23)
        dtpHasta.Size = New Size(130, 30)
        dtpHasta.Format = DateTimePickerFormat.Short
        dtpHasta.Font = New Font("Segoe UI", 11)
        panelFiltros.Controls.Add(dtpHasta)

        ' === SEGUNDA FILA - BOTONES ===
        btnExportarPDF = New Button()
        btnExportarPDF.Text = "EXPORTAR PDF"
        btnExportarPDF.Location = New Point(30, 70) ' <-- Cambiado
        btnExportarPDF.Size = New Size(150, 45)
        btnExportarPDF.BackColor = Color.FromArgb(231, 76, 60)
        btnExportarPDF.ForeColor = Color.White
        btnExportarPDF.FlatStyle = FlatStyle.Flat
        btnExportarPDF.Cursor = Cursors.Hand
        btnExportarPDF.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnExportarPDF.FlatAppearance.BorderSize = 0
        btnExportarPDF.FlatAppearance.MouseOverBackColor = Color.FromArgb(192, 57, 43)
        panelFiltros.Controls.Add(btnExportarPDF)

        btnBuscarPagos = New Button()
        btnBuscarPagos.Text = "🔍 BUSCAR"
        btnBuscarPagos.Location = New Point(190, 70) ' <-- Cambiado
        btnBuscarPagos.Size = New Size(130, 45)
        btnBuscarPagos.BackColor = Color.FromArgb(52, 152, 219)
        btnBuscarPagos.ForeColor = Color.White
        btnBuscarPagos.FlatStyle = FlatStyle.Flat
        btnBuscarPagos.Cursor = Cursors.Hand
        btnBuscarPagos.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnBuscarPagos.FlatAppearance.BorderSize = 0
        btnBuscarPagos.FlatAppearance.MouseOverBackColor = Color.FromArgb(41, 128, 185)
        panelFiltros.Controls.Add(btnBuscarPagos)


        ' === PANEL DE ESTADÍSTICAS VERTICAL, GRANDE Y SIN ESPACIO EN BLANCO ===
        lblTotalesPagos = New Label()
        lblTotalesPagos.Location = New Point(340, 70)
        lblTotalesPagos.Size = New Size(300, 45)  ' MÁS ANCHO PARA EVITAR ESPACIO EN BLANCO
        lblTotalesPagos.Font = New Font("Segoe UI", 10, FontStyle.Bold)  ' FUENTE MÁS GRANDE
        lblTotalesPagos.ForeColor = Color.FromArgb(44, 62, 80)
        lblTotalesPagos.Text = "📊 Seleccione filtros y presione BUSCAR para ver estadísticas"
        lblTotalesPagos.BackColor = Color.FromArgb(236, 240, 241)
        lblTotalesPagos.BorderStyle = BorderStyle.FixedSingle
        lblTotalesPagos.TextAlign = ContentAlignment.TopLeft
        lblTotalesPagos.Padding = New Padding(15, 8, 15, 8)
        lblTotalesPagos.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right  ' ANCLAJE PARA OCUPAR TODO EL ANCHO
        panelFiltros.Controls.Add(lblTotalesPagos)


        ' DataGridView principal
        dgvPagos = New DataGridView()
        dgvPagos.Dock = DockStyle.Fill
        dgvPagos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill  ' LLENAR TODO EL ANCHO
        dgvPagos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvPagos.ReadOnly = True
        dgvPagos.AllowUserToAddRows = False
        dgvPagos.AllowUserToDeleteRows = False
        dgvPagos.BackgroundColor = Color.White
        dgvPagos.BorderStyle = BorderStyle.None
        dgvPagos.RowHeadersVisible = False
        dgvPagos.Font = New Font("Segoe UI", 10.0F)
        dgvPagos.ScrollBars = ScrollBars.Both
        dgvPagos.AllowUserToResizeColumns = True
        dgvPagos.MultiSelect = False
        ConfigurarDataGridView(dgvPagos)


        ' Agregar al formulario
        Me.Controls.Add(dgvPagos)
        Me.Controls.Add(panelFiltros)
    End Sub

    Private Sub ConfigurarDataGridView(dgv As DataGridView)
        ' Configuración visual moderna y profesional - CABECERAS CORREGIDAS
        dgv.EnableHeadersVisualStyles = False

        ' Encabezados con diseño moderno - LETRAS MÁS PEQUEÑAS
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(44, 62, 80)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)  ' FUENTE MÁS PEQUEÑA
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(3, 8, 3, 8)  ' PADDING REDUCIDO
        dgv.ColumnHeadersHeight = 35  ' ALTURA REDUCIDA ENTRE Y LA COLUMNAS DE INFORMACION

        ' Estilo de celdas mejorado - LETRAS MÁS PEQUEÑAS
        dgv.DefaultCellStyle.Font = New Font("Segoe UI", 8.5F)  ' FUENTE MÁS PEQUEÑA
        dgv.DefaultCellStyle.Padding = New Padding(6, 4, 6, 4)  ' PADDING REDUCIDO
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
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single  ' BORDES VISIBLES
        dgv.MultiSelect = False
        dgv.AutoGenerateColumns = True
    End Sub

    Private Sub CargarDatos()
        Try
            DiagnosticarFechasProblematicas()
            CargarApartamentos()

            If cmbApartamento.Items.Count > 0 Then
                cmbApartamento.SelectedIndex = 0
                CargarHistorialPagos(0)
            End If

        Catch ex As Exception
            MessageBox.Show("Error al cargar datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarApartamentos()
        Try
            cmbApartamento.Items.Clear()
            cmbApartamento.Items.Add("TODOS LOS APARTAMENTOS")

            ' Cargar directamente desde la base de datos
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT id_torre, numero_apartamento FROM Apartamentos ORDER BY id_torre, numero_apartamento"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim torre As String = reader("id_torre").ToString()
                            Dim numero As String = reader("numero_apartamento").ToString()
                            cmbApartamento.Items.Add("Torre " & torre & " - " & numero)
                        End While
                    End Using
                End Using
            End Using

            If cmbApartamento.Items.Count > 0 Then
                cmbApartamento.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show("Error al cargar apartamentos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Al menos tener la opción "TODOS"
            If cmbApartamento.Items.Count = 0 Then
                cmbApartamento.Items.Add("TODOS LOS APARTAMENTOS")
                cmbApartamento.SelectedIndex = 0
            End If
        End Try
    End Sub

    Private Sub CargarHistorialPagos(Optional idApartamento As Integer = 0)
        Try
            Me.Cursor = Cursors.WaitCursor
            lblTotalesPagos.Text = "🔄 Cargando datos..."

            dtPagos = New DataTable()
            dtPagos.Columns.Add("Fecha", GetType(DateTime))
            dtPagos.Columns.Add("NumeroRecibo", GetType(String))
            dtPagos.Columns.Add("Apartamento", GetType(String))
            dtPagos.Columns.Add("Residente", GetType(String))
            dtPagos.Columns.Add("SaldoAnterior", GetType(Decimal))
            dtPagos.Columns.Add("PagoAdministracion", GetType(Decimal))
            dtPagos.Columns.Add("PagoIntereses", GetType(Decimal))
            dtPagos.Columns.Add("TotalPagado", GetType(Decimal))
            dtPagos.Columns.Add("SaldoActual", GetType(Decimal))
            dtPagos.Columns.Add("EstadoPago", GetType(String))
            dtPagos.Columns.Add("Observaciones", GetType(String))

            Dim totalPagado As Decimal = 0
            Dim totalAdministracion As Decimal = 0
            Dim totalIntereses As Decimal = 0
            Dim totalRegistros As Integer = 0

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        CAST(p.fecha_pago AS TEXT) as fecha_pago_texto,
                        p.numero_recibo,
                        p.saldo_anterior,
                        p.vr_pagado_administracion,
                        p.vr_pagado_intereses,
                        p.total_pagado,
                        p.saldo_actual,
                        p.estado_pago,
                        p.observacion,
                        a.id_torre,
                        a.numero_apartamento,
                        a.nombre_residente
                    FROM pagos p
                    INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos
                    WHERE p.fecha_pago IS NOT NULL"

                If idApartamento > 0 Then
                    consulta &= " AND p.id_apartamentos = @idApartamento"
                End If

                consulta &= " ORDER BY p.fecha_pago DESC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    If idApartamento > 0 Then
                        comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    End If

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim fechaPago As DateTime = DateTime.MinValue
                            Dim fechaTexto As String = ""
                            Dim fechaParseada As Boolean = False

                            Try
                                If Not IsDBNull(reader("fecha_pago_texto")) Then
                                    fechaTexto = reader("fecha_pago_texto").ToString().Trim()
                                End If
                            Catch ex As Exception
                                Console.WriteLine("Error al leer fecha_pago_texto: " & ex.Message)
                                Continue While
                            End Try

                            If Not String.IsNullOrEmpty(fechaTexto) Then
                                fechaTexto = fechaTexto.Replace(" ", "").Replace(vbTab, "").Replace(vbCrLf, "").Replace(vbLf, "")

                                If Not String.IsNullOrEmpty(fechaTexto) Then
                                    If DateTime.TryParseExact(fechaTexto, "d/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaPago) Then
                                        fechaParseada = True
                                    ElseIf DateTime.TryParseExact(fechaTexto, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaPago) Then
                                        fechaParseada = True
                                    ElseIf DateTime.TryParseExact(fechaTexto, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaPago) Then
                                        fechaParseada = True
                                    ElseIf DateTime.TryParseExact(fechaTexto, "dd/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaPago) Then
                                        fechaParseada = True
                                    ElseIf DateTime.TryParseExact(fechaTexto, "M/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaPago) Then
                                        fechaParseada = True
                                    ElseIf DateTime.TryParseExact(fechaTexto, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, fechaPago) Then
                                        fechaParseada = True
                                    Else
                                        If DateTime.TryParse(fechaTexto, fechaPago) Then
                                            fechaParseada = True
                                        Else
                                            Console.WriteLine("Fecha no reconocida: '" & fechaTexto & "'")
                                            Continue While
                                        End If
                                    End If
                                Else
                                    Continue While
                                End If
                            Else
                                Continue While
                            End If

                            If fechaParseada AndAlso fechaPago >= dtpDesde.Value.Date AndAlso fechaPago <= dtpHasta.Value.Date.AddDays(1) Then
                                Dim row As DataRow = dtPagos.NewRow()

                                row("Fecha") = fechaPago
                                row("NumeroRecibo") = If(IsDBNull(reader("numero_recibo")), "", reader("numero_recibo").ToString())
                                row("Apartamento") = "Torre " & reader("id_torre").ToString() & " - " & reader("numero_apartamento").ToString()
                                row("Residente") = If(IsDBNull(reader("nombre_residente")), "Sin registrar", reader("nombre_residente").ToString())
                                row("SaldoAnterior") = Convert.ToDecimal(If(IsDBNull(reader("saldo_anterior")), 0, reader("saldo_anterior")))
                                row("PagoAdministracion") = Convert.ToDecimal(If(IsDBNull(reader("vr_pagado_administracion")), 0, reader("vr_pagado_administracion")))
                                row("PagoIntereses") = Convert.ToDecimal(If(IsDBNull(reader("vr_pagado_intereses")), 0, reader("vr_pagado_intereses")))
                                row("TotalPagado") = Convert.ToDecimal(If(IsDBNull(reader("total_pagado")), 0, reader("total_pagado")))
                                row("SaldoActual") = Convert.ToDecimal(If(IsDBNull(reader("saldo_actual")), 0, reader("saldo_actual")))
                                row("EstadoPago") = If(IsDBNull(reader("estado_pago")), "REGISTRADO", reader("estado_pago").ToString())
                                row("Observaciones") = If(IsDBNull(reader("observacion")), "", reader("observacion").ToString())

                                dtPagos.Rows.Add(row)

                                totalPagado += Convert.ToDecimal(If(IsDBNull(reader("total_pagado")), 0, reader("total_pagado")))
                                totalAdministracion += Convert.ToDecimal(If(IsDBNull(reader("vr_pagado_administracion")), 0, reader("vr_pagado_administracion")))
                                totalIntereses += Convert.ToDecimal(If(IsDBNull(reader("vr_pagado_intereses")), 0, reader("vr_pagado_intereses")))
                                totalRegistros += 1
                            End If
                        End While
                    End Using
                End Using
            End Using

            dgvPagos.DataSource = dtPagos
            FormatearColumnasPagos()
            ActualizarLabelTotales(totalRegistros, totalPagado, totalAdministracion, totalIntereses)

        Catch ex As Exception
            MessageBox.Show("Error al cargar historial de pagos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FormatearColumnasPagos()
        If dgvPagos.Columns.Count > 0 Then
            With dgvPagos
                .Columns("Fecha").HeaderText = "FECHA"
                .Columns("Fecha").DefaultCellStyle.Format = "dd/MM/yyyy"
                .Columns("Fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns("Fecha").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

                .Columns("NumeroRecibo").HeaderText = "N° RECIBO"
                .Columns("NumeroRecibo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns("NumeroRecibo").DefaultCellStyle.Font = New Font("Consolas", 8, FontStyle.Bold)

                .Columns("Apartamento").HeaderText = "APARTAMENTO"
                .Columns("Apartamento").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns("Apartamento").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

                .Columns("Residente").HeaderText = "RESIDENTE"
                .Columns("Residente").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

                .Columns("SaldoAnterior").HeaderText = "SALDO ANTERIOR"
                .Columns("SaldoAnterior").DefaultCellStyle.Format = "C0"
                .Columns("SaldoAnterior").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns("SaldoAnterior").DefaultCellStyle.Font = New Font("Consolas", 8)

                .Columns("PagoAdministracion").HeaderText = "ADMINISTRACIÓN"
                .Columns("PagoAdministracion").DefaultCellStyle.Format = "C0"
                .Columns("PagoAdministracion").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns("PagoAdministracion").DefaultCellStyle.Font = New Font("Consolas", 8, FontStyle.Bold)
                .Columns("PagoAdministracion").DefaultCellStyle.BackColor = Color.FromArgb(230, 247, 255)

                .Columns("PagoIntereses").HeaderText = "INTERESES"
                .Columns("PagoIntereses").DefaultCellStyle.Format = "C0"
                .Columns("PagoIntereses").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns("PagoIntereses").DefaultCellStyle.Font = New Font("Consolas", 8)

                .Columns("TotalPagado").HeaderText = "TOTAL PAGADO"
                .Columns("TotalPagado").DefaultCellStyle.Format = "C0"
                .Columns("TotalPagado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns("TotalPagado").DefaultCellStyle.Font = New Font("Consolas", 9, FontStyle.Bold)
                .Columns("TotalPagado").DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 220)
                .Columns("TotalPagado").DefaultCellStyle.ForeColor = Color.FromArgb(133, 100, 4)

                .Columns("SaldoActual").HeaderText = "SALDO ACTUAL"
                .Columns("SaldoActual").DefaultCellStyle.Format = "C0"
                .Columns("SaldoActual").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns("SaldoActual").DefaultCellStyle.Font = New Font("Consolas", 8)

                .Columns("EstadoPago").HeaderText = "ESTADO"
                .Columns("EstadoPago").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns("EstadoPago").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)

                .Columns("Observaciones").HeaderText = "OBSERVACIONES"
                .Columns("Observaciones").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns("Observaciones").DefaultCellStyle.WrapMode = DataGridViewTriState.True
            End With
        End If
    End Sub

    Private Sub ActualizarLabelTotales(totalRegistros As Integer, totalPagado As Decimal,
                                      totalAdministracion As Decimal, totalIntereses As Decimal)
        If totalRegistros > 0 Then
            Dim promedioMensual As Decimal = If(totalRegistros > 0, totalPagado / totalRegistros, 0)
            Dim porcentajeIntereses As Decimal = If(totalPagado > 0, (totalIntereses / totalPagado) * 100, 0)


            lblTotalesPagos.Text = String.Format(
                "📊 RESUMEN: {0} pagos registrados - 💰 Total Recaudado: {1} - 🏢 Administración: {2} - 📈 Intereses: {3} ({4:F1}%) - 📊 Promedio por pago: {5} - ⏱️ Actualizado: {8}",
                totalRegistros.ToString("N0"),
                totalPagado.ToString("C0"),
                totalAdministracion.ToString("C0"),
                totalIntereses.ToString("C0"),
                porcentajeIntereses,
                promedioMensual.ToString("C0"),
                dtpDesde.Value.ToString("dd/MM/yyyy"),
                dtpHasta.Value.ToString("dd/MM/yyyy"),
                DateTime.Now.ToString("HH:mm")
            )
            lblTotalesPagos.BackColor = Color.FromArgb(212, 237, 218)
            lblTotalesPagos.ForeColor = Color.FromArgb(21, 87, 36)
        Else
            lblTotalesPagos.Text = "⚠️ Escoger apartamento y el rango de fecha"
            lblTotalesPagos.BackColor = Color.FromArgb(248, 215, 218)
            lblTotalesPagos.ForeColor = Color.FromArgb(114, 28, 36)
            lblTotalesPagos.Font = New Font("Segoe UI", 10, FontStyle.Regular)

        End If
    End Sub

    ' ============================================================================
    ' EVENTOS DE INTERFAZ
    ' ============================================================================

    Private Sub CmbApartamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbApartamento.SelectedIndexChanged
        Try
            If cmbApartamento.SelectedIndex >= 0 Then
                If cmbApartamento.SelectedIndex = 0 Then
                    CargarHistorialPagos(0)
                Else
                    Dim textoSeleccionado As String = cmbApartamento.SelectedItem.ToString()
                    Dim partes() As String = textoSeleccionado.Split("-"c)
                    If partes.Length = 2 Then
                        Dim torreTexto As String = partes(0).Trim()
                        Dim numero As String = partes(1).Trim()

                        Dim torre As Integer = 0
                        If torreTexto.StartsWith("Torre ") Then
                            Integer.TryParse(torreTexto.Substring(6), torre)
                        End If

                        ' Buscar directamente en la base de datos
                        Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                            conexion.Open()
                            Dim consulta As String = "SELECT id_apartamentos FROM Apartamentos WHERE id_torre = @torre AND numero_apartamento = @numero"
                            Using comando As New SQLiteCommand(consulta, conexion)
                                comando.Parameters.AddWithValue("@torre", torre)
                                comando.Parameters.AddWithValue("@numero", numero)
                                Dim resultado = comando.ExecuteScalar()
                                If resultado IsNot Nothing Then
                                    CargarHistorialPagos(Convert.ToInt32(resultado))
                                Else
                                    CargarHistorialPagos(0)
                                End If
                            End Using
                        End Using
                    End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error en CmbApartamento_SelectedIndexChanged: " & ex.Message)
            CargarHistorialPagos(0)
        End Try
    End Sub

    Private Sub BtnBuscarPagos_Click(sender As Object, e As EventArgs) Handles btnBuscarPagos.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            lblTotalesPagos.Text = "🔄 Buscando registros..."

            If dtpHasta.Value < dtpDesde.Value Then
                MessageBox.Show("La fecha 'Hasta' debe ser mayor o igual a la fecha 'Desde'.",
                              "Rango de fechas inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            CmbApartamento_SelectedIndexChanged(sender, e)
        Catch ex As Exception
            MessageBox.Show("Error al buscar pagos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnExportarPagos_Click(sender As Object, e As EventArgs) Handles btnExportarPagos.Click
        Try
            If dtPagos Is Nothing OrElse dtPagos.Rows.Count = 0 Then
                MessageBox.Show("No hay datos de pagos para exportar.", "Sin datos",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Archivo CSV|*.csv"
            saveDialog.Title = "Exportar Historial de Pagos"
            saveDialog.FileName = "HistorialPagos_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                ExportarHistorialPagosCSV(saveDialog.FileName)
                MessageBox.Show("Historial de pagos exportado exitosamente.", "Éxito",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)

                If MessageBox.Show("¿Desea abrir el archivo exportado?", "Abrir archivo",
                                 MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(saveDialog.FileName) With {.UseShellExecute = True})
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error al exportar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnExportarPDF_Click(sender As Object, e As EventArgs) Handles btnExportarPDF.Click
        Try
            If dtPagos Is Nothing OrElse dtPagos.Rows.Count = 0 Then
                MessageBox.Show("No hay datos de pagos para exportar a PDF.", "Sin datos",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            PrepararDatosParaPDF()

            printDocument = New PrintDocument()
            printDocument.DefaultPageSettings.PaperSize = New PaperSize("A4", 827, 1169)
            printDocument.DefaultPageSettings.Margins = New Margins(40, 40, 40, 40)
            printDocument.DefaultPageSettings.Landscape = True

            Dim previewDialog As New PrintPreviewDialog()
            previewDialog.Document = printDocument
            previewDialog.Size = New Size(900, 700)
            previewDialog.ShowDialog()

        Catch ex As Exception
            MessageBox.Show("Error al generar PDF: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnVolver_Click(sender As Object, e As EventArgs) Handles btnVolver.Click
        Try
            Me.Close()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub


    ' ============================================================================
    ' FORMATEO DE CELDAS
    ' ============================================================================

    Private Sub DgvPagos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvPagos.CellFormatting
        If e.RowIndex >= 0 Then
            ' Estados de pago con colores modernos
            If e.ColumnIndex = dgvPagos.Columns("EstadoPago").Index Then
                If e.Value IsNot Nothing Then
                    Select Case e.Value.ToString().ToUpper()
                        Case "PAGADO", "REGISTRADO", "CONFIRMADO"
                            e.CellStyle.BackColor = Color.FromArgb(40, 167, 69)
                            e.CellStyle.ForeColor = Color.White
                            e.CellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                        Case "ANULADO", "CANCELADO"
                            e.CellStyle.BackColor = Color.FromArgb(220, 53, 69)
                            e.CellStyle.ForeColor = Color.White
                            e.CellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                        Case "PENDIENTE"
                            e.CellStyle.BackColor = Color.FromArgb(255, 193, 7)
                            e.CellStyle.ForeColor = Color.Black
                            e.CellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                        Case Else
                            e.CellStyle.BackColor = Color.FromArgb(108, 117, 125)
                            e.CellStyle.ForeColor = Color.White
                            e.CellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                    End Select
                End If
            End If

            ' Resaltar montos altos
            If e.ColumnIndex = dgvPagos.Columns("TotalPagado").Index Then
                If e.Value IsNot Nothing Then
                    Dim valor As Decimal = Convert.ToDecimal(e.Value)
                    If valor > 200000 Then
                        e.CellStyle.BackColor = Color.FromArgb(255, 235, 205)
                        e.CellStyle.ForeColor = Color.FromArgb(102, 77, 3)
                        e.CellStyle.Font = New Font("Consolas", 10, FontStyle.Bold)
                    ElseIf valor > 100000 Then
                        e.CellStyle.BackColor = Color.FromArgb(255, 243, 224)
                        e.CellStyle.ForeColor = Color.FromArgb(133, 100, 4)
                        e.CellStyle.Font = New Font("Consolas", 10, FontStyle.Bold)
                    End If
                End If
            End If

            ' Destacar saldos pendientes
            If e.ColumnIndex = dgvPagos.Columns("SaldoActual").Index Then
                If e.Value IsNot Nothing Then
                    Dim saldo As Decimal = Convert.ToDecimal(e.Value)
                    If saldo > 0 Then
                        e.CellStyle.BackColor = Color.FromArgb(255, 243, 243)
                        e.CellStyle.ForeColor = Color.FromArgb(169, 68, 66)
                        e.CellStyle.Font = New Font("Consolas", 10, FontStyle.Bold)
                    End If
                End If
            End If

            ' Destacar fechas recientes
            If e.ColumnIndex = dgvPagos.Columns("Fecha").Index Then
                If e.Value IsNot Nothing Then
                    Dim fechaPago As DateTime = Convert.ToDateTime(e.Value)
                    If fechaPago >= DateTime.Now.AddDays(-30) Then
                        e.CellStyle.BackColor = Color.FromArgb(240, 248, 255)
                        e.CellStyle.ForeColor = Color.FromArgb(52, 73, 94)
                        e.CellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                    End If
                End If
            End If
        End If
    End Sub

    ' ============================================================================
    ' MÉTODOS DE EXPORTACIÓN
    ' ============================================================================

    Private Sub ExportarHistorialPagosCSV(rutaArchivo As String)
        Using writer As New StreamWriter(rutaArchivo, False, System.Text.Encoding.UTF8)
            writer.WriteLine("sep=,")
            writer.WriteLine("Fecha,Número Recibo,Apartamento,Residente,Saldo Anterior,Pago Administración,Pago Intereses,Total Pagado,Saldo Actual,Estado,Observaciones")

            For Each row As DataRow In dtPagos.Rows
                Dim fecha As String = ""
                If row("Fecha") IsNot Nothing AndAlso Not IsDBNull(row("Fecha")) Then
                    fecha = Convert.ToDateTime(row("Fecha")).ToString("dd/MM/yyyy")
                End If

                Dim linea As String = String.Join(",", {
                    """" & fecha & """",
                    """" & ObtenerValorCeldaSeguro(row("NumeroRecibo"), "") & """",
                    """" & ObtenerValorCeldaSeguro(row("Apartamento"), "") & """",
                    """" & ObtenerValorCeldaSeguro(row("Residente"), "") & """",
                    ObtenerValorCeldaSeguro(row("SaldoAnterior"), "0"),
                    ObtenerValorCeldaSeguro(row("PagoAdministracion"), "0"),
                    ObtenerValorCeldaSeguro(row("PagoIntereses"), "0"),
                    ObtenerValorCeldaSeguro(row("TotalPagado"), "0"),
                    ObtenerValorCeldaSeguro(row("SaldoActual"), "0"),
                    """" & ObtenerValorCeldaSeguro(row("EstadoPago"), "") & """",
                    """" & ObtenerValorCeldaSeguro(row("Observaciones"), "").Replace("""", """""") & """"
                })
                writer.WriteLine(linea)
            Next

            writer.WriteLine()
            writer.WriteLine("RESUMEN")
            writer.WriteLine("Total de registros:," & dtPagos.Rows.Count.ToString())
            writer.WriteLine("Período consultado:," & dtpDesde.Value.ToString("dd/MM/yyyy") & " - " & dtpHasta.Value.ToString("dd/MM/yyyy"))
            writer.WriteLine("Fecha de exportación:," & DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
        End Using
    End Sub

    Private Function ObtenerValorCeldaSeguro(valor As Object, valorPorDefecto As String) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return valorPorDefecto
        End If
        Return valor.ToString().Trim()
    End Function

    ' ============================================================================
    ' MÉTODOS PARA PDF
    ' ============================================================================

    Private Sub PrepararDatosParaPDF()
        datosPagoParaImprimir = New List(Of Dictionary(Of String, Object))()

        If cmbApartamento.SelectedIndex = 0 Then
            apartamentoSeleccionadoPDF = "TODOS LOS APARTAMENTOS"
        Else
            apartamentoSeleccionadoPDF = cmbApartamento.SelectedItem.ToString()
        End If

        fechaImpresionPDF = DateTime.Now.ToString("dd/MMMM/yyyy HH:mm", New CultureInfo("es-ES"))

        For Each row As DataRow In dtPagos.Rows
            Dim pago As New Dictionary(Of String, Object)()
            pago("Fecha") = Convert.ToDateTime(row("Fecha")).ToString("dd/MM/yyyy", New CultureInfo("es-ES"))
            pago("Recibo") = row("NumeroRecibo").ToString()
            pago("Apartamento") = row("Apartamento").ToString()
            pago("Residente") = row("Residente").ToString()
            pago("PagoAdm") = Convert.ToDecimal(row("PagoAdministracion"))
            pago("PagoInt") = Convert.ToDecimal(row("PagoIntereses"))
            pago("Total") = Convert.ToDecimal(row("TotalPagado"))
            pago("SaldoAnterior") = Convert.ToDecimal(row("SaldoAnterior"))
            pago("SaldoActual") = Convert.ToDecimal(row("SaldoActual"))
            pago("Estado") = row("EstadoPago").ToString()
            pago("Observacion") = row("Observaciones").ToString()
            datosPagoParaImprimir.Add(pago)
        Next

        totalPaginas = Math.Ceiling(datosPagoParaImprimir.Count / filasPorPagina)
        paginaActual = 1
    End Sub

    Private Sub PrintDocument_PrintPage(sender As Object, e As PrintPageEventArgs) Handles printDocument.PrintPage
        Try
            Dim graphics As Graphics = e.Graphics
            Dim fontSmall As New Font("Arial", 8)
            Dim fontNormal As New Font("Arial", 8)
            Dim fontBold As New Font("Arial", 8, FontStyle.Bold)
            Dim fontTitle As New Font("Arial", 15, FontStyle.Bold)

            Dim brush As New SolidBrush(Color.Black)
            Dim pen As New Pen(Color.Black, 0.5)

            Dim y As Integer = e.MarginBounds.Top
            Dim x As Integer = e.MarginBounds.Left
            Dim lineHeight As Integer = 12

            ' ENCABEZADO
            Dim titulo As String = "HISTÓRICO DE PAGOS - COOPDIASAM"
            Dim tituloSize As SizeF = graphics.MeasureString(titulo, fontTitle)
            graphics.DrawString(titulo, fontTitle, brush, x + (e.MarginBounds.Width - tituloSize.Width) / 2, y)
            y += 25

            graphics.DrawString("CONJUNTO HABITACIONAL COOPDIASAM", fontBold, brush, x, y)
            graphics.DrawString("FECHA: " & fechaImpresionPDF, fontNormal, brush, e.MarginBounds.Right - 200, y)
            y += lineHeight

            graphics.DrawString("NIT: 900.225.635-8", fontNormal, brush, x, y)
            graphics.DrawString("Página " & paginaActual.ToString() & " de " & totalPaginas.ToString(), fontNormal, brush, e.MarginBounds.Right - 100, y)
            y += lineHeight

            graphics.DrawString("APARTAMENTO: " & apartamentoSeleccionadoPDF, fontBold, brush, x, y)
            y += lineHeight

            graphics.DrawString("PERÍODO: " & dtpDesde.Value.ToString("dd/MM/yyyy") & " - " & dtpHasta.Value.ToString("dd/MM/yyyy"), fontNormal, brush, x, y)
            y += 20

            ' TABLA DE DATOS
            Dim colWidths() As Integer = {65, 130, 100, 180, 100, 60, 80, 90, 60, 240}
            Dim headers() As String = {"FECHA", "RECIBO", "APART.", "RESIDENTE", "PAGO ADM", "INTERÉS", "TOTAL", "SALDO ANT", "ESTADO", "OBSERVACIÓN"}

            ' Encabezados
            Dim headerY As Integer = y
            Dim currentX As Integer = x

            For i As Integer = 0 To headers.Length - 1
                graphics.FillRectangle(New SolidBrush(Color.LightGray), currentX, headerY, colWidths(i), 18)
                graphics.DrawRectangle(pen, currentX, headerY, colWidths(i), 18)

                Dim headerSize As SizeF = graphics.MeasureString(headers(i), fontBold)
                Dim headerX As Single = currentX + (colWidths(i) - headerSize.Width) / 2
                graphics.DrawString(headers(i), fontBold, brush, headerX, headerY + 2)
                currentX += colWidths(i)
            Next

            y += 20

            ' Datos de la tabla
            Dim startIndex As Integer = (paginaActual - 1) * filasPorPagina
            Dim endIndex As Integer = Math.Min(startIndex + filasPorPagina - 1, datosPagoParaImprimir.Count - 1)

            For i As Integer = startIndex To endIndex
                If y > e.MarginBounds.Bottom - 80 Then Exit For

                Dim pago = datosPagoParaImprimir(i)
                currentX = x

                Dim cellData() As String = {
                    pago("Fecha").ToString(),
                    pago("Recibo").ToString(),
                    TruncateText(pago("Apartamento").ToString(), 15),
                    TruncateText(pago("Residente").ToString(), 25),
                    Convert.ToDecimal(pago("PagoAdm")).ToString("N0"),
                    Convert.ToDecimal(pago("PagoInt")).ToString("N0"),
                    Convert.ToDecimal(pago("Total")).ToString("N0"),
                    Convert.ToDecimal(pago("SaldoAnterior")).ToString("N0"),
                    TruncateText(pago("Estado").ToString(), 11),
                    TruncateText(pago("Observacion").ToString(), 50)
                }

                For j As Integer = 0 To cellData.Length - 1
                    graphics.DrawRectangle(pen, currentX, y, colWidths(j), 16)

                    Dim cellRect As New RectangleF(currentX + 2, y + 1, colWidths(j) - 4, 14)
                    Dim format As New StringFormat()

                    If j >= 4 AndAlso j <= 7 Then
                        format.Alignment = StringAlignment.Far
                    Else
                        format.Alignment = StringAlignment.Near
                    End If

                    format.LineAlignment = StringAlignment.Center
                    graphics.DrawString(cellData(j), fontSmall, brush, cellRect, format)
                    currentX += colWidths(j)
                Next

                y += 16
            Next

            ' Pie de página
            Dim footerY As Integer = e.MarginBounds.Bottom - 20
            graphics.DrawString("Generado el: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm"), fontSmall, brush, x, footerY)
            graphics.DrawString("Sistema de Gestión ClickApt", fontSmall, brush, e.MarginBounds.Right - 200, footerY)

            paginaActual += 1
            e.HasMorePages = (paginaActual <= totalPaginas)

        Catch ex As Exception
            MessageBox.Show("Error al generar página PDF: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function TruncateText(text As String, maxLength As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLength Then Return text
        Return text.Substring(0, maxLength - 3) & "..."
    End Function

    ' ============================================================================
    ' MÉTODOS AUXILIARES
    ' ============================================================================

    Private Sub FormHistorial_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Try
            ' Reposicionar el botón VOLVER centrado en el panel inferior
            If btnVolver IsNot Nothing Then
                btnVolver.Location = New Point((Me.Width - btnVolver.Width) \ 2, 10)
            End If

            ' Redimensionar panel de totales para usar todo el ancho disponible
            If lblTotalesPagos IsNot Nothing Then
                Dim anchoDisponible As Integer = Me.Width - 550
                If anchoDisponible > 400 Then
                    lblTotalesPagos.Width = anchoDisponible
                End If
            End If

        Catch ex As Exception
            ' Ignorar errores de redimensionamiento
        End Try
    End Sub

    Private Sub DiagnosticarFechasProblematicas()
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT CAST(fecha_pago AS TEXT) as fecha_texto, COUNT(*) as cantidad FROM pagos WHERE fecha_pago IS NOT NULL GROUP BY fecha_pago ORDER BY cantidad DESC LIMIT 10"

                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        Console.WriteLine("=== DIAGNÓSTICO DE FECHAS ===")
                        While reader.Read()
                            Dim fechaTexto As String = reader("fecha_texto").ToString()
                            Dim cantidad As Integer = Convert.ToInt32(reader("cantidad"))
                            Console.WriteLine("Fecha: '" & fechaTexto & "' | Registros: " & cantidad.ToString())
                        End While
                    End Using
                End Using

            End Using
        Catch ex As Exception
            Console.WriteLine("Error en diagnóstico: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            If dtPagos IsNot Nothing Then
                dtPagos.Dispose()
            End If
            If printDocument IsNot Nothing Then
                printDocument.Dispose()
            End If
        Catch ex As Exception
            ' Ignorar errores de limpieza
        End Try

        MyBase.OnFormClosed(e)
    End Sub

End Class