Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data.SQLite

Public Class FormEstados
    Inherits Form

    Private dgvEstados As DataGridView
    Private lblTitulo As Label
    Private WithEvents btnActualizar As Button
    Private WithEvents btnVolver As Button
    Private lblContadores As Label
    Private txtBuscar As TextBox
    Private lblBuscar As Label
    Private WithEvents btnBuscar As Button
    Private WithEvents btnLimpiarBusqueda As Button
    Private datosCompletos As DataTable ' Para almacenar todos los datos sin filtrar

    Private Sub FormEstados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarEstados()
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configuración de ventana completa - MISMO ESTILO QUE FORMHISTORIAL
            Me.Text = "Estados de Cuentas - COOPDIASAM"
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

            lblTitulo = New Label()
            lblTitulo.Text = "📊 ESTADOS DE CUENTAS"
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
        ' Panel de filtros con mejor organización y altura REDUCIDA
        Dim panelFiltros As New Panel()
        panelFiltros.Dock = DockStyle.Top
        panelFiltros.Height = 100  ' REDUCIDO PARA ELIMINAR ESPACIO EXTRA
        panelFiltros.BackColor = Color.FromArgb(248, 249, 250)  ' MISMO COLOR
        panelFiltros.Padding = New Padding(20)  ' PADDING REDUCIDO

        ' === PRIMERA FILA DE CONTROLES ===
        ' Búsqueda
        lblBuscar = New Label()
        lblBuscar.Text = "🔍 Buscar por apartamento, residente, teléfono o correo:"
        lblBuscar.Location = New Point(20, 15)  ' POSICIÓN AJUSTADA
        lblBuscar.Size = New Size(400, 25)
        lblBuscar.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblBuscar.ForeColor = Color.FromArgb(52, 73, 94)
        lblBuscar.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblBuscar)

        txtBuscar = New TextBox()
        txtBuscar.Location = New Point(430, 13)  ' POSICIÓN AJUSTADA
        txtBuscar.Size = New Size(300, 30)
        txtBuscar.Font = New Font("Segoe UI", 11)
        txtBuscar.BorderStyle = BorderStyle.FixedSingle
        AddHandler txtBuscar.KeyPress, AddressOf txtBuscar_KeyPress
        panelFiltros.Controls.Add(txtBuscar)

        ' === SEGUNDA FILA - BOTONES MÁS COMPACTOS ===
        btnActualizar = New Button()
        btnActualizar.Text = "🔄 ACTUALIZAR"
        btnActualizar.Location = New Point(20, 50)  ' POSICIÓN AJUSTADA
        btnActualizar.Size = New Size(130, 35)  ' TAMAÑO REDUCIDO
        btnActualizar.BackColor = Color.FromArgb(39, 174, 96)
        btnActualizar.ForeColor = Color.White
        btnActualizar.FlatStyle = FlatStyle.Flat
        btnActualizar.Cursor = Cursors.Hand
        btnActualizar.Font = New Font("Segoe UI", 7, FontStyle.Bold)  ' FUENTE REDUCIDA
        btnActualizar.FlatAppearance.BorderSize = 0
        btnActualizar.FlatAppearance.MouseOverBackColor = Color.FromArgb(34, 153, 84)
        AddHandler btnActualizar.Click, AddressOf btnActualizar_Click
        panelFiltros.Controls.Add(btnActualizar)

        btnBuscar = New Button()
        btnBuscar.Text = "🔍 BUSCAR"
        btnBuscar.Location = New Point(160, 50)  ' POSICIÓN AJUSTADA
        btnBuscar.Size = New Size(110, 35)  ' TAMAÑO REDUCIDO
        btnBuscar.BackColor = Color.FromArgb(52, 152, 219)
        btnBuscar.ForeColor = Color.White
        btnBuscar.FlatStyle = FlatStyle.Flat
        btnBuscar.Cursor = Cursors.Hand
        btnBuscar.Font = New Font("Segoe UI", 10, FontStyle.Bold)  ' FUENTE REDUCIDA
        btnBuscar.FlatAppearance.BorderSize = 0
        btnBuscar.FlatAppearance.MouseOverBackColor = Color.FromArgb(41, 128, 185)
        AddHandler btnBuscar.Click, AddressOf btnBuscar_Click
        panelFiltros.Controls.Add(btnBuscar)

        btnLimpiarBusqueda = New Button()
        btnLimpiarBusqueda.Text = "🧹 LIMPIAR"
        btnLimpiarBusqueda.Location = New Point(280, 50)  ' POSICIÓN AJUSTADA
        btnLimpiarBusqueda.Size = New Size(110, 35)  ' TAMAÑO REDUCIDO
        btnLimpiarBusqueda.BackColor = Color.FromArgb(149, 165, 166)
        btnLimpiarBusqueda.ForeColor = Color.White
        btnLimpiarBusqueda.FlatStyle = FlatStyle.Flat
        btnLimpiarBusqueda.Cursor = Cursors.Hand
        btnLimpiarBusqueda.Font = New Font("Segoe UI", 10, FontStyle.Bold)  ' FUENTE REDUCIDA
        btnLimpiarBusqueda.FlatAppearance.BorderSize = 0
        btnLimpiarBusqueda.FlatAppearance.MouseOverBackColor = Color.FromArgb(127, 140, 141)
        AddHandler btnLimpiarBusqueda.Click, AddressOf btnLimpiarBusqueda_Click
        panelFiltros.Controls.Add(btnLimpiarBusqueda)

        ' === PANEL DE ESTADÍSTICAS MÁS COMPACTO ===
        lblContadores = New Label()
        lblContadores.Location = New Point(410, 50)  ' POSICIÓN AJUSTADA
        lblContadores.Size = New Size(400, 35)  ' ALTURA REDUCIDA
        lblContadores.Font = New Font("Segoe UI", 10, FontStyle.Bold)  ' FUENTE REDUCIDA
        lblContadores.ForeColor = Color.FromArgb(44, 62, 80)
        lblContadores.Text = "📊 Actualizando estadísticas..."
        lblContadores.BackColor = Color.FromArgb(236, 240, 241)
        lblContadores.BorderStyle = BorderStyle.FixedSingle
        lblContadores.TextAlign = ContentAlignment.MiddleLeft  ' CENTRADO VERTICAL
        lblContadores.Padding = New Padding(10, 5, 10, 5)  ' PADDING REDUCIDO
        lblContadores.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right  ' ANCLAJE PARA OCUPAR TODO EL ANCHO
        panelFiltros.Controls.Add(lblContadores)

        ' DataGridView principal
        dgvEstados = New DataGridView()
        dgvEstados.Dock = DockStyle.Fill
        dgvEstados.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill  ' LLENAR TODO EL ANCHO
        dgvEstados.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvEstados.ReadOnly = True
        dgvEstados.AllowUserToAddRows = False
        dgvEstados.AllowUserToDeleteRows = False
        dgvEstados.BackgroundColor = Color.White
        dgvEstados.BorderStyle = BorderStyle.None
        dgvEstados.RowHeadersVisible = False
        dgvEstados.Font = New Font("Segoe UI", 10.0F)
        dgvEstados.ScrollBars = ScrollBars.Both
        dgvEstados.AllowUserToResizeColumns = True
        dgvEstados.MultiSelect = False
        ConfigurarDataGridView(dgvEstados)

        ' Agregar al formulario
        Me.Controls.Add(dgvEstados)
        Me.Controls.Add(panelFiltros)
    End Sub

    Private Sub ConfigurarDataGridView(dgv As DataGridView)
        ' Configuración visual moderna y profesional - CABECERAS IGUALES A FORMHISTORIAL
        dgv.EnableHeadersVisualStyles = False

        ' Encabezados con diseño moderno - LETRAS MÁS PEQUEÑAS (IGUAL QUE FORMHISTORIAL)
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(44, 62, 80)  ' MISMO COLOR
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)  ' FUENTE MÁS PEQUEÑA
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(3, 8, 3, 8)  ' PADDING REDUCIDO
        dgv.ColumnHeadersHeight = 35  ' ALTURA REDUCIDA

        ' Estilo de celdas mejorado - LETRAS MÁS PEQUEÑAS (IGUAL QUE FORMHISTORIAL)
        dgv.DefaultCellStyle.Font = New Font("Segoe UI", 8.5F)  ' FUENTE MÁS PEQUEÑA
        dgv.DefaultCellStyle.Padding = New Padding(6, 4, 6, 4)  ' PADDING REDUCIDO
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv.DefaultCellStyle.ForeColor = Color.FromArgb(52, 73, 94)
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
        dgv.DefaultCellStyle.SelectionForeColor = Color.White

        ' Filas alternadas con color suave
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 249, 250)
        dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.FromArgb(52, 73, 94)

        ' Configuración de filas y bordes - FILAS MÁS PEQUEÑAS (IGUAL QUE FORMHISTORIAL)
        dgv.RowTemplate.Height = 32  ' FILAS MÁS PEQUEÑAS
        dgv.GridColor = Color.FromArgb(189, 195, 199)
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single  ' BORDES VISIBLES
        dgv.MultiSelect = False
        dgv.AutoGenerateColumns = True
    End Sub

    Private Sub CargarEstados()
        Try
            Me.Cursor = Cursors.WaitCursor
            lblContadores.Text = "🔄 Cargando datos..."

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' CONSULTA CORREGIDA - Usando los nombres correctos de las columnas
                Dim consulta As String = "
                    SELECT 
                        a.id_apartamentos as 'ID',
                        a.id_torre as 'Torre',
                        a.id_piso as 'Piso',
                        a.numero_apartamento as 'Apartamento',
                        COALESCE(a.nombre_residente, 'Sin asignar') as 'Residente',
                        COALESCE(a.telefono, '-') as 'Teléfono',
                        COALESCE(a.correo, '-') as 'Correo',
                        COALESCE(ultimo_saldo.saldo_actual, 0) as 'Saldo Actual',
                        CASE 
                            WHEN COALESCE(ultimo_saldo.saldo_actual, 0) = 0 THEN 'Al día'
                            WHEN COALESCE(ultimo_saldo.saldo_actual, 0) < 0 THEN 'A favor'
                            WHEN COALESCE(ultimo_saldo.saldo_actual, 0) > 0 THEN 'Pendiente'
                            ELSE 'Sin movimientos'
                        END as 'Estado',
                        COALESCE(cuotas_pendientes.cantidad, 0) as 'Cuotas Pendientes',
                        COALESCE(ultimo_saldo.fecha_pago, 'Sin pagos') as 'Último Pago'
                    FROM Apartamentos a
                    LEFT JOIN (
                        SELECT 
                            p.id_apartamentos,
                            p.saldo_actual,
                            p.fecha_pago,
                            ROW_NUMBER() OVER (PARTITION BY p.id_apartamentos ORDER BY p.fecha_pago DESC) as rn
                        FROM pagos p
                    ) ultimo_saldo ON a.id_apartamentos = ultimo_saldo.id_apartamentos AND ultimo_saldo.rn = 1
                    LEFT JOIN (
                        SELECT 
                            cga.id_apartamentos,
                            COUNT(*) as cantidad
                        FROM cuotas_generadas_apartamento cga
                        WHERE cga.estado = 'pendiente'
                        GROUP BY cga.id_apartamentos
                    ) cuotas_pendientes ON a.id_apartamentos = cuotas_pendientes.id_apartamentos
                    ORDER BY a.id_torre, a.id_piso, a.numero_apartamento"

                Using adapter As New Data.SQLite.SQLiteDataAdapter(consulta, conexion)
                    datosCompletos = New DataTable()  ' Guardar todos los datos
                    adapter.Fill(datosCompletos)
                    dgvEstados.DataSource = datosCompletos

                    ' Configurar columnas
                    ConfigurarColumnas()

                    ' Aplicar colores según el estado
                    AplicarColoresPorEstado()

                    ' Actualizar contadores
                    ActualizarContadores(datosCompletos)
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al cargar estados: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ConfigurarColumnas()
        Try
            ' Asegurar que las columnas están configuradas correctamente
            If dgvEstados.Columns.Count = 0 Then Return

            ' Ocultar ID
            If dgvEstados.Columns.Contains("ID") Then
                dgvEstados.Columns("ID").Visible = False
            End If

            ' CONFIGURAR ENCABEZADOS Y ANCHOS - ESTILO FORMHISTORIAL
            If dgvEstados.Columns.Contains("Torre") Then
                With dgvEstados.Columns("Torre")
                    .HeaderText = "TORRE"
                    .FillWeight = 8
                    .MinimumWidth = 30
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                End With
            End If

            If dgvEstados.Columns.Contains("Piso") Then
                With dgvEstados.Columns("Piso")
                    .HeaderText = "PISO"
                    .FillWeight = 8
                    .MinimumWidth = 30
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                End With
            End If

            If dgvEstados.Columns.Contains("Apartamento") Then
                With dgvEstados.Columns("Apartamento")
                    .HeaderText = "APART-"
                    .FillWeight = 10
                    .MinimumWidth = 30
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                End With
            End If

            If dgvEstados.Columns.Contains("Residente") Then
                With dgvEstados.Columns("Residente")
                    .HeaderText = "RESIDENTE"
                    .FillWeight = 25
                    .MinimumWidth = 150
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                End With
            End If

            If dgvEstados.Columns.Contains("Teléfono") Then
                With dgvEstados.Columns("Teléfono")
                    .HeaderText = "TELÉFONO"
                    .FillWeight = 15
                    .MinimumWidth = 120
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Consolas", 8)
                End With
            End If

            If dgvEstados.Columns.Contains("Correo") Then
                With dgvEstados.Columns("Correo")
                    .HeaderText = "CORREO ELECTRÓNICO"
                    .FillWeight = 20
                    .MinimumWidth = 180
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                End With
            End If

            If dgvEstados.Columns.Contains("Saldo Actual") Then
                With dgvEstados.Columns("Saldo Actual")
                    .HeaderText = "SALDO ACTUAL"
                    .FillWeight = 12
                    .MinimumWidth = 150
                    .DefaultCellStyle.Format = "C0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .DefaultCellStyle.Font = New Font("Consolas", 8, FontStyle.Bold)
                End With
            End If

            If dgvEstados.Columns.Contains("Estado") Then
                With dgvEstados.Columns("Estado")
                    .HeaderText = "ESTADO"
                    .FillWeight = 10
                    .MinimumWidth = 120
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                End With
            End If

            If dgvEstados.Columns.Contains("Cuotas Pendientes") Then
                With dgvEstados.Columns("Cuotas Pendientes")
                    .HeaderText = "CUOTAS PEND."
                    .FillWeight = 8
                    .MinimumWidth = 120
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Consolas", 8)
                End With
            End If

            If dgvEstados.Columns.Contains("Último Pago") Then
                With dgvEstados.Columns("Último Pago")
                    .HeaderText = "ÚLTIMO PAGO"
                    .FillWeight = 12
                    .MinimumWidth = 150
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Font = New Font("Segoe UI", 8)
                End With
            End If

        Catch ex As Exception
            Console.WriteLine($"Error configurando columnas: {ex.Message}")
        End Try
    End Sub

    Private Sub AplicarColoresPorEstado()
        Try
            For Each row As DataGridViewRow In dgvEstados.Rows
                If row.Cells("Estado") IsNot Nothing AndAlso row.Cells("Estado").Value IsNot Nothing Then
                    Select Case row.Cells("Estado").Value.ToString()
                        Case "Al día"
                            row.Cells("Estado").Style.BackColor = Color.FromArgb(40, 167, 69)
                            row.Cells("Estado").Style.ForeColor = Color.White
                            row.Cells("Estado").Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                        Case "Pendiente"
                            row.Cells("Estado").Style.BackColor = Color.FromArgb(220, 53, 69)
                            row.Cells("Estado").Style.ForeColor = Color.White
                            row.Cells("Estado").Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                        Case "A favor"
                            row.Cells("Estado").Style.BackColor = Color.FromArgb(23, 162, 184)
                            row.Cells("Estado").Style.ForeColor = Color.White
                            row.Cells("Estado").Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                        Case "Sin movimientos"
                            row.Cells("Estado").Style.BackColor = Color.FromArgb(255, 193, 7)
                            row.Cells("Estado").Style.ForeColor = Color.Black
                            row.Cells("Estado").Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                    End Select
                End If

                ' Resaltar saldos pendientes (similar a FormHistorial)
                If row.Cells("Saldo Actual") IsNot Nothing AndAlso row.Cells("Saldo Actual").Value IsNot Nothing Then
                    Dim saldo As Decimal = Convert.ToDecimal(row.Cells("Saldo Actual").Value)
                    If saldo > 0 Then
                        row.Cells("Saldo Actual").Style.BackColor = Color.FromArgb(255, 243, 243)
                        row.Cells("Saldo Actual").Style.ForeColor = Color.FromArgb(169, 68, 66)
                        row.Cells("Saldo Actual").Style.Font = New Font("Consolas", 9, FontStyle.Bold)
                    ElseIf saldo < 0 Then
                        row.Cells("Saldo Actual").Style.BackColor = Color.FromArgb(240, 248, 255)
                        row.Cells("Saldo Actual").Style.ForeColor = Color.FromArgb(52, 73, 94)
                        row.Cells("Saldo Actual").Style.Font = New Font("Consolas", 9, FontStyle.Bold)
                    End If
                End If
            Next
        Catch ex As Exception
            Console.WriteLine($"Error aplicando colores: {ex.Message}")
        End Try
    End Sub

    Private Sub ActualizarContadores(tabla As DataTable)
        Try
            Dim totalApartamentos As Integer = tabla.Rows.Count
            Dim alDia As Integer = tabla.Select("Estado = 'Al día'").Length
            Dim pendientes As Integer = tabla.Select("Estado = 'Pendiente'").Length
            Dim aFavor As Integer = tabla.Select("Estado = 'A favor'").Length
            Dim sinMovimientos As Integer = tabla.Select("Estado = 'Sin movimientos'").Length

            lblContadores.Text = String.Format(
                "📊 RESUMEN: {0} apartamentos - 🟢 Al día: {1} - 🔴 Pendientes: {2} - 🔵 A favor: {3} - ⚪ Sin movimientos: {4} - ⏱️ Actualizado: {5}",
                totalApartamentos,
                alDia,
                pendientes,
                aFavor,
                sinMovimientos,
                DateTime.Now.ToString("HH:mm")
            )
            lblContadores.BackColor = Color.FromArgb(212, 237, 218)
            lblContadores.ForeColor = Color.FromArgb(21, 87, 36)

        Catch ex As Exception
            lblContadores.Text = "⚠️ Error al calcular estadísticas"
            lblContadores.BackColor = Color.FromArgb(248, 215, 218)
            lblContadores.ForeColor = Color.FromArgb(114, 28, 36)
        End Try
    End Sub

    ' ============================================================================
    ' EVENTOS DE INTERFAZ
    ' ============================================================================

    Private Sub btnActualizar_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            lblContadores.Text = "🔄 Actualizando datos..."
            CargarEstados()
        Catch ex As Exception
            MessageBox.Show("Error al actualizar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnVolver_Click(sender As Object, e As EventArgs) Handles btnVolver.Click
        Try
            Me.Close()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    ' MÉTODOS PARA LA FUNCIONALIDAD DE BÚSQUEDA
    Private Sub txtBuscar_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            RealizarBusqueda()
            e.Handled = True
        End If
    End Sub

    Private Sub RealizarBusqueda()
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim textoBusqueda As String = txtBuscar.Text.Trim()

            If String.IsNullOrEmpty(textoBusqueda) Then
                dgvEstados.DataSource = datosCompletos
                ActualizarContadores(datosCompletos)
            Else
                Dim tablaFiltrada As New DataTable()
                tablaFiltrada = datosCompletos.Clone()

                For Each fila As DataRow In datosCompletos.Rows
                    Dim coincide As Boolean = False

                    If fila("Torre").ToString().Contains(textoBusqueda) Then coincide = True
                    If fila("Piso").ToString().Contains(textoBusqueda) Then coincide = True
                    If fila("Apartamento").ToString().Contains(textoBusqueda) Then coincide = True
                    If fila("Residente").ToString().ToUpper().Contains(textoBusqueda.ToUpper()) Then coincide = True
                    If fila("Teléfono").ToString().Contains(textoBusqueda) Then coincide = True
                    If fila("Correo").ToString().ToUpper().Contains(textoBusqueda.ToUpper()) Then coincide = True
                    If fila("Estado").ToString().ToUpper().Contains(textoBusqueda.ToUpper()) Then coincide = True

                    If coincide Then
                        tablaFiltrada.ImportRow(fila)
                    End If
                Next

                dgvEstados.DataSource = tablaFiltrada

                If tablaFiltrada.Rows.Count > 0 Then
                    ActualizarContadores(tablaFiltrada)
                    lblContadores.BackColor = Color.FromArgb(212, 237, 218)
                    lblContadores.ForeColor = Color.FromArgb(21, 87, 36)
                Else
                    lblContadores.Text = $"🔍 No se encontraron resultados para '{textoBusqueda}'"
                    lblContadores.BackColor = Color.FromArgb(248, 215, 218)
                    lblContadores.ForeColor = Color.FromArgb(114, 28, 36)
                End If
            End If

            ConfigurarColumnas()
            AplicarColoresPorEstado()

        Catch ex As Exception
            MessageBox.Show($"Error en la búsqueda: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LimpiarBusqueda()
        Try
            txtBuscar.Text = ""
            dgvEstados.DataSource = datosCompletos
            ConfigurarColumnas()
            AplicarColoresPorEstado()
            ActualizarContadores(datosCompletos)
            txtBuscar.Focus()
        Catch ex As Exception
            MessageBox.Show("Error al limpiar búsqueda: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Eventos de botones de búsqueda
    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        RealizarBusqueda()
    End Sub

    Private Sub btnLimpiarBusqueda_Click(sender As Object, e As EventArgs) Handles btnLimpiarBusqueda.Click
        LimpiarBusqueda()
    End Sub

    ' ============================================================================
    ' MÉTODOS AUXILIARES
    ' ============================================================================

    Private Sub FormEstados_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Try
            ' Reposicionar el botón VOLVER centrado en el panel inferior (IGUAL QUE FORMHISTORIAL)
            If btnVolver IsNot Nothing Then
                btnVolver.Location = New Point((Me.Width - btnVolver.Width) \ 2, 10)
            End If

            ' Redimensionar panel de estadísticas para usar todo el ancho disponible
            If lblContadores IsNot Nothing Then
                Dim anchoDisponible As Integer = Me.Width - 550
                If anchoDisponible > 400 Then
                    lblContadores.Width = anchoDisponible
                End If
            End If

        Catch ex As Exception
            ' Ignorar errores de redimensionamiento
        End Try
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            If datosCompletos IsNot Nothing Then
                datosCompletos.Dispose()
            End If
        Catch ex As Exception
            ' Ignorar errores de limpieza
        End Try

        MyBase.OnFormClosed(e)
    End Sub

End Class