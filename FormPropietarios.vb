Imports System.Windows.Forms
Imports System.Drawing

Public Class FormPropietarios
    Inherits Form

    Private dgvPropietarios As DataGridView
    Private lblTitulo As Label
    Private lblBuscador As Label
    Private txtBuscador As TextBox
    Private btnLimpiar As Button
    Private btnEditar As Button
    Private btnVolver As Button
    Private lblEstadisticas As Label
    Private listaCompletaApartamentos As List(Of Apartamento)

    Private Sub FormPropietarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarPropietarios()
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configuración de ventana completa - MISMO ESTILO
            Me.Text = "Propietarios - COOPDIASAM"
            Me.WindowState = FormWindowState.Maximized
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MinimumSize = New Size(1400, 700)
            Me.MaximizeBox = True
            Me.MinimizeBox = True
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' Panel superior con mejor diseño - MISMO COLOR AZUL
            Dim panelSuperior As New Panel()
            panelSuperior.Dock = DockStyle.Top
            panelSuperior.Height = 80
            panelSuperior.BackColor = Color.FromArgb(46, 132, 188)

            lblTitulo = New Label()
            lblTitulo.Text = "📊 GESTIÓN DE PROPIETARIOS"
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
            AddHandler btnVolver.Click, AddressOf btnVolver_Click
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
        panelFiltros.Height = 180  ' AUMENTADO PARA ACOMODAR TODO
        panelFiltros.BackColor = Color.FromArgb(248, 249, 250)
        panelFiltros.Padding = New Padding(30)

        ' === PRIMERA FILA DE CONTROLES ===
        ' Label del buscador
        lblBuscador = New Label()
        lblBuscador.Text = "Buscar por Apartamento (ej: 1202 para Torre 1, Apt 202):"
        lblBuscador.Location = New Point(30, 20)
        lblBuscador.Size = New Size(400, 25)
        lblBuscador.Font = New Font("Segoe UI", 11, FontStyle.Italic)
        lblBuscador.ForeColor = Color.FromArgb(52, 73, 94)
        lblBuscador.TextAlign = ContentAlignment.MiddleLeft
        panelFiltros.Controls.Add(lblBuscador)

        ' TextBox del buscador
        txtBuscador = New TextBox()
        txtBuscador.Location = New Point(30, 55)
        txtBuscador.Size = New Size(300, 30)
        txtBuscador.Font = New Font("Segoe UI", 11)
        txtBuscador.BorderStyle = BorderStyle.FixedSingle
        AddHandler txtBuscador.TextChanged, AddressOf txtBuscador_TextChanged
        panelFiltros.Controls.Add(txtBuscador)

        ' Botón Limpiar búsqueda
        btnLimpiar = New Button()
        btnLimpiar.Text = "🧹 LIMPIAR"
        btnLimpiar.Location = New Point(340, 55)
        btnLimpiar.Size = New Size(120, 30)
        btnLimpiar.BackColor = Color.FromArgb(149, 165, 166)
        btnLimpiar.ForeColor = Color.White
        btnLimpiar.FlatStyle = FlatStyle.Flat
        btnLimpiar.Cursor = Cursors.Hand
        btnLimpiar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnLimpiar.FlatAppearance.BorderSize = 0
        btnLimpiar.FlatAppearance.MouseOverBackColor = Color.FromArgb(127, 140, 141)
        AddHandler btnLimpiar.Click, AddressOf btnLimpiar_Click
        panelFiltros.Controls.Add(btnLimpiar)

        ' === SEGUNDA FILA - BOTÓN EDITAR ===
        btnEditar = New Button()
        btnEditar.Text = "✏️ EDITAR"
        btnEditar.Location = New Point(30, 100)
        btnEditar.Size = New Size(180, 45)
        btnEditar.BackColor = Color.FromArgb(231, 76, 60) ' Mismo color base
        btnEditar.ForeColor = Color.White
        btnEditar.FlatStyle = FlatStyle.Flat
        btnEditar.Cursor = Cursors.Hand
        btnEditar.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnEditar.FlatAppearance.BorderSize = 0
        btnEditar.FlatAppearance.MouseOverBackColor = Color.FromArgb(192, 57, 43) ' Mismo color hover
        AddHandler btnEditar.Click, AddressOf btnEditar_Click
        panelFiltros.Controls.Add(btnEditar)
        btnEditar.TextAlign = ContentAlignment.MiddleCenter
        btnEditar.Padding = New Padding(10, 0, 0, 0) ' Mueve el texto un poco a la derecha


        ' === PANEL DE ESTADÍSTICAS VERTICAL Y MÁS GRANDE ===
        lblEstadisticas = New Label()
        lblEstadisticas.Location = New Point(220, 100)
        lblEstadisticas.Size = New Size(300, 45)  ' MÁS ANCHO PARA EVITAR ESPACIO EN BLANCO
        lblEstadisticas.Font = New Font("Segoe UI", 10, FontStyle.Regular)  ' FUENTE MÁS GRANDE
        lblEstadisticas.ForeColor = Color.FromArgb(44, 62, 80)
        lblEstadisticas.Text = "📊 Seleccione un apartamento para ver detalles del propietario"
        lblEstadisticas.BackColor = Color.FromArgb(236, 240, 241)
        lblEstadisticas.BorderStyle = BorderStyle.FixedSingle
        lblEstadisticas.TextAlign = ContentAlignment.TopLeft
        lblEstadisticas.Padding = New Padding(15, 8, 15, 8)
        lblEstadisticas.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right  ' ANCLAJE PARA OCUPAR TODO EL ANCHO
        panelFiltros.Controls.Add(lblEstadisticas)

        ' DataGridView principal
        dgvPropietarios = New DataGridView()
        dgvPropietarios.Dock = DockStyle.Fill
        dgvPropietarios.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None  ' CAMBIADO PARA CONTROL MANUAL
        dgvPropietarios.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvPropietarios.ReadOnly = True
        dgvPropietarios.AllowUserToAddRows = False
        dgvPropietarios.AllowUserToDeleteRows = False
        dgvPropietarios.BackgroundColor = Color.White
        dgvPropietarios.BorderStyle = BorderStyle.None
        dgvPropietarios.RowHeadersVisible = False
        dgvPropietarios.Font = New Font("Segoe UI", 8.5F)
        dgvPropietarios.ScrollBars = ScrollBars.Both
        dgvPropietarios.AllowUserToResizeColumns = True
        dgvPropietarios.MultiSelect = False
        ' Evento para actualizar estadísticas al seleccionar
        AddHandler dgvPropietarios.SelectionChanged, AddressOf dgvPropietarios_SelectionChanged
        ConfigurarDataGridView(dgvPropietarios)

        ' Agregar al formulario
        Me.Controls.Add(dgvPropietarios)
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
        dgv.ColumnHeadersHeight = 35  ' ALTURA REDUCIDA

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

    Private Sub CargarPropietarios(Optional filtroApartamento As String = "")
        Try
            listaCompletaApartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()

            If listaCompletaApartamentos IsNot Nothing Then
                Dim listaFiltrada As List(Of Apartamento) = listaCompletaApartamentos

                ' Aplicar filtro si se proporciona
                If Not String.IsNullOrWhiteSpace(filtroApartamento) Then
                    listaFiltrada = FiltrarApartamentos(listaCompletaApartamentos, filtroApartamento)
                End If

                dgvPropietarios.DataSource = listaFiltrada
                FormatearColumnas()
                ActualizarEstadisticas(listaFiltrada)

            End If
        Catch ex As Exception
            MessageBox.Show($"Error al cargar la lista de propietarios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FormatearColumnas()
        If dgvPropietarios.Columns.Count > 0 Then
            ' Ocultar columnas no relevantes
            If dgvPropietarios.Columns.Contains("IdApartamento") Then dgvPropietarios.Columns("IdApartamento").Visible = False

            ' Configurar columnas visibles con iconos y anchos específicos
            If dgvPropietarios.Columns.Contains("Torre") Then
                dgvPropietarios.Columns("Torre").HeaderText = "🏢 TORRE"
                dgvPropietarios.Columns("Torre").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvPropietarios.Columns("Torre").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                dgvPropietarios.Columns("Torre").Width = 90  ' MÁS ANGOSTA
                dgvPropietarios.Columns("Torre").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            End If

            If dgvPropietarios.Columns.Contains("Piso") Then
                dgvPropietarios.Columns("Piso").HeaderText = "🏠 PISO"
                dgvPropietarios.Columns("Piso").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvPropietarios.Columns("Piso").DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                dgvPropietarios.Columns("Piso").Width = 90  ' MÁS ANGOSTA
                dgvPropietarios.Columns("Piso").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            End If

            If dgvPropietarios.Columns.Contains("NumeroApartamento") Then
                dgvPropietarios.Columns("NumeroApartamento").HeaderText = "🚪 APARTAMENTO"
                dgvPropietarios.Columns("NumeroApartamento").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvPropietarios.Columns("NumeroApartamento").DefaultCellStyle.Font = New Font("Consolas", 8, FontStyle.Bold)
                dgvPropietarios.Columns("NumeroApartamento").Width = 120  ' ANCHO FIJO
                dgvPropietarios.Columns("NumeroApartamento").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            End If

            If dgvPropietarios.Columns.Contains("NombreResidente") Then
                dgvPropietarios.Columns("NombreResidente").HeaderText = "👤 NOMBRE RESIDENTE"
                dgvPropietarios.Columns("NombreResidente").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dgvPropietarios.Columns("NombreResidente").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgvPropietarios.Columns("NombreResidente").FillWeight = 150  ' 40% del espacio restante
            End If

            If dgvPropietarios.Columns.Contains("Correo") Then
                dgvPropietarios.Columns("Correo").HeaderText = "📧 CORREO ELECTRÓNICO"
                dgvPropietarios.Columns("Correo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dgvPropietarios.Columns("Correo").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgvPropietarios.Columns("Correo").FillWeight = 150  ' 40% del espacio restante
            End If

            If dgvPropietarios.Columns.Contains("Telefono") Then
                dgvPropietarios.Columns("Telefono").HeaderText = "📱 TELÉFONO"
                dgvPropietarios.Columns("Telefono").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvPropietarios.Columns("Telefono").DefaultCellStyle.Font = New Font("Consolas", 8)
                dgvPropietarios.Columns("Telefono").Width = 120  ' ANCHO FIJO
                dgvPropietarios.Columns("Telefono").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            End If

            If dgvPropietarios.Columns.Contains("MatriculaInmobiliaria") Then
                dgvPropietarios.Columns("MatriculaInmobiliaria").HeaderText = "📋 MATRÍCULA"
                dgvPropietarios.Columns("MatriculaInmobiliaria").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvPropietarios.Columns("MatriculaInmobiliaria").DefaultCellStyle.Font = New Font("Consolas", 8)
                dgvPropietarios.Columns("MatriculaInmobiliaria").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgvPropietarios.Columns("MatriculaInmobiliaria").FillWeight = 60 ' 20% del espacio restante
            End If

            ' Ocultar columnas no relevantes
            If dgvPropietarios.Columns.Contains("Activo") Then dgvPropietarios.Columns("Activo").Visible = False
            If dgvPropietarios.Columns.Contains("FechaRegistro") Then dgvPropietarios.Columns("FechaRegistro").Visible = False
            If dgvPropietarios.Columns.Contains("SaldoActual") Then dgvPropietarios.Columns("SaldoActual").Visible = False
            If dgvPropietarios.Columns.Contains("EstadoCuenta") Then dgvPropietarios.Columns("EstadoCuenta").Visible = False
            If dgvPropietarios.Columns.Contains("UltimoPago") Then dgvPropietarios.Columns("UltimoPago").Visible = False
            If dgvPropietarios.Columns.Contains("TieneUltimoPago") Then dgvPropietarios.Columns("TieneUltimoPago").Visible = False
            If dgvPropietarios.Columns.Contains("DiasEnMora") Then dgvPropietarios.Columns("DiasEnMora").Visible = False
            If dgvPropietarios.Columns.Contains("TotalIntereses") Then dgvPropietarios.Columns("TotalIntereses").Visible = False
        End If
    End Sub

    Private Sub ActualizarEstadisticas(lista As List(Of Apartamento))
        If lista IsNot Nothing Then
            Dim totalApartamentos As Integer = lista.Count
            Dim conResidente As Integer = 0
            Dim conCorreo As Integer = 0
            Dim conTelefono As Integer = 0

            ' Contar apartamentos con datos
            For Each apartamento As Apartamento In lista
                If Not String.IsNullOrWhiteSpace(apartamento.NombreResidente) Then
                    conResidente += 1
                End If
                If Not String.IsNullOrWhiteSpace(apartamento.Correo) Then
                    conCorreo += 1
                End If
                If Not String.IsNullOrWhiteSpace(apartamento.Telefono) Then
                    conTelefono += 1
                End If
            Next

            ' FORMATO VERTICAL CORREGIDO Y BIEN VISIBLE
            lblEstadisticas.Text = String.Format(
                "📊 RESUMEN: {0} apartamentos encontrados - 👤 Con residente: {1} ({2:F1}%) - 📧 Con correo: {3} ({4:F1}%)   📱 Con teléfono: {5} ({6:F1}%)",
                totalApartamentos,
                conResidente,
                If(totalApartamentos > 0, (conResidente / totalApartamentos) * 100, 0),
                conCorreo,
                If(totalApartamentos > 0, (conCorreo / totalApartamentos) * 100, 0),
                conTelefono,
                If(totalApartamentos > 0, (conTelefono / totalApartamentos) * 100, 0),
                DateTime.Now.ToString("HH:mm")
            )
            lblEstadisticas.BackColor = Color.FromArgb(248, 215, 218)
            lblEstadisticas.ForeColor = Color.FromArgb(114, 28, 36)
        Else
            lblEstadisticas.Text = "⚠️ No se pudieron cargar los datos de propietarios" & vbCrLf &
                                  "💡 Verifique la conexión a la base de datos"
            lblEstadisticas.BackColor = Color.FromArgb(248, 215, 218)
            lblEstadisticas.ForeColor = Color.FromArgb(114, 28, 36)
        End If
    End Sub

    Private Sub dgvPropietarios_SelectionChanged(sender As Object, e As EventArgs)
        Try
            If dgvPropietarios.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = dgvPropietarios.SelectedRows(0)
                Dim torre As String = If(selectedRow.Cells("Torre").Value?.ToString(), "N/A")
                Dim apartamento As String = If(selectedRow.Cells("NumeroApartamento").Value?.ToString(), "N/A")
                Dim residente As String = If(selectedRow.Cells("NombreResidente").Value?.ToString(), "Sin registrar")
                Dim correo As String = If(selectedRow.Cells("Correo").Value?.ToString(), "Sin correo")
                Dim telefono As String = If(selectedRow.Cells("Telefono").Value?.ToString(), "Sin teléfono")

                lblEstadisticas.Text = String.Format(
                    "🏢 APARTAMENTO SELECCIONADO: 📍 Torre {0} - Apartamento {1} - 👤 RESIDENTE: {2} - 📧 Correo: {3}   📱 Teléfono: {4}",
                    torre, apartamento, residente, correo, telefono
                )
                lblEstadisticas.BackColor = Color.FromArgb(64, 64, 64)      ' Fondo gris oscuro
                lblEstadisticas.ForeColor = Color.White                      ' Texto blanco para buen contraste

            End If
        Catch ex As Exception
            ' Ignorar errores de selección
        End Try
    End Sub

    Private Function FiltrarApartamentos(apartamentos As List(Of Apartamento), filtro As String) As List(Of Apartamento)
        If String.IsNullOrWhiteSpace(filtro) Then
            Return apartamentos
        End If

        Dim listaFiltrada As New List(Of Apartamento)

        For Each apartamento As Apartamento In apartamentos
            Dim coincide As Boolean = False

            ' Buscar por ID de apartamento completo (ej: 1202)
            If filtro.Length = 4 AndAlso IsNumeric(filtro) Then
                Dim torre As String = filtro.Substring(0, 1)
                Dim numeroApt As String = filtro.Substring(1, 3)

                If apartamento.Torre = torre AndAlso apartamento.NumeroApartamento = numeroApt Then
                    coincide = True
                End If
            End If

            ' Buscar por coincidencias parciales en torre, piso, apartamento
            If Not coincide Then
                If apartamento.Torre.ToString().Contains(filtro) OrElse
                   apartamento.Piso.ToString().Contains(filtro) OrElse
                   apartamento.NumeroApartamento.ToString().Contains(filtro) OrElse
                   (apartamento.NombreResidente IsNot Nothing AndAlso apartamento.NombreResidente.ToUpper().Contains(filtro.ToUpper())) Then
                    coincide = True
                End If
            End If

            If coincide Then
                listaFiltrada.Add(apartamento)
            End If
        Next

        Return listaFiltrada
    End Function

    Private Sub txtBuscador_TextChanged(sender As Object, e As EventArgs)
        Dim filtro As String = txtBuscador.Text.Trim()
        CargarPropietarios(filtro)
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs)
        txtBuscador.Text = ""
        CargarPropietarios()
    End Sub

    Private Sub btnEditar_Click(sender As Object, e As EventArgs)
        Try
            If dgvPropietarios.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = dgvPropietarios.SelectedRows(0)
                Dim idApartamento As Integer = CInt(selectedRow.Cells("IdApartamento").Value)

                Dim apartamentoAEditar As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

                If apartamentoAEditar IsNot Nothing Then
                    Dim formDetallePropietario As New FormDetallePropietario(apartamentoAEditar)
                    Dim resultado As DialogResult = formDetallePropietario.ShowDialog()

                    If resultado = DialogResult.OK Then
                        ' Mostrar mensaje de confirmación
                        MessageBox.Show("Información del propietario actualizada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' Recargar la lista manteniendo el filtro
                        CargarPropietarios(txtBuscador.Text.Trim())

                        ' Mantener la selección en el registro editado si es posible
                        For Each row As DataGridViewRow In dgvPropietarios.Rows
                            If CInt(row.Cells("IdApartamento").Value) = idApartamento Then
                                row.Selected = True
                                dgvPropietarios.CurrentCell = row.Cells(1) ' Seleccionar primera columna visible
                                Exit For
                            End If
                        Next
                    End If
                Else
                    MessageBox.Show("No se pudo cargar la información del apartamento para editar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Seleccione un propietario para editar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al editar propietario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnVolver_Click(sender As Object, e As EventArgs)
        Try
            Me.Close()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    Private Sub FormPropietarios_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Try
            ' Reposicionar el botón VOLVER centrado en el panel inferior
            If btnVolver IsNot Nothing Then
                btnVolver.Location = New Point((Me.Width - btnVolver.Width) \ 2, 10)
            End If

            ' Redimensionar panel de estadísticas para usar todo el ancho disponible
            If lblEstadisticas IsNot Nothing Then
                Dim anchoDisponible As Integer = Me.Width - 250
                If anchoDisponible > 400 Then
                    lblEstadisticas.Width = anchoDisponible
                End If
            End If

        Catch ex As Exception
            ' Ignorar errores de redimensionamiento
        End Try
    End Sub

End Class