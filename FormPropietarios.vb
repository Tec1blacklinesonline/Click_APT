Imports System.Windows.Forms
Imports System.Drawing

Public Class FormPropietarios
    Inherits Form

    Private dgvPropietarios As DataGridView
    Private lblTitulo As Label
    Private lblBuscador As Label
    Private txtBuscador As TextBox
    Private btnLimpiar As Button
    Private btnAgregar As Button
    Private btnEditar As Button
    Private btnVolver As Button
    Private listaCompletaApartamentos As List(Of Apartamento)

    Private Sub FormPropietarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Gestión de Propietarios"
        Me.Size = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ConfigurarFormulario()
        CargarPropietarios()
    End Sub

    Private Sub ConfigurarFormulario()
        ' Título del formulario
        lblTitulo = New Label() With {
            .Text = "Listado y Gestión de Propietarios",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.FromArgb(41, 128, 185),
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        Me.Controls.Add(lblTitulo)

        ' Label del buscador
        lblBuscador = New Label() With {
            .Text = "Buscar por Apartamento (ej: 1202 para Torre 1, Apt 202):",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .ForeColor = Color.FromArgb(52, 73, 94),
            .AutoSize = True,
            .Location = New Point(20, 60)
        }
        Me.Controls.Add(lblBuscador)

        ' TextBox del buscador
        txtBuscador = New TextBox() With {
            .Location = New Point(20, 85),
            .Size = New Size(250, 25),
            .Font = New Font("Segoe UI", 10),
            .BorderStyle = BorderStyle.FixedSingle
        }
        AddHandler txtBuscador.TextChanged, AddressOf txtBuscador_TextChanged
        Me.Controls.Add(txtBuscador)

        ' Botón Limpiar búsqueda
        btnLimpiar = New Button() With {
            .Text = "Limpiar",
            .Size = New Size(80, 25),
            .Location = New Point(280, 85),
            .BackColor = Color.FromArgb(149, 165, 166),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .FlatStyle = FlatStyle.Flat
        }
        btnLimpiar.FlatAppearance.BorderSize = 0
        AddHandler btnLimpiar.Click, AddressOf btnLimpiar_Click
        Me.Controls.Add(btnLimpiar)

        ' DataGridView para mostrar los apartamentos/propietarios
        dgvPropietarios = New DataGridView()
        With dgvPropietarios
            .Location = New Point(20, 125)
            .Size = New Size(Me.ClientSize.Width - 40, Me.ClientSize.Height - 235)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .BackgroundColor = Color.White
            .BorderStyle = BorderStyle.None
            .ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(41, 128, 185)
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            .EnableHeadersVisualStyles = False
            .ColumnHeadersHeight = 35
        End With
        Me.Controls.Add(dgvPropietarios)

        ' Botón Agregar
        btnAgregar = New Button() With {
            .Text = "Agregar Nuevo",
            .Size = New Size(120, 40),
            .Location = New Point(20, dgvPropietarios.Bottom + 20),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        btnAgregar.FlatAppearance.BorderSize = 0
        AddHandler btnAgregar.Click, AddressOf btnAgregar_Click
        Me.Controls.Add(btnAgregar)

        ' Botón Editar
        btnEditar = New Button() With {
            .Text = "Editar",
            .Size = New Size(100, 40),
            .Location = New Point(150, dgvPropietarios.Bottom + 20),
            .BackColor = Color.FromArgb(243, 156, 18),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        btnEditar.FlatAppearance.BorderSize = 0
        AddHandler btnEditar.Click, AddressOf btnEditar_Click
        Me.Controls.Add(btnEditar)

        ' Botón Volver
        btnVolver = New Button() With {
            .Text = "Volver",
            .Size = New Size(100, 40),
            .Location = New Point(260, dgvPropietarios.Bottom + 20),
            .BackColor = Color.FromArgb(52, 73, 94),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        btnVolver.FlatAppearance.BorderSize = 0
        AddHandler btnVolver.Click, AddressOf btnVolver_Click
        Me.Controls.Add(btnVolver)
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

                ' Configurar visibilidad y nombres de columnas con anchos específicos
                If dgvPropietarios.Columns.Contains("IdApartamento") Then dgvPropietarios.Columns("IdApartamento").Visible = False

                If dgvPropietarios.Columns.Contains("Torre") Then
                    dgvPropietarios.Columns("Torre").HeaderText = "Torre"
                    dgvPropietarios.Columns("Torre").Width = 60
                End If

                If dgvPropietarios.Columns.Contains("Piso") Then
                    dgvPropietarios.Columns("Piso").HeaderText = "Piso"
                    dgvPropietarios.Columns("Piso").Width = 60
                End If

                If dgvPropietarios.Columns.Contains("NumeroApartamento") Then
                    dgvPropietarios.Columns("NumeroApartamento").HeaderText = "Apartamento"
                    dgvPropietarios.Columns("NumeroApartamento").Width = 90
                End If

                If dgvPropietarios.Columns.Contains("NombreResidente") Then
                    dgvPropietarios.Columns("NombreResidente").HeaderText = "Nombre Residente"
                    dgvPropietarios.Columns("NombreResidente").Width = 200
                End If

                If dgvPropietarios.Columns.Contains("Correo") Then
                    dgvPropietarios.Columns("Correo").HeaderText = "Correo Electrónico"
                    dgvPropietarios.Columns("Correo").Width = 220
                End If

                If dgvPropietarios.Columns.Contains("Telefono") Then
                    dgvPropietarios.Columns("Telefono").HeaderText = "Teléfono"
                    dgvPropietarios.Columns("Telefono").Width = 100
                End If

                If dgvPropietarios.Columns.Contains("MatriculaInmobiliaria") Then
                    dgvPropietarios.Columns("MatriculaInmobiliaria").HeaderText = "Matrícula Inmobiliaria"
                    dgvPropietarios.Columns("MatriculaInmobiliaria").Width = 150
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

                ' Hacer que la última columna visible se ajuste al espacio restante
                For i As Integer = dgvPropietarios.Columns.Count - 1 To 0 Step -1
                    If dgvPropietarios.Columns(i).Visible Then
                        dgvPropietarios.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al cargar la lista de propietarios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub btnAgregar_Click(sender As Object, e As EventArgs)
        Try
            Dim formDetallePropietario As New FormDetallePropietario()
            If formDetallePropietario.ShowDialog() = DialogResult.OK Then
                CargarPropietarios(txtBuscador.Text.Trim())
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al abrir formulario de agregar propietario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
        Me.Close()
    End Sub

End Class