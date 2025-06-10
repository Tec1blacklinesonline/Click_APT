Imports System.Windows.Forms
Imports System.Drawing
Imports BCrypt.Net
Imports System.Data.SQLite

Public Class FormRegistro
    Inherits Form

    Private tabControl As TabControl
    Private btnGuardar As Button
    Private btnCancelar As Button

    ' Controles para registro de usuarios
    Private txtNombreUsuario As TextBox
    Private txtNombreCompleto As TextBox
    Private txtEmail As TextBox
    Private txtContrasena As TextBox
    Private txtConfirmarContrasena As TextBox
    Private cboRol As ComboBox

    ' Controles para registro de cuotas
    Private cboAsamblea As ComboBox
    Private cboTipoPiso As ComboBox
    Private txtValorCuota As TextBox
    Private dtpFechaVencimiento As DateTimePicker
    Private txtDescripcionCuota As TextBox
    Private chkGenerarParaTodos As CheckBox

    Private Sub FormRegistro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarDatosIniciales()
    End Sub

    Private Sub ConfigurarFormulario()
        Me.Text = "Registro de Usuarios y Cuotas"
        Me.Size = New Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' Panel superior
        Dim panelSuperior As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(231, 76, 60)
        }

        Dim lblTitulo As New Label With {
            .Text = "📝 REGISTRO DE USUARIOS Y CUOTAS",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(20, 25),
            .AutoSize = True
        }
        panelSuperior.Controls.Add(lblTitulo)
        Me.Controls.Add(panelSuperior)

        ' Panel inferior con botones
        Dim panelInferior As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 60,
            .BackColor = Color.FromArgb(236, 240, 241)
        }

        btnGuardar = New Button With {
            .Text = "💾 Guardar",
            .Size = New Size(120, 35),
            .Location = New Point(350, 12),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnGuardar.FlatAppearance.BorderSize = 0
        AddHandler btnGuardar.Click, AddressOf btnGuardar_Click
        panelInferior.Controls.Add(btnGuardar)

        btnCancelar = New Button With {
            .Text = "❌ Cancelar",
            .Size = New Size(120, 35),
            .Location = New Point(480, 12),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
        panelInferior.Controls.Add(btnCancelar)

        Me.Controls.Add(panelInferior)

        ' TabControl
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        ' Tab Registro de Usuarios
        Dim tabUsuarios As New TabPage("👤 Registrar Usuario")
        CrearTabUsuarios(tabUsuarios)

        ' Tab Registro de Cuotas
        Dim tabCuotas As New TabPage("💰 Generar Cuotas")
        CrearTabCuotas(tabCuotas)

        tabControl.TabPages.Add(tabUsuarios)
        tabControl.TabPages.Add(tabCuotas)

        Me.Controls.Add(tabControl)
    End Sub

    Private Sub CrearTabUsuarios(tab As TabPage)
        tab.BackColor = Color.White

        ' Nombre de usuario
        Dim lblNombreUsuario As New Label With {
            .Text = "Nombre de Usuario:",
            .Location = New Point(30, 30),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblNombreUsuario)

        txtNombreUsuario = New TextBox With {
            .Location = New Point(180, 27),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(txtNombreUsuario)

        ' Nombre completo
        Dim lblNombreCompleto As New Label With {
            .Text = "Nombre Completo:",
            .Location = New Point(30, 70),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblNombreCompleto)

        txtNombreCompleto = New TextBox With {
            .Location = New Point(180, 67),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(txtNombreCompleto)

        ' Email
        Dim lblEmail As New Label With {
            .Text = "Correo Electrónico:",
            .Location = New Point(30, 110),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblEmail)

        txtEmail = New TextBox With {
            .Location = New Point(180, 107),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(txtEmail)

        ' Rol
        Dim lblRol As New Label With {
            .Text = "Rol:",
            .Location = New Point(30, 150),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblRol)

        cboRol = New ComboBox With {
            .Location = New Point(180, 147),
            .Size = New Size(200, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 10)
        }
        cboRol.Items.AddRange({"Administrador", "Operador", "Consulta"})
        cboRol.SelectedIndex = 1 ' Operador por defecto
        tab.Controls.Add(cboRol)

        ' Contraseña
        Dim lblContrasena As New Label With {
            .Text = "Contraseña:",
            .Location = New Point(30, 190),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblContrasena)

        txtContrasena = New TextBox With {
            .Location = New Point(180, 187),
            .Size = New Size(300, 25),
            .PasswordChar = "*"c,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(txtContrasena)

        ' Confirmar contraseña
        Dim lblConfirmar As New Label With {
            .Text = "Confirmar Contraseña:",
            .Location = New Point(30, 230),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblConfirmar)

        txtConfirmarContrasena = New TextBox With {
            .Location = New Point(180, 227),
            .Size = New Size(300, 25),
            .PasswordChar = "*"c,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(txtConfirmarContrasena)

        ' Nota informativa
        Dim lblNota As New Label With {
            .Text = "Nota: La contraseña debe tener al menos 8 caracteres.",
            .Location = New Point(30, 270),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.Gray
        }
        tab.Controls.Add(lblNota)
    End Sub

    Private Sub CrearTabCuotas(tab As TabPage)
        tab.BackColor = Color.White

        ' Asamblea
        Dim lblAsamblea As New Label With {
            .Text = "Asamblea:",
            .Location = New Point(30, 30),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblAsamblea)

        cboAsamblea = New ComboBox With {
            .Location = New Point(180, 27),
            .Size = New Size(300, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(cboAsamblea)

        ' Tipo de piso
        Dim lblTipoPiso As New Label With {
            .Text = "Tipo de Piso:",
            .Location = New Point(30, 70),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblTipoPiso)

        cboTipoPiso = New ComboBox With {
            .Location = New Point(180, 67),
            .Size = New Size(200, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 10)
        }
        cboTipoPiso.Items.AddRange({"Todos", "Primer Piso", "Segundo Piso", "Tercer Piso", "Cuarto Piso", "Quinto Piso"})
        cboTipoPiso.SelectedIndex = 0
        tab.Controls.Add(cboTipoPiso)

        ' Valor de la cuota
        Dim lblValor As New Label With {
            .Text = "Valor de la Cuota:",
            .Location = New Point(30, 110),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblValor)

        txtValorCuota = New TextBox With {
            .Location = New Point(180, 107),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10),
            .TextAlign = HorizontalAlignment.Right
        }
        tab.Controls.Add(txtValorCuota)

        Dim lblPesos As New Label With {
            .Text = "COP",
            .Location = New Point(340, 110),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblPesos)

        ' Fecha de vencimiento
        Dim lblFechaVenc As New Label With {
            .Text = "Fecha de Vencimiento:",
            .Location = New Point(30, 150),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblFechaVenc)

        dtpFechaVencimiento = New DateTimePicker With {
            .Location = New Point(180, 147),
            .Size = New Size(200, 25),
            .Format = DateTimePickerFormat.Short,
            .Font = New Font("Segoe UI", 10),
            .Value = DateTime.Now.AddDays(30)
        }
        tab.Controls.Add(dtpFechaVencimiento)

        ' Descripción
        Dim lblDescripcion As New Label With {
            .Text = "Descripción:",
            .Location = New Point(30, 190),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(lblDescripcion)

        txtDescripcionCuota = New TextBox With {
            .Location = New Point(180, 187),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        tab.Controls.Add(txtDescripcionCuota)

        ' Checkbox generar para todos
        chkGenerarParaTodos = New CheckBox With {
            .Text = "Generar cuota para todos los apartamentos",
            .Location = New Point(30, 230),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10),
            .Checked = True
        }
        tab.Controls.Add(chkGenerarParaTodos)

        ' Nota informativa
        Dim lblNotaCuotas As New Label With {
            .Text = "Nota: Se generarán cuotas individuales para cada apartamento según el tipo de piso seleccionado.",
            .Location = New Point(30, 270),
            .Size = New Size(500, 40),
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.Gray
        }
        tab.Controls.Add(lblNotaCuotas)
    End Sub

    Private Sub CargarDatosIniciales()
        Try
            ' Cargar asambleas
            Dim asambleas = AsambleasDAL.ObtenerTodasLasAsambleas()
            cboAsamblea.Items.Clear()
            For Each asamblea In asambleas
                cboAsamblea.Items.Add($"{asamblea.NombreAsamblea} ({asamblea.FechaAsamblea:dd/MM/yyyy})")
            Next

            If cboAsamblea.Items.Count > 0 Then
                cboAsamblea.SelectedIndex = 0
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al cargar datos iniciales: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)
        Try
            If tabControl.SelectedTab.Text.Contains("Usuario") Then
                GuardarUsuario()
            Else
                GenerarCuotas()
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GuardarUsuario()
        ' Validaciones
        If String.IsNullOrWhiteSpace(txtNombreUsuario.Text) Then
            MessageBox.Show("El nombre de usuario es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreUsuario.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtNombreCompleto.Text) Then
            MessageBox.Show("El nombre completo es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreCompleto.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtEmail.Text) OrElse Not txtEmail.Text.Contains("@") Then
            MessageBox.Show("Debe ingresar un correo electrónico válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtEmail.Focus()
            Return
        End If

        If txtContrasena.Text.Length < 8 Then
            MessageBox.Show("La contraseña debe tener al menos 8 caracteres.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContrasena.Focus()
            Return
        End If

        If txtContrasena.Text <> txtConfirmarContrasena.Text Then
            MessageBox.Show("Las contraseñas no coinciden.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtConfirmarContrasena.Focus()
            Return
        End If

        Try
            ' Encriptar contraseña
            Dim hashContrasena As String = BCrypt.Net.BCrypt.HashPassword(txtContrasena.Text)

            ' Insertar usuario en base de datos
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Verificar que no exista el usuario
                Dim consultaExiste As String = "SELECT COUNT(*) FROM Usuarios WHERE nombre_usuario = @usuario"
                Using comandoExiste As New SQLiteCommand(consultaExiste, conexion)
                    comandoExiste.Parameters.AddWithValue("@usuario", txtNombreUsuario.Text.Trim())
                    If Convert.ToInt32(comandoExiste.ExecuteScalar()) > 0 Then
                        MessageBox.Show("Ya existe un usuario con ese nombre.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using

                ' Insertar nuevo usuario
                Dim consultaInsert As String = "INSERT INTO Usuarios (nombre_usuario, contrasena_hash, nombre_completo, email, rol, fecha_creacion, activo) VALUES (@usuario, @contrasena, @nombre, @email, @rol, datetime('now'), 1)"
                Using comando As New SQLiteCommand(consultaInsert, conexion)
                    comando.Parameters.AddWithValue("@usuario", txtNombreUsuario.Text.Trim())
                    comando.Parameters.AddWithValue("@contrasena", hashContrasena)
                    comando.Parameters.AddWithValue("@nombre", txtNombreCompleto.Text.Trim())
                    comando.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
                    comando.Parameters.AddWithValue("@rol", cboRol.SelectedItem.ToString())

                    If comando.ExecuteNonQuery() > 0 Then
                        MessageBox.Show("Usuario registrado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LimpiarFormularioUsuario()
                    Else
                        MessageBox.Show("No se pudo registrar el usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al registrar usuario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GenerarCuotas()
        ' Validaciones
        If cboAsamblea.SelectedIndex = -1 Then
            MessageBox.Show("Debe seleccionar una asamblea.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cboAsamblea.Focus()
            Return
        End If

        Dim valorCuota As Decimal
        If Not Decimal.TryParse(txtValorCuota.Text, valorCuota) OrElse valorCuota <= 0 Then
            MessageBox.Show("Debe ingresar un valor de cuota válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtValorCuota.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtDescripcionCuota.Text) Then
            MessageBox.Show("La descripción de la cuota es obligatoria.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDescripcionCuota.Focus()
            Return
        End If

        Try
            Dim asambleas = AsambleasDAL.ObtenerTodasLasAsambleas()
            Dim asambleaSeleccionada = asambleas(cboAsamblea.SelectedIndex)
            Dim idAsamblea As Integer = asambleaSeleccionada.IdAsamblea

            ' Generar cuotas
            Dim apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()
            Dim cuotasGeneradas As Integer = 0

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        For Each apartamento In apartamentos
                            ' Filtrar por tipo de piso si no es "Todos"
                            If cboTipoPiso.SelectedIndex > 0 Then
                                Dim pisoSeleccionado As Integer = cboTipoPiso.SelectedIndex
                                If apartamento.Piso <> pisoSeleccionado Then
                                    Continue For
                                End If
                            End If

                            ' Insertar cuota generada para el apartamento
                            Dim consultaInsert As String = "INSERT INTO cuotas_generadas_apartamento (id_apartamentos, matricula_inmobiliaria, fecha_cuota, valor_cuota, fecha_vencimiento, estado, tipo_cuota, tipo_piso, id_asamblea) VALUES (@idApt, @matricula, date('now'), @valor, @fechaVenc, 'pendiente', @descripcion, @tipoPiso, @idAsamblea)"

                            Using comando As New SQLiteCommand(consultaInsert, conexion, transaccion)
                                comando.Parameters.AddWithValue("@idApt", apartamento.IdApartamento)
                                comando.Parameters.AddWithValue("@matricula", apartamento.MatriculaInmobiliaria)
                                comando.Parameters.AddWithValue("@valor", valorCuota)
                                comando.Parameters.AddWithValue("@fechaVenc", dtpFechaVencimiento.Value.Date)
                                comando.Parameters.AddWithValue("@descripcion", txtDescripcionCuota.Text.Trim())
                                comando.Parameters.AddWithValue("@tipoPiso", cboTipoPiso.SelectedItem.ToString())
                                comando.Parameters.AddWithValue("@idAsamblea", idAsamblea)

                                If comando.ExecuteNonQuery() > 0 Then
                                    cuotasGeneradas += 1
                                End If
                            End Using
                        Next

                        transaccion.Commit()
                        MessageBox.Show($"Se generaron {cuotasGeneradas} cuotas exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LimpiarFormularioCuotas()

                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al generar cuotas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LimpiarFormularioUsuario()
        txtNombreUsuario.Clear()
        txtNombreCompleto.Clear()
        txtEmail.Clear()
        txtContrasena.Clear()
        txtConfirmarContrasena.Clear()
        cboRol.SelectedIndex = 1
        txtNombreUsuario.Focus()
    End Sub

    Private Sub LimpiarFormularioCuotas()
        txtValorCuota.Clear()
        txtDescripcionCuota.Clear()
        dtpFechaVencimiento.Value = DateTime.Now.AddDays(30)
        cboTipoPiso.SelectedIndex = 0
        chkGenerarParaTodos.Checked = True
        txtValorCuota.Focus()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

End Class