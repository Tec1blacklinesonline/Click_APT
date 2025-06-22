Imports System.Windows.Forms
Imports System.Data.SQLite
Imports BCrypt.Net
Imports System.Drawing

Public Class FormCrearUsuario
    Inherits Form

    Private btnCrearUsuario As Button
    Private lblResultado As Label
    Private txtUsuario As TextBox
    Private txtContrasena As TextBox
    Private lblUsuario As Label
    Private lblContrasena As Label

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Crear Usuario de Emergencia"
        Me.Size = New Size(400, 250)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' Etiqueta Usuario
        lblUsuario = New Label() With {
            .Text = "Usuario:",
            .Location = New Point(20, 20),
            .Size = New Size(80, 23),
            .Font = New Font("Segoe UI", 10)
        }
        Me.Controls.Add(lblUsuario)

        ' TextBox Usuario
        txtUsuario = New TextBox() With {
            .Location = New Point(110, 20),
            .Size = New Size(200, 23),
            .Text = "CESAR",
            .Font = New Font("Segoe UI", 10)
        }
        Me.Controls.Add(txtUsuario)

        ' Etiqueta Contraseña
        lblContrasena = New Label() With {
            .Text = "Contraseña:",
            .Location = New Point(20, 60),
            .Size = New Size(80, 23),
            .Font = New Font("Segoe UI", 10)
        }
        Me.Controls.Add(lblContrasena)

        ' TextBox Contraseña
        txtContrasena = New TextBox() With {
            .Location = New Point(110, 60),
            .Size = New Size(200, 23),
            .Text = "72501",
            .Font = New Font("Segoe UI", 10)
        }
        Me.Controls.Add(txtContrasena)

        ' Botón crear usuario
        btnCrearUsuario = New Button() With {
            .Text = "Crear Usuario",
            .Location = New Point(150, 100),
            .Size = New Size(120, 35),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnCrearUsuario.FlatAppearance.BorderSize = 0
        AddHandler btnCrearUsuario.Click, AddressOf btnCrearUsuario_Click
        Me.Controls.Add(btnCrearUsuario)

        ' Label resultado
        lblResultado = New Label() With {
            .Location = New Point(20, 150),
            .Size = New Size(350, 60),
            .Font = New Font("Segoe UI", 9),
            .Text = "Haga clic en 'Crear Usuario' para continuar.",
            .ForeColor = Color.Gray
        }
        Me.Controls.Add(lblResultado)
    End Sub

    Private Sub btnCrearUsuario_Click(sender As Object, e As EventArgs)
        Try
            lblResultado.Text = "Procesando..."
            lblResultado.ForeColor = Color.Blue
            Application.DoEvents()

            ' Validaciones
            If String.IsNullOrWhiteSpace(txtUsuario.Text) Then
                lblResultado.Text = "❌ Ingrese un nombre de usuario."
                lblResultado.ForeColor = Color.Red
                Return
            End If

            If String.IsNullOrWhiteSpace(txtContrasena.Text) Then
                lblResultado.Text = "❌ Ingrese una contraseña."
                lblResultado.ForeColor = Color.Red
                Return
            End If

            ' Crear usuario
            CrearUsuarioEnBD(txtUsuario.Text.Trim(), txtContrasena.Text.Trim())

        Catch ex As Exception
            lblResultado.Text = $"❌ Error: {ex.Message}"
            lblResultado.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub CrearUsuarioEnBD(usuario As String, contrasena As String)
        Try
            ' Usar la misma cadena de conexión de tu App.config
            Dim connectionString As String = "Data Source=C:\Users\DELL\Dropbox\BD_COOPDIASAM\CONJUNTO_2025.db;Version=3;"

            Using conexion As New SQLiteConnection(connectionString)
                conexion.Open()
                lblResultado.Text = "✅ Conectado a la base de datos..."
                lblResultado.ForeColor = Color.Green
                Application.DoEvents()

                ' Verificar si el usuario ya existe
                Dim consultaExiste As String = "SELECT COUNT(*) FROM Usuarios WHERE nombre_usuario = @usuario"
                Using comandoExiste As New SQLiteCommand(consultaExiste, conexion)
                    comandoExiste.Parameters.AddWithValue("@usuario", usuario)
                    Dim existe As Integer = Convert.ToInt32(comandoExiste.ExecuteScalar())

                    If existe > 0 Then
                        ' Usuario existe, actualizarlo
                        ActualizarUsuarioExistente(conexion, usuario, contrasena)
                    Else
                        ' Usuario no existe, crearlo
                        CrearNuevoUsuario(conexion, usuario, contrasena)
                    End If
                End Using
            End Using

        Catch ex As Exception
            lblResultado.Text = $"❌ Error de conexión: {ex.Message}"
            lblResultado.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub CrearNuevoUsuario(conexion As SQLiteConnection, usuario As String, contrasena As String)
        Try
            ' Generar hash BCrypt
            Dim hashContrasena As String = BCrypt.Net.BCrypt.HashPassword(contrasena)

            ' Insertar nuevo usuario
            Dim consultaInsert As String = "INSERT INTO Usuarios (nombre_usuario, contrasena_hash, nombre_completo, email, rol, fecha_creacion, activo) VALUES (@usuario, @contrasena, @nombre, @email, @rol, datetime('now'), 1)"

            Using comando As New SQLiteCommand(consultaInsert, conexion)
                comando.Parameters.AddWithValue("@usuario", usuario)
                comando.Parameters.AddWithValue("@contrasena", hashContrasena)
                comando.Parameters.AddWithValue("@nombre", $"{usuario} Administrador")
                comando.Parameters.AddWithValue("@email", $"{usuario.ToLower()}@coopdiasam.com")
                comando.Parameters.AddWithValue("@rol", "Administrador")

                Dim resultado As Integer = comando.ExecuteNonQuery()

                If resultado > 0 Then
                    lblResultado.Text = $"✅ Usuario '{usuario}' creado exitosamente!" & vbCrLf & $"Contraseña: {contrasena}" & vbCrLf & "Ya puedes hacer login."
                    lblResultado.ForeColor = Color.Green
                Else
                    lblResultado.Text = "❌ No se pudo crear el usuario."
                    lblResultado.ForeColor = Color.Red
                End If
            End Using

        Catch ex As Exception
            lblResultado.Text = $"❌ Error al crear usuario: {ex.Message}"
            lblResultado.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub ActualizarUsuarioExistente(conexion As SQLiteConnection, usuario As String, contrasena As String)
        Try
            ' Generar hash BCrypt
            Dim hashContrasena As String = BCrypt.Net.BCrypt.HashPassword(contrasena)

            ' Actualizar usuario existente
            Dim consultaUpdate As String = "UPDATE Usuarios SET contrasena_hash = @contrasena, activo = 1 WHERE nombre_usuario = @usuario"

            Using comando As New SQLiteCommand(consultaUpdate, conexion)
                comando.Parameters.AddWithValue("@usuario", usuario)
                comando.Parameters.AddWithValue("@contrasena", hashContrasena)

                Dim resultado As Integer = comando.ExecuteNonQuery()

                If resultado > 0 Then
                    lblResultado.Text = $"✅ Usuario '{usuario}' actualizado exitosamente!" & vbCrLf & $"Nueva contraseña: {contrasena}" & vbCrLf & "Ya puedes hacer login."
                    lblResultado.ForeColor = Color.Green
                Else
                    lblResultado.Text = "❌ No se pudo actualizar el usuario."
                    lblResultado.ForeColor = Color.Red
                End If
            End Using

        Catch ex As Exception
            lblResultado.Text = $"❌ Error al actualizar usuario: {ex.Message}"
            lblResultado.ForeColor = Color.Red
        End Try
    End Sub

    ' Método para probar desde tu proyecto principal
    Public Shared Sub MostrarFormularioCrearUsuario()
        Try
            Dim form As New FormCrearUsuario()
            form.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir formulario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class