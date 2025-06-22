' Software ClikApt 20250603

Public Class Inicio
    Private WithEvents btnLogin As New Button()
    Private WithEvents btnSalir As New Button() ' Nuevo botón Salir
    Private txtUsuario As New TextBox()
    Private txtPassword As New TextBox()
    Private lblEstadoConexion As New Label()


    Private Sub Inicio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configuración del formulario
        Me.Text = "Login COOPDIASAM"
        Me.ClientSize = New Size(350, 350) ' Aumenté la altura para el nuevo botón
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Panel contenedor
        Dim panelLogin As New Panel With {
            .Size = New Size(250, 220), ' Aumenté la altura para el nuevo botón
            .Location = New Point((Me.ClientSize.Width - 250) \ 2, 40)
        }

        ' Etiqueta usuario
        Dim lblUsuario As New Label With {
            .Text = "Usuario:", .Location = New Point(10, 20), .Width = 80
        }

        txtUsuario = New TextBox With {
            .Location = New Point(90, 20), .Width = 150
        }

        ' Etiqueta contraseña
        Dim lblPassword As New Label With {
            .Text = "Contraseña:", .Location = New Point(10, 60), .Width = 80
        }

        txtPassword = New TextBox With {
            .Location = New Point(90, 60), .Width = 150, .PasswordChar = "*"c
        }

        ' Botón login
        btnLogin = New Button With {
            .Text = "Iniciar Sesión",
            .Location = New Point(90, 100),
            .Width = 120
        }

        ' Nuevo botón Salir
        btnSalir = New Button With {
            .Text = "Salir",
            .Location = New Point(90, 140),
            .Width = 120,
            .BackColor = Color.LightCoral,
            .FlatStyle = FlatStyle.Flat
        }
        btnSalir.FlatAppearance.BorderSize = 0

        ' Etiqueta de estado de conexión
        lblEstadoConexion = New Label With {
            .AutoSize = True,
            .ForeColor = Color.Red,
            .Location = New Point((Me.ClientSize.Width - 200) \ 2, 270), ' Ajusté la posición
            .TextAlign = ContentAlignment.MiddleCenter
        }

        panelLogin.Controls.AddRange({lblUsuario, txtUsuario, lblPassword, txtPassword, btnLogin, btnSalir})
        Me.Controls.Add(panelLogin)
        Me.Controls.Add(lblEstadoConexion)

        ' Probar conexión al cargar
        If ConexionBD.ProbarConexion() Then
            lblEstadoConexion.ForeColor = Color.Green
            lblEstadoConexion.Text = "✔ Conectado a la base de datos"
        Else
            lblEstadoConexion.ForeColor = Color.Red
            lblEstadoConexion.Text = "✖ No se pudo conectar a la base de datos"
            btnLogin.Enabled = False
        End If
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        ' Validar campos vacíos
        If String.IsNullOrWhiteSpace(txtUsuario.Text) Then
            MessageBox.Show("Ingrese nombre de usuario", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUsuario.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show("Ingrese contraseña", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPassword.Focus()
            Return
        End If

        ' Intentar iniciar sesión
        Try
            If ConexionBD.ValidarUsuario(txtUsuario.Text, txtPassword.Text) Then
                ' SOLUCIÓN: Establecer el usuario actual en la sesión
                ConexionBD.EstablecerUsuarioActual(txtUsuario.Text)

                ' Verificar que se estableció correctamente
                Dim usuarioEstablecido = ConexionBD.ObtenerUsuarioActual()
                If usuarioEstablecido IsNot Nothing Then
                    ' Inicio de sesión exitoso, abrir COOPDIASAM
                    Me.Hide()
                    Dim formPrincipal As New COOPDIASAM()
                    formPrincipal.Show()
                Else
                    MessageBox.Show("Error al establecer la sesión del usuario", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Credenciales incorrectas", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPassword.SelectAll()
                txtPassword.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show("Error al iniciar sesión: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Evento para el nuevo botón Salir
    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Application.ExitThread() ' Cierra todos los hilos y la aplicación inmediatamente
    End Sub



End Class