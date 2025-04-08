Public Class Inicio
    Private panelLogin As Panel
    Private txtUsuario As TextBox
    Private txtPassword As TextBox

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configurar el formulario
        Me.Text = "Login COOPDIASAM"
        Me.Width = 350
        Me.Height = 250
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Crear un panel para contener los elementos del login
        panelLogin = New Panel()
        panelLogin.Width = 250
        panelLogin.Height = 150
        panelLogin.Left = (Me.ClientSize.Width - panelLogin.Width) / 2
        panelLogin.Top = (Me.ClientSize.Height - panelLogin.Height) / 2

        ' Crear etiqueta de usuario
        Dim lblUsuario As New Label()
        lblUsuario.Text = "Usuario:"
        lblUsuario.Location = New Point(10, 20)
        lblUsuario.Width = 80

        ' Caja de texto para usuario
        txtUsuario = New TextBox()
        txtUsuario.Location = New Point(90, 20)
        txtUsuario.Width = 150

        ' Etiqueta de contraseña
        Dim lblPassword As New Label()
        lblPassword.Text = "Contraseña:"
        lblPassword.Location = New Point(10, 60)
        lblPassword.Width = 80

        ' Caja de texto para contraseña
        txtPassword = New TextBox()
        txtPassword.Location = New Point(90, 60)
        txtPassword.Width = 150
        txtPassword.UseSystemPasswordChar = True ' Reemplaza a PasswordChar

        ' Botón de login
        Dim btnLogin As New Button()
        btnLogin.Text = "Iniciar Sesión"
        btnLogin.Location = New Point(90, 100)
        btnLogin.Width = 120
        AddHandler btnLogin.Click, AddressOf btnLogin_Click

        ' Agregar controles al panel
        panelLogin.Controls.Add(lblUsuario)
        panelLogin.Controls.Add(txtUsuario)
        panelLogin.Controls.Add(lblPassword)
        panelLogin.Controls.Add(txtPassword)
        panelLogin.Controls.Add(btnLogin)

        ' Agregar panel al formulario
        Me.Controls.Add(panelLogin)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs)
        ' Validaciones simples
        If txtUsuario.Text.Trim() = "" Then
            MessageBox.Show("Por favor ingrese un nombre de usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUsuario.Focus()
            Return
        End If

        If txtPassword.Text.Trim() = "" Then
            MessageBox.Show("Por favor ingrese una contraseña.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPassword.Focus()
            Return
        End If

        Try
            If ConexionBD.ValidarUsuario(txtUsuario.Text, txtPassword.Text) Then
                MessageBox.Show("Inicio de sesión exitoso", "Bienvenido", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Dim formAdmin As New FormAdministracion()
                Me.Hide()
            Else
                MessageBox.Show("Usuario o contraseña incorrectos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPassword.Clear()
                txtPassword.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show("Error al intentar iniciar sesión: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
