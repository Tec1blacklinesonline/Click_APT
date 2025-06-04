Imports System.Windows.Forms
Imports System.Drawing

Public Class FormConfiguracionSMTP
    Inherits Form

    Private txtServidor As TextBox
    Private txtPuerto As TextBox
    Private txtUsuario As TextBox
    Private txtContrasena As TextBox
    Private chkSSL As CheckBox
    Private txtCorreoRemitente As TextBox
    Private txtNombreRemitente As TextBox
    Private txtTimeout As TextBox
    Private cboProveedores As ComboBox
    Private btnGuardar As Button
    Private btnCancelar As Button
    Private btnProbar As Button

    Public Sub New()
        InitializeComponent()
        CargarConfiguracion()
    End Sub

    Private Sub InitializeComponent()
        ' Configuración del formulario
        Me.Text = "Configuración de Correo Electrónico"
        Me.Size = New Size(500, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' Panel principal
        Dim panelPrincipal As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .Padding = New Padding(20)
        }
        Me.Controls.Add(panelPrincipal)

        ' Título
        Dim lblTitulo As New Label With {
            .Text = "CONFIGURACIÓN SMTP",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(52, 73, 94),
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblTitulo)

        ' ComboBox Proveedores
        Dim lblProveedor As New Label With {
            .Text = "Proveedor de correo:",
            .Location = New Point(20, 60),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblProveedor)

        cboProveedores = New ComboBox With {
            .Location = New Point(20, 80),
            .Size = New Size(200, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboProveedores.Items.AddRange({"Personalizado", "Gmail", "Yahoo", "Outlook", "Office365"})
        cboProveedores.SelectedIndex = 0
        AddHandler cboProveedores.SelectedIndexChanged, AddressOf cboProveedores_SelectedIndexChanged
        panelPrincipal.Controls.Add(cboProveedores)

        ' Servidor SMTP
        Dim lblServidor As New Label With {
            .Text = "Servidor SMTP:",
            .Location = New Point(20, 120),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblServidor)

        txtServidor = New TextBox With {
            .Location = New Point(20, 140),
            .Size = New Size(250, 25)
        }
        panelPrincipal.Controls.Add(txtServidor)

        ' Puerto
        Dim lblPuerto As New Label With {
            .Text = "Puerto:",
            .Location = New Point(280, 120),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblPuerto)

        txtPuerto = New TextBox With {
            .Location = New Point(280, 140),
            .Size = New Size(80, 25)
        }
        panelPrincipal.Controls.Add(txtPuerto)

        ' SSL
        chkSSL = New CheckBox With {
            .Text = "Usar SSL/TLS",
            .Location = New Point(370, 140),
            .AutoSize = True,
            .Checked = True
        }
        panelPrincipal.Controls.Add(chkSSL)

        ' Usuario
        Dim lblUsuario As New Label With {
            .Text = "Usuario (correo electrónico):",
            .Location = New Point(20, 180),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblUsuario)

        txtUsuario = New TextBox With {
            .Location = New Point(20, 200),
            .Size = New Size(440, 25)
        }
        panelPrincipal.Controls.Add(txtUsuario)

        ' Contraseña
        Dim lblContrasena As New Label With {
            .Text = "Contraseña:",
            .Location = New Point(20, 240),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblContrasena)

        txtContrasena = New TextBox With {
            .Location = New Point(20, 260),
            .Size = New Size(440, 25),
            .PasswordChar = "*"c
        }
        panelPrincipal.Controls.Add(txtContrasena)

        ' Correo remitente
        Dim lblRemitente As New Label With {
            .Text = "Correo del remitente:",
            .Location = New Point(20, 300),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblRemitente)

        txtCorreoRemitente = New TextBox With {
            .Location = New Point(20, 320),
            .Size = New Size(440, 25)
        }
        panelPrincipal.Controls.Add(txtCorreoRemitente)

        ' Nombre remitente
        Dim lblNombre As New Label With {
            .Text = "Nombre del remitente:",
            .Location = New Point(20, 360),
            .AutoSize = True
        }
        panelPrincipal.Controls.Add(lblNombre)

        txtNombreRemitente = New TextBox With {
            .Location = New Point(20, 380),
            .Size = New Size(440, 25)
        }
        panelPrincipal.Controls.Add(txtNombreRemitente)

        ' Panel de botones
        Dim panelBotones As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 60,
            .BackColor = Color.FromArgb(240, 240, 240)
        }
        Me.Controls.Add(panelBotones)

        ' Botón Probar
        btnProbar = New Button With {
            .Text = "Probar Conexión",
            .Size = New Size(120, 35),
            .Location = New Point(20, 12),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnProbar.FlatAppearance.BorderSize = 0
        AddHandler btnProbar.Click, AddressOf btnProbar_Click
        panelBotones.Controls.Add(btnProbar)

        ' Botón Guardar
        btnGuardar = New Button With {
            .Text = "Guardar",
            .Size = New Size(100, 35),
            .Location = New Point(250, 12),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnGuardar.FlatAppearance.BorderSize = 0
        AddHandler btnGuardar.Click, AddressOf btnGuardar_Click
        panelBotones.Controls.Add(btnGuardar)

        ' Botón Cancelar
        btnCancelar = New Button With {
            .Text = "Cancelar",
            .Size = New Size(100, 35),
            .Location = New Point(360, 12),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
        panelBotones.Controls.Add(btnCancelar)
    End Sub

    Private Sub CargarConfiguracion()
        Try
            Dim config As ConfiguracionSMTP = ConfiguracionSMTP.Cargar()

            txtServidor.Text = config.ServidorSMTP
            txtPuerto.Text = config.Puerto.ToString()
            txtUsuario.Text = config.UsuarioSMTP
            txtContrasena.Text = config.ContrasenaSMTP
            chkSSL.Checked = config.UsarSSL
            txtCorreoRemitente.Text = config.CorreoRemitente
            txtNombreRemitente.Text = config.NombreRemitente

        Catch ex As Exception
            MessageBox.Show($"Error al cargar configuración: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboProveedores_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cboProveedores.SelectedIndex = 0 Then Return ' Personalizado

        Dim configs = ConfiguracionSMTP.ObtenerConfiguracionesPredefinidas()
        Dim nombreProveedor As String = cboProveedores.SelectedItem.ToString()

        If configs.ContainsKey(nombreProveedor) Then
            Dim config = configs(nombreProveedor)
            txtServidor.Text = config.ServidorSMTP
            txtPuerto.Text = config.Puerto.ToString()
            chkSSL.Checked = config.UsarSSL

            ' Para Gmail, mostrar nota sobre contraseña de aplicación
            If nombreProveedor = "Gmail" Then
                MessageBox.Show("Para Gmail, debe usar una 'Contraseña de aplicación' en lugar de su contraseña normal." & vbCrLf & vbCrLf