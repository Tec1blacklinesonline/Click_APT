' ============================================================================
' FORMULARIO DE CONFIGURACIÓN SMTP
' Permite configurar dinámicamente los parámetros de correo
' ============================================================================

Imports System.Configuration
Imports System.Drawing
Imports System.Windows.Forms

Public Class FormConfiguracionSMTP
    Inherits Form

    ' Controles del formulario
    Private txtServidor As TextBox
    Private txtPuerto As TextBox
    Private txtUsuario As TextBox
    Private txtContrasena As TextBox
    Private chkSSL As CheckBox
    Private txtRemitente As TextBox
    Private txtNombreRemitente As TextBox
    Private btnProbarConexion As Button
    Private btnGuardar As Button
    Private btnCancelar As Button
    Private lblEstado As Label

    Private Sub FormConfiguracionSMTP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarConfiguracionActual()
    End Sub

    Private Sub ConfigurarFormulario()
        ' Configuración del formulario
        Me.Text = "Configuración de Correo SMTP"
        Me.Size = New Size(500, 480)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' Panel superior con título
        Dim panelTitulo As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.FromArgb(52, 73, 94)
        }

        Dim lblTitulo As New Label With {
            .Text = "📧 CONFIGURACIÓN DE CORREO SMTP",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.White,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Fill
        }
        panelTitulo.Controls.Add(lblTitulo)
        Me.Controls.Add(panelTitulo)

        ' Panel principal
        Dim panelPrincipal As New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(20)
        }

        Dim yPos As Integer = 20

        ' Servidor SMTP
        CrearCampo(panelPrincipal, "Servidor SMTP:", txtServidor, yPos, "smtp.gmail.com")
        yPos += 40

        ' Puerto
        CrearCampo(panelPrincipal, "Puerto:", txtPuerto, yPos, "587")
        yPos += 40

        ' Usuario
        CrearCampo(panelPrincipal, "Usuario (Email):", txtUsuario, yPos, "")
        yPos += 40

        ' Contraseña
        CrearCampo(panelPrincipal, "Contraseña:", txtContrasena, yPos, "", True)
        yPos += 40

        ' SSL/TLS
        Dim lblSSL As New Label With {
            .Text = "Usar SSL/TLS:",
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25),
            .Font = New Font("Segoe UI", 10)
        }
        panelPrincipal.Controls.Add(lblSSL)

        chkSSL = New CheckBox With {
            .Location = New Point(150, yPos),
            .Size = New Size(200, 25),
            .Text = "Habilitar conexión segura",
            .Checked = True,
            .Font = New Font("Segoe UI", 9)
        }
        panelPrincipal.Controls.Add(chkSSL)
        yPos += 40

        ' Nombre del remitente
        CrearCampo(panelPrincipal, "Nombre Remitente:", txtNombreRemitente, yPos, "Administración COOPDIASAM")
        yPos += 40

        ' Correo del remitente
        CrearCampo(panelPrincipal, "Correo Remitente:", txtRemitente, yPos, "")
        yPos += 50

        ' Estado
        lblEstado = New Label With {
            .Location = New Point(20, yPos),
            .Size = New Size(400, 40),
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.Gray,
            .Text = "💡 Complete los datos y pruebe la conexión antes de guardar"
        }
        panelPrincipal.Controls.Add(lblEstado)
        yPos += 50

        ' Botones
        btnProbarConexion = New Button With {
            .Text = "🔍 Probar Conexión",
            .Location = New Point(20, yPos),
            .Size = New Size(140, 35),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnProbarConexion.FlatAppearance.BorderSize = 0
        AddHandler btnProbarConexion.Click, AddressOf btnProbarConexion_Click
        panelPrincipal.Controls.Add(btnProbarConexion)

        btnGuardar = New Button With {
            .Text = "💾 Guardar",
            .Location = New Point(180, yPos),
            .Size = New Size(100, 35),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnGuardar.FlatAppearance.BorderSize = 0
        AddHandler btnGuardar.Click, AddressOf btnGuardar_Click
        panelPrincipal.Controls.Add(btnGuardar)

        btnCancelar = New Button With {
            .Text = "❌ Cancelar",
            .Location = New Point(300, yPos),
            .Size = New Size(100, 35),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
        panelPrincipal.Controls.Add(btnCancelar)

        Me.Controls.Add(panelPrincipal)
    End Sub

    Private Sub CrearCampo(panel As Panel, labelText As String, ByRef textBox As TextBox, yPos As Integer, valorDefecto As String, Optional esPassword As Boolean = False)
        Dim lbl As New Label With {
            .Text = labelText,
            .Location = New Point(20, yPos),
            .Size = New Size(120, 25),
            .Font = New Font("Segoe UI", 10)
        }
        panel.Controls.Add(lbl)

        textBox = New TextBox With {
            .Location = New Point(150, yPos - 2),
            .Size = New Size(280, 25),
            .Font = New Font("Segoe UI", 9),
            .Text = valorDefecto
        }

        If esPassword Then
            textBox.PasswordChar = "*"c
        End If

        panel.Controls.Add(textBox)
    End Sub

    Private Sub CargarConfiguracionActual()
        Try
            ' Cargar desde App.config si existe - VB.NET SINTAXIS
            txtServidor.Text = If(ConfigurationManager.AppSettings("SmtpServidor"), "smtp.gmail.com")
            txtPuerto.Text = If(ConfigurationManager.AppSettings("SmtpPuerto"), "587")
            txtUsuario.Text = If(ConfigurationManager.AppSettings("SmtpUsuario"), "")
            txtContrasena.Text = If(ConfigurationManager.AppSettings("SmtpContrasena"), "")
            chkSSL.Checked = Boolean.Parse(If(ConfigurationManager.AppSettings("SmtpSSL"), "True"))
            txtNombreRemitente.Text = If(ConfigurationManager.AppSettings("NombreRemitente"), "Administración COOPDIASAM")
            txtRemitente.Text = If(ConfigurationManager.AppSettings("CorreoRemitente"), txtUsuario.Text)

        Catch ex As Exception
            ' Si hay error cargando configuración, usar valores por defecto
            lblEstado.Text = "⚠️ Error cargando configuración actual, usando valores por defecto"
            lblEstado.ForeColor = Color.Orange
        End Try
    End Sub

    Private Sub btnProbarConexion_Click(sender As Object, e As EventArgs)
        If Not ValidarCampos() Then
            Return
        End If

        Try
            btnProbarConexion.Enabled = False
            btnProbarConexion.Text = "🔄 Probando..."
            lblEstado.Text = "Probando conexión SMTP..."
            lblEstado.ForeColor = Color.Blue
            Application.DoEvents()

            ' Probar conexión SMTP
            Using client As New System.Net.Mail.SmtpClient(txtServidor.Text.Trim())
                client.Port = Integer.Parse(txtPuerto.Text.Trim())
                client.EnableSsl = chkSSL.Checked
                client.UseDefaultCredentials = False
                client.Credentials = New System.Net.NetworkCredential(txtUsuario.Text.Trim(), txtContrasena.Text)
                client.Timeout = 10000 ' 10 segundos

                ' Enviar correo de prueba
                Using mensaje As New System.Net.Mail.MailMessage()
                    mensaje.From = New System.Net.Mail.MailAddress(txtRemitente.Text.Trim(), txtNombreRemitente.Text.Trim())
                    mensaje.To.Add(txtUsuario.Text.Trim()) ' Enviar a sí mismo como prueba
                    mensaje.Subject = "Prueba de Configuración SMTP - COOPDIASAM"
                    mensaje.Body = $"Esta es una prueba de configuración SMTP realizada el {DateTime.Now:dd/MM/yyyy HH:mm:ss}" & vbCrLf &
                                  "Si recibe este mensaje, la configuración es correcta."
                    mensaje.IsBodyHtml = False

                    client.Send(mensaje)
                End Using
            End Using

            lblEstado.Text = "✅ Conexión exitosa! Configuración correcta."
            lblEstado.ForeColor = Color.Green

        Catch ex As System.Net.Mail.SmtpException
            lblEstado.Text = $"❌ Error SMTP: {ex.Message}"
            lblEstado.ForeColor = Color.Red

        Catch ex As Exception
            lblEstado.Text = $"❌ Error de conexión: {ex.Message}"
            lblEstado.ForeColor = Color.Red

        Finally
            btnProbarConexion.Enabled = True
            btnProbarConexion.Text = "🔍 Probar Conexión"
        End Try
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)
        If Not ValidarCampos() Then
            Return
        End If

        Try
            ' Guardar configuración (en un entorno real, esto debería ir a un archivo de configuración o base de datos)
            ' Por simplicidad, mostraremos los valores que se guardarían

            Dim configuracion As String = "Configuración SMTP guardada:" & vbCrLf &
                                        $"Servidor: {txtServidor.Text.Trim()}" & vbCrLf &
                                        $"Puerto: {txtPuerto.Text.Trim()}" & vbCrLf &
                                        $"Usuario: {txtUsuario.Text.Trim()}" & vbCrLf &
                                        $"SSL: {chkSSL.Checked}" & vbCrLf &
                                        $"Nombre Remitente: {txtNombreRemitente.Text.Trim()}" & vbCrLf &
                                        $"Correo Remitente: {txtRemitente.Text.Trim()}"

            ' En una implementación real, aquí guardarías en App.config o base de datos
            ' GuardarEnConfiguracion()

            MessageBox.Show("⚠️ NOTA: Esta configuración se guarda temporalmente." & vbCrLf &
                          "Para una implementación completa, debe modificar el archivo App.config" & vbCrLf &
                          "o implementar persistencia en base de datos." & vbCrLf & vbCrLf &
                          configuracion,
                          "Configuración",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Information)

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            MessageBox.Show($"Error al guardar configuración: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Function ValidarCampos() As Boolean
        ' Validar servidor
        If String.IsNullOrWhiteSpace(txtServidor.Text) Then
            MessageBox.Show("El servidor SMTP es obligatorio.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtServidor.Focus()
            Return False
        End If

        ' Validar puerto
        Dim puerto As Integer
        If Not Integer.TryParse(txtPuerto.Text.Trim(), puerto) OrElse puerto < 1 OrElse puerto > 65535 Then
            MessageBox.Show("El puerto debe ser un número válido entre 1 y 65535.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPuerto.Focus()
            Return False
        End If

        ' Validar usuario
        If String.IsNullOrWhiteSpace(txtUsuario.Text) Then
            MessageBox.Show("El usuario (email) es obligatorio.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUsuario.Focus()
            Return False
        End If

        ' Validar formato de email del usuario
        If Not ValidarEmail(txtUsuario.Text.Trim()) Then
            MessageBox.Show("El usuario debe ser una dirección de correo válida.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUsuario.Focus()
            Return False
        End If

        ' Validar contraseña
        If String.IsNullOrWhiteSpace(txtContrasena.Text) Then
            MessageBox.Show("La contraseña es obligatoria.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContrasena.Focus()
            Return False
        End If

        ' Validar nombre del remitente
        If String.IsNullOrWhiteSpace(txtNombreRemitente.Text) Then
            MessageBox.Show("El nombre del remitente es obligatorio.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreRemitente.Focus()
            Return False
        End If

        ' Validar correo del remitente
        If String.IsNullOrWhiteSpace(txtRemitente.Text) Then
            txtRemitente.Text = txtUsuario.Text.Trim() ' Usar el mismo del usuario si está vacío
        End If

        If Not ValidarEmail(txtRemitente.Text.Trim()) Then
            MessageBox.Show("El correo del remitente debe ser una dirección válida.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtRemitente.Focus()
            Return False
        End If

        Return True
    End Function

    Private Function ValidarEmail(email As String) As Boolean
        Try
            Dim addr As New System.Net.Mail.MailAddress(email)
            Return addr.Address = email AndAlso email.Contains("@") AndAlso email.Contains(".")
        Catch
            Return False
        End Try
    End Function

    ' Método para guardar en App.config (implementación futura)
    Private Sub GuardarEnConfiguracion()
        ' Este método requeriría permisos especiales y manipulación del App.config
        ' Por ahora solo mostramos la configuración

        ' Ejemplo de cómo se haría:
        ' Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        ' config.AppSettings.Settings("SmtpServidor").Value = txtServidor.Text.Trim()
        ' config.AppSettings.Settings("SmtpPuerto").Value = txtPuerto.Text.Trim()
        ' config.Save(ConfigurationSaveMode.Modified)
        ' ConfigurationManager.RefreshSection("appSettings")
    End Sub

    ' Propiedades públicas para acceder a la configuración desde otros formularios
    Public ReadOnly Property ServidorSMTP As String
        Get
            Return txtServidor.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property PuertoSMTP As Integer
        Get
            Return Integer.Parse(txtPuerto.Text.Trim())
        End Get
    End Property

    Public ReadOnly Property UsuarioSMTP As String
        Get
            Return txtUsuario.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property ContrasenaSMTP As String
        Get
            Return txtContrasena.Text
        End Get
    End Property

    Public ReadOnly Property UsarSSL As Boolean
        Get
            Return chkSSL.Checked
        End Get
    End Property

    Public ReadOnly Property NombreRemitente As String
        Get
            Return txtNombreRemitente.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property CorreoRemitente As String
        Get
            Return txtRemitente.Text.Trim()
        End Get
    End Property

End Class