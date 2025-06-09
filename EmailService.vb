' EmailService.vb
' Este archivo maneja el envío de correos electrónicos.

Imports System.Net.Mail
Imports System.Windows.Forms ' Para MessageBox (considerar logging para producción)
Imports System.IO ' Para adjuntar archivos

Public Class EmailService

    ' Configuración de SMTP para Gmail (corregida)
    Private Const SmtpHost As String = "smtp.gmail.com" ' Cambiado a Gmail ya que usas @gmail.com
    Private Const SmtpPort As Integer = 587 ' Puerto correcto para Gmail
    Private Const SmtpUser As String = "correomensajeriablacklines@gmail.com" ' Tu dirección de correo
    Private Const SmtpPass As String = "trwp azlx qehm gtby" ' Contraseña de aplicación de Gmail

    ''' <summary>
    ''' Envía un recibo de pago por correo electrónico al propietario.
    ''' </summary>
    ''' <param name="destinatarioCorreo">Dirección de correo del propietario.</param>
    ''' <param name="destinatarioNombre">Nombre del propietario.</param>
    ''' <param name="numeroRecibo">Número del recibo de pago.</param>
    ''' <param name="rutaPdfAdjunto">Ruta completa del archivo PDF a adjuntar.</param>
    ''' <returns>True si el correo se envió exitosamente, False en caso contrario.</returns>
    Public Shared Function EnviarRecibo(destinatarioCorreo As String, destinatarioNombre As String, numeroRecibo As String, rutaPdfAdjunto As String) As Boolean
        If String.IsNullOrWhiteSpace(destinatarioCorreo) Then
            MessageBox.Show("El correo del destinatario está vacío. No se puede enviar el recibo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Try
            Using client As New SmtpClient(SmtpHost)
                ' Configuración de cliente SMTP
                client.Port = SmtpPort
                client.EnableSsl = True ' Habilitar SSL/TLS
                client.DeliveryMethod = SmtpDeliveryMethod.Network
                client.UseDefaultCredentials = False ' Importante: debe estar en False
                client.Credentials = New System.Net.NetworkCredential(SmtpUser, SmtpPass)

                ' Configuración adicional para Gmail
                client.Timeout = 30000 ' 30 segundos de timeout

                Using mail As New MailMessage()
                    mail.From = New MailAddress(SmtpUser, "Administración COOPDIASAM")
                    mail.To.Add(destinatarioCorreo)
                    mail.Subject = $"Envío copia de recibo de caja No. {numeroRecibo}"
                    mail.IsBodyHtml = True ' Si quieres usar HTML en el cuerpo del correo

                    ' Cuerpo del correo
                    Dim body As String = $"Señor usuario {destinatarioNombre},<br/><br/>" &
                                         $"La administración del conjunto residencial COOPDIASAM se permite emitir el siguiente recibo de pago No. {numeroRecibo}.<br/><br/>" &
                                         $"Adjunto encontrará el recibo en formato PDF para su registro.<br/><br/>" &
                                         "Atentamente,<br/>" &
                                         "Fernando Gamba<br/>" &
                                         "Administrador Conjunto Residencial COOPDIASAM<br/>" &
                                         "Teléfono: +57 321-9597100"
                    mail.Body = body

                    ' Adjuntar el PDF
                    If File.Exists(rutaPdfAdjunto) Then
                        Using attachment As New Attachment(rutaPdfAdjunto, System.Net.Mime.MediaTypeNames.Application.Pdf)
                            mail.Attachments.Add(attachment)

                            ' Enviar el correo
                            client.Send(mail)
                        End Using
                    Else
                        MessageBox.Show($"El archivo PDF '{rutaPdfAdjunto}' no se encontró para adjuntar al correo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return False
                    End If
                End Using

                Return True

            End Using

        Catch ex As SmtpException
            ' Mensajes de error más específicos
            Dim mensajeError As String = ""

            Select Case ex.StatusCode
                Case SmtpStatusCode.MailboxBusy
                    mensajeError = "Error de autenticación o servidor ocupado. Verifique que:" & vbCrLf &
                                  "1. La contraseña de aplicación sea correcta" & vbCrLf &
                                  "2. Tenga habilitada la verificación en 2 pasos en Gmail" & vbCrLf &
                                  "3. Esté usando una 'Contraseña de aplicación' no la contraseña normal"
                Case SmtpStatusCode.MailboxUnavailable
                    mensajeError = "La dirección de correo no es válida o no existe."
                Case SmtpStatusCode.InsufficientStorage
                    mensajeError = "No hay suficiente espacio en el servidor de correo."
                Case SmtpStatusCode.TransactionFailed
                    mensajeError = "Falló la transacción de correo. Verifique la conexión a internet."
                Case SmtpStatusCode.GeneralFailure
                    mensajeError = "Error general del servidor SMTP. Verifique la configuración de autenticación."
                Case Else
                    mensajeError = $"Error SMTP: {ex.Message}" & vbCrLf & $"Código: {ex.StatusCode}"
            End Select

            MessageBox.Show(mensajeError, "Error de Correo", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Catch ex As Exception
            MessageBox.Show($"Error general al enviar el correo: {ex.Message}", "Error de Correo", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

End Class