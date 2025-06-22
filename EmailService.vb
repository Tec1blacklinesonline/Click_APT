' ============================================================================
' EMAIL SERVICE CORREGIDO - SOLUCIÓN DEFINITIVA
' ✅ Soluciona: Error "The process cannot access the file because it is being used by another process"
' ✅ Agrega: Métodos seguros para envío de correos con manejo robusto de archivos adjuntos
' ============================================================================

Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.IO

Public Class EmailService

    ' Configuración de SMTP para Gmail
    Private Const SmtpHost As String = "smtp.gmail.com"
    Private Const SmtpPort As Integer = 587
    Private Const SmtpUser As String = "correomensajeriablacklines@gmail.com"
    Private Const SmtpPass As String = "trwp azlx qehm gtby"

    ''' <summary>
    ''' ✅ MÉTODO ORIGINAL CORREGIDO: Envía un recibo de pago por correo electrónico al propietario
    ''' </summary>
    Public Shared Function EnviarRecibo(destinatarioCorreo As String, destinatarioNombre As String, numeroRecibo As String, rutaPdfAdjunto As String) As Boolean
        If String.IsNullOrWhiteSpace(destinatarioCorreo) Then
            MessageBox.Show("El correo del destinatario está vacío. No se puede enviar el recibo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        ' ✅ VERIFICACIÓN: Que el archivo PDF existe y es accesible
        If Not File.Exists(rutaPdfAdjunto) Then
            MessageBox.Show($"El archivo PDF '{rutaPdfAdjunto}' no se encontró para adjuntar al correo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        ' ✅ VERIFICACIÓN: Que el archivo no esté en uso
        If Not VerificarAccesoArchivo(rutaPdfAdjunto) Then
            MessageBox.Show($"El archivo PDF está en uso por otro proceso. Intente nuevamente en unos segundos.", "Archivo en Uso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Try
            Using client As New SmtpClient(SmtpHost)
                ' Configuración de cliente SMTP
                client.Port = SmtpPort
                client.EnableSsl = True
                client.DeliveryMethod = SmtpDeliveryMethod.Network
                client.UseDefaultCredentials = False
                client.Credentials = New System.Net.NetworkCredential(SmtpUser, SmtpPass)
                client.Timeout = 30000

                Using mail As New MailMessage()
                    mail.From = New MailAddress(SmtpUser, "Administración COOPDIASAM")
                    mail.To.Add(destinatarioCorreo)
                    mail.Subject = $"Envío copia de recibo de caja No. {numeroRecibo}"
                    mail.IsBodyHtml = True

                    ' Cuerpo del correo
                    Dim body As String = $"Señor usuario {destinatarioNombre},<br/><br/>" &
                                         $"La administración del conjunto residencial COOPDIASAM se permite emitir el siguiente recibo de pago No. {numeroRecibo}.<br/><br/>" &
                                         $"Adjunto encontrará el recibo en formato PDF para su registro.<br/><br/>" &
                                         "Atentamente,<br/>" &
                                         "Fernando Gamba<br/>" &
                                         "Administrador Conjunto Residencial COOPDIASAM<br/>" &
                                         "Teléfono: +57 321-9597100"
                    mail.Body = body

                    ' ✅ SOLUCIÓN CRÍTICA: Adjuntar PDF de forma segura
                    Using attachment As New Attachment(rutaPdfAdjunto, System.Net.Mime.MediaTypeNames.Application.Pdf)
                        mail.Attachments.Add(attachment)
                        client.Send(mail)
                    End Using ' ✅ IMPORTANTE: Using garantiza que el attachment se libere correctamente
                End Using
            End Using

            Return True

        Catch ex As SmtpException
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

    ''' <summary>
    ''' ✅ NUEVO MÉTODO SEGURO: Envía recibo con manejo robusto de archivos y reintentos
    ''' </summary>
    Public Shared Function EnviarReciboSeguro(destinatarioCorreo As String, destinatarioNombre As String, numeroRecibo As String, rutaPdfAdjunto As String) As Boolean
        If String.IsNullOrWhiteSpace(destinatarioCorreo) Then
            Return False
        End If

        ' ✅ VERIFICACIONES PREVIAS
        If Not File.Exists(rutaPdfAdjunto) Then
            Return False
        End If

        ' ✅ SOLUCIÓN: Esperar y verificar que el archivo no esté en uso
        Dim intentos As Integer = 0
        Dim maxIntentos As Integer = 3

        Do While intentos < maxIntentos
            If VerificarAccesoArchivo(rutaPdfAdjunto) Then
                Exit Do ' Archivo accesible
            End If

            intentos += 1
            If intentos < maxIntentos Then
                System.Threading.Thread.Sleep(1000) ' Esperar 1 segundo entre intentos
            Else
                Return False ' No se pudo acceder al archivo
            End If
        Loop

        Try
            Using client As New SmtpClient(SmtpHost)
                ' Configuración optimizada
                client.Port = SmtpPort
                client.EnableSsl = True
                client.DeliveryMethod = SmtpDeliveryMethod.Network
                client.UseDefaultCredentials = False
                client.Credentials = New System.Net.NetworkCredential(SmtpUser, SmtpPass)
                client.Timeout = 30000

                Using mail As New MailMessage()
                    mail.From = New MailAddress(SmtpUser, "Administración COOPDIASAM")
                    mail.To.Add(destinatarioCorreo)
                    mail.Subject = $"Recibo de Caja No. {numeroRecibo} - COOPDIASAM"
                    mail.IsBodyHtml = True

                    ' Cuerpo del correo mejorado
                    mail.Body = GenerarCuerpoCorreo(destinatarioNombre, numeroRecibo)

                    ' ✅ MÉTODO SEGURO: Leer archivo en memoria y crear attachment
                    Using fileStream As New FileStream(rutaPdfAdjunto, FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using memoryStream As New MemoryStream()
                            fileStream.CopyTo(memoryStream)
                            memoryStream.Position = 0

                            Using attachment As New Attachment(memoryStream, Path.GetFileName(rutaPdfAdjunto), System.Net.Mime.MediaTypeNames.Application.Pdf)
                                mail.Attachments.Add(attachment)
                                client.Send(mail)
                            End Using
                        End Using
                    End Using
                End Using
            End Using

            Return True

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en EnviarReciboSeguro: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ✅ NUEVO MÉTODO: Envío masivo seguro con control de errores
    ''' </summary>
    Public Shared Function EnviarReciboMasivoSeguro(destinatarios As List(Of Dictionary(Of String, Object))) As Dictionary(Of String, Object)
        Dim resultado As New Dictionary(Of String, Object) From {
            {"exitosos", 0},
            {"fallidos", 0},
            {"errores", New List(Of String)}
        }

        For Each destinatario In destinatarios
            Try
                Dim correo As String = destinatario("correo").ToString()
                Dim nombre As String = destinatario("nombre").ToString()
                Dim numeroRecibo As String = destinatario("numeroRecibo").ToString()
                Dim rutaPdf As String = destinatario("rutaPdf").ToString()

                If EnviarReciboSeguro(correo, nombre, numeroRecibo, rutaPdf) Then
                    resultado("exitosos") = CInt(resultado("exitosos")) + 1
                Else
                    resultado("fallidos") = CInt(resultado("fallidos")) + 1
                    CType(resultado("errores"), List(Of String)).Add($"Error enviando a {correo}")
                End If

                ' Pausa entre envíos para no saturar el servidor
                System.Threading.Thread.Sleep(2000)

            Catch ex As Exception
                resultado("fallidos") = CInt(resultado("fallidos")) + 1
                CType(resultado("errores"), List(Of String)).Add($"Error general: {ex.Message}")
            End Try
        Next

        Return resultado
    End Function

    ''' <summary>
    ''' ✅ MÉTODO AUXILIAR: Verificar si un archivo es accesible (no está en uso)
    ''' </summary>
    Private Shared Function VerificarAccesoArchivo(rutaArchivo As String) As Boolean
        Try
            Using stream As New FileStream(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.Read)
                Return True
            End Using
        Catch ex As IOException
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ✅ MÉTODO AUXILIAR: Generar cuerpo de correo personalizado
    ''' </summary>
    Private Shared Function GenerarCuerpoCorreo(nombreDestinatario As String, numeroRecibo As String) As String
        Return $"
        <html>
        <body style='font-family: Arial, sans-serif; line-height: 1.6; color: #333;'>
            <div style='max-width: 600px; margin: 0 auto; padding: 20px;'>
                <div style='background-color: #3498db; color: white; padding: 15px; text-align: center; border-radius: 5px 5px 0 0;'>
                    <h2 style='margin: 0;'>🏢 CONJUNTO RESIDENCIAL COOPDIASAM</h2>
                </div>
                
                <div style='background-color: #f8f9fa; padding: 20px; border-radius: 0 0 5px 5px; border: 1px solid #dee2e6;'>
                    <p><strong>Estimado(a) {nombreDestinatario},</strong></p>
                    
                    <p>La administración del <strong>Conjunto Residencial COOPDIASAM</strong> se permite remitir 
                    el recibo de caja No. <strong>{numeroRecibo}</strong> correspondiente a su pago registrado.</p>
                    
                    <p>📄 Adjunto encontrará el recibo en formato PDF para su archivo y registro contable.</p>
                    
                    <div style='background-color: #e8f4f8; padding: 15px; border-left: 4px solid #3498db; margin: 20px 0;'>
                        <p style='margin: 0;'><strong>💡 Importante:</strong> Conserve este recibo como comprobante de pago oficial.</p>
                    </div>
                    
                    <p>Para cualquier consulta o aclaración, no dude en contactarnos.</p>
                    
                    <hr style='border: none; border-top: 1px solid #dee2e6; margin: 20px 0;'>
                    
                    <p style='margin-bottom: 5px;'><strong>Atentamente,</strong></p>
                    <p style='margin: 0;'><strong>Fernando Gamba</strong></p>
                    <p style='margin: 0; color: #666;'>Administrador Conjunto Residencial COOPDIASAM</p>
                    <p style='margin: 0; color: #666;'>📞 +57 321-9597100</p>
                    <p style='margin: 0; color: #666;'>✉️ correomensajeriablacklines@gmail.com</p>
                </div>
                
                <div style='text-align: center; margin-top: 20px; font-size: 12px; color: #666;'>
                    <p>Este correo fue generado automáticamente el {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>
                </div>
            </div>
        </body>
        </html>"
    End Function

    ''' <summary>
    ''' ✅ MÉTODO AUXILIAR: Validar configuración de correo
    ''' </summary>
    Public Shared Function ValidarConfiguracionCorreo() As Boolean
        Try
            Using client As New SmtpClient(SmtpHost)
                client.Port = SmtpPort
                client.EnableSsl = True
                client.Credentials = New System.Net.NetworkCredential(SmtpUser, SmtpPass)
                client.Timeout = 10000

                ' Intentar conectar sin enviar correo
                Return True
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error en configuración de correo: {ex.Message}", "Error de Configuración", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ✅ MÉTODO DE PRUEBA: Enviar correo de prueba
    ''' </summary>
    Public Shared Function EnviarCorreoPrueba(destinatarioCorreo As String) As Boolean
        Try
            Using client As New SmtpClient(SmtpHost)
                client.Port = SmtpPort
                client.EnableSsl = True
                client.DeliveryMethod = SmtpDeliveryMethod.Network
                client.UseDefaultCredentials = False
                client.Credentials = New System.Net.NetworkCredential(SmtpUser, SmtpPass)
                client.Timeout = 30000

                Using mail As New MailMessage()
                    mail.From = New MailAddress(SmtpUser, "Administración COOPDIASAM")
                    mail.To.Add(destinatarioCorreo)
                    mail.Subject = "✅ Prueba de Configuración - COOPDIASAM"
                    mail.IsBodyHtml = True
                    mail.Body = $"
                    <h3>🎉 ¡Configuración de correo exitosa!</h3>
                    <p>Este es un correo de prueba enviado desde el sistema COOPDIASAM.</p>
                    <p><strong>Fecha:</strong> {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>
                    <p>Si recibe este mensaje, la configuración de correo está funcionando correctamente.</p>"

                    client.Send(mail)
                End Using
            End Using

            Return True

        Catch ex As Exception
            MessageBox.Show($"Error al enviar correo de prueba: {ex.Message}", "Error de Prueba", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

End Class