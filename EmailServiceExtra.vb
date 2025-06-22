' ============================================================================
' EMAIL SERVICE EXTRA - SERVICIO DE CORREO PARA PAGOS EXTRA
' ✅ Especializado para envío de recibos de multas, adiciones, sanciones, etc.
' ✅ Extiende EmailService para pagos extra
' ============================================================================

Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.IO

Public Class EmailServiceExtra

    ' Usar la misma configuración SMTP del EmailService principal
    Private Const SmtpHost As String = "smtp.gmail.com"
    Private Const SmtpPort As Integer = 587
    Private Const SmtpUser As String = "correomensajeriablacklines@gmail.com"
    Private Const SmtpPass As String = "trwp azlx qehm gtby"

    ''' <summary>
    ''' Envía recibo de pago extra por correo electrónico
    ''' </summary>
    Public Shared Function EnviarReciboPagoExtra(destinatarioCorreo As String, destinatarioNombre As String, numeroRecibo As String, tipoPago As String, rutaPdfAdjunto As String) As Boolean
        If String.IsNullOrWhiteSpace(destinatarioCorreo) Then
            Return False
        End If

        If Not File.Exists(rutaPdfAdjunto) Then
            Return False
        End If

        ' Verificar acceso al archivo
        If Not VerificarAccesoArchivo(rutaPdfAdjunto) Then
            Return False
        End If

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
                    mail.Subject = GenerarAsuntoPagoExtra(tipoPago, numeroRecibo)
                    mail.IsBodyHtml = True
                    mail.Body = GenerarCuerpoCorreoPagoExtra(destinatarioNombre, numeroRecibo, tipoPago)

                    ' Adjuntar PDF de forma segura
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
            System.Diagnostics.Debug.WriteLine($"Error en EnviarReciboPagoExtra: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Genera asunto personalizado según tipo de pago extra
    ''' </summary>
    Private Shared Function GenerarAsuntoPagoExtra(tipoPago As String, numeroRecibo As String) As String
        Select Case tipoPago.ToUpper()
            Case "MULTA"
                Return $"🚨 Recibo de Multa No. {numeroRecibo} - COOPDIASAM"
            Case "SANCION"
                Return $"⚠️ Recibo de Sanción No. {numeroRecibo} - COOPDIASAM"
            Case "ADICION"
                Return $"💰 Recibo de Cobro Adicional No. {numeroRecibo} - COOPDIASAM"
            Case "REPARACION"
                Return $"🔧 Recibo de Reparación No. {numeroRecibo} - COOPDIASAM"
            Case "PAGO_ATRASADO"
                Return $"⏰ Recibo de Pago Atrasado No. {numeroRecibo} - COOPDIASAM"
            Case "SERVICIO_EXTRA"
                Return $"🛠️ Recibo de Servicio Extra No. {numeroRecibo} - COOPDIASAM"
            Case Else
                Return $"📄 Recibo de Pago Extra No. {numeroRecibo} - COOPDIASAM"
        End Select
    End Function

    ''' <summary>
    ''' Genera cuerpo de correo personalizado para pagos extra
    ''' </summary>
    Private Shared Function GenerarCuerpoCorreoPagoExtra(nombreDestinatario As String, numeroRecibo As String, tipoPago As String) As String
        Dim emoji As String = ObtenerEmojiTipoPago(tipoPago)
        Dim descripcionTipo As String = ObtenerDescripcionTipoPago(tipoPago)
        Dim colorTema As String = ObtenerColorTema(tipoPago)

        Return $"
        <html>
        <body style='font-family: Arial, sans-serif; line-height: 1.6; color: #333;'>
            <div style='max-width: 600px; margin: 0 auto; padding: 20px;'>
                <div style='background-color: {colorTema}; color: white; padding: 15px; text-align: center; border-radius: 5px 5px 0 0;'>
                    <h2 style='margin: 0;'>{emoji} CONJUNTO RESIDENCIAL COOPDIASAM</h2>
                </div>
                
                <div style='background-color: #f8f9fa; padding: 20px; border-radius: 0 0 5px 5px; border: 1px solid #dee2e6;'>
                    <p><strong>Estimado(a) {nombreDestinatario},</strong></p>
                    
                    <p>La administración del <strong>Conjunto Residencial COOPDIASAM</strong> se permite remitir 
                    el recibo No. <strong>{numeroRecibo}</strong> correspondiente a un <strong>{descripcionTipo}</strong>.</p>
                    
                    <div style='background-color: {ObtenerColorFondo(tipoPago)}; padding: 15px; border-left: 4px solid {colorTema}; margin: 20px 0;'>
                        <p style='margin: 0;'><strong>{emoji} Tipo de Pago:</strong> {tipoPago.Replace("_", " ")}</p>
                        <p style='margin: 5px 0 0 0;'><strong>📄 Número de Recibo:</strong> {numeroRecibo}</p>
                    </div>
                    
                    <p>📎 Adjunto encontrará el recibo en formato PDF para su archivo y registro contable.</p>
                    
                    {GenerarMensajeEspecificoTipo(tipoPago)}
                    
                    <p>Para cualquier consulta o aclaración sobre este pago, no dude en contactarnos.</p>
                    
                    <hr style='border: none; border-top: 1px solid #dee2e6; margin: 20px 0;'>
                    
                    <p style='margin-bottom: 5px;'><strong>Atentamente,</strong></p>
                    <p style='margin: 0;'><strong>Fernando Gamba</strong></p>
                    <p style='margin: 0; color: #666;'>Administrador Conjunto Residencial COOPDIASAM</p>
                    <p style='margin: 0; color: #666;'>📞 +57 321-9597100</p>
                    <p style='margin: 0; color: #666;'>✉️ correomensajeriablacklines@gmail.com</p>
                </div>
                
                <div style='text-align: center; margin-top: 20px; font-size: 12px; color: #666;'>
                    <p>Este correo fue generado automáticamente el {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>
                    <p style='color: #999;'>IMPORTANTE: Conserve este recibo como comprobante de pago oficial</p>
                </div>
            </div>
        </body>
        </html>"
    End Function

    ''' <summary>
    ''' Obtiene emoji según tipo de pago
    ''' </summary>
    Private Shared Function ObtenerEmojiTipoPago(tipoPago As String) As String
        Select Case tipoPago.ToUpper()
            Case "MULTA"
                Return "🚨"
            Case "SANCION"
                Return "⚠️"
            Case "ADICION"
                Return "💰"
            Case "REPARACION"
                Return "🔧"
            Case "PAGO_ATRASADO"
                Return "⏰"
            Case "SERVICIO_EXTRA"
                Return "🛠️"
            Case Else
                Return "📄"
        End Select
    End Function

    ''' <summary>
    ''' Obtiene descripción amigable del tipo de pago
    ''' </summary>
    Private Shared Function ObtenerDescripcionTipoPago(tipoPago As String) As String
        Select Case tipoPago.ToUpper()
            Case "MULTA"
                Return "pago de multa"
            Case "SANCION"
                Return "sanción económica"
            Case "ADICION"
                Return "cobro adicional"
            Case "REPARACION"
                Return "cobro por reparación"
            Case "PAGO_ATRASADO"
                Return "pago atrasado"
            Case "SERVICIO_EXTRA"
                Return "servicio extraordinario"
            Case Else
                Return "pago especial"
        End Select
    End Function

    ''' <summary>
    ''' Obtiene color de tema según tipo de pago
    ''' </summary>
    Private Shared Function ObtenerColorTema(tipoPago As String) As String
        Select Case tipoPago.ToUpper()
            Case "MULTA", "SANCION"
                Return "#dc3545" ' Rojo
            Case "ADICION", "SERVICIO_EXTRA"
                Return "#007bff" ' Azul
            Case "REPARACION"
                Return "#fd7e14" ' Naranja
            Case "PAGO_ATRASADO"
                Return "#6f42c1" ' Morado
            Case Else
                Return "#6c757d" ' Gris
        End Select
    End Function

    ''' <summary>
    ''' Obtiene color de fondo según tipo de pago
    ''' </summary>
    Private Shared Function ObtenerColorFondo(tipoPago As String) As String
        Select Case tipoPago.ToUpper()
            Case "MULTA", "SANCION"
                Return "#f8d7da" ' Rojo claro
            Case "ADICION", "SERVICIO_EXTRA"
                Return "#d1ecf1" ' Azul claro
            Case "REPARACION"
                Return "#ffeaa7" ' Naranja claro
            Case "PAGO_ATRASADO"
                Return "#e2d9f3" ' Morado claro
            Case Else
                Return "#e9ecef" ' Gris claro
        End Select
    End Function

    ''' <summary>
    ''' Genera mensaje específico según tipo de pago
    ''' </summary>
    Private Shared Function GenerarMensajeEspecificoTipo(tipoPago As String) As String
        Select Case tipoPago.ToUpper()
            Case "MULTA"
                Return "<div style='background-color: #f8d7da; padding: 10px; border-radius: 5px; margin: 15px 0;'>" &
                       "<p style='margin: 0; color: #721c24;'><strong>📌 Importante:</strong> Esta multa fue impuesta según el reglamento de propiedad horizontal. " &
                       "Le solicitamos revisar las normas de convivencia para evitar futuras infracciones.</p></div>"

            Case "SANCION"
                Return "<div style='background-color: #f8d7da; padding: 10px; border-radius: 5px; margin: 15px 0;'>" &
                       "<p style='margin: 0; color: #721c24;'><strong>⚠️ Advertencia:</strong> Esta sanción se aplicó por incumplimiento del reglamento interno. " &
                       "El pago no exime de futuras sanciones por reincidencia.</p></div>"

            Case "REPARACION"
                Return "<div style='background-color: #ffeaa7; padding: 10px; border-radius: 5px; margin: 15px 0;'>" &
                       "<p style='margin: 0; color: #856404;'><strong>🔧 Reparación:</strong> Este cobro corresponde a daños causados a las áreas comunes. " &
                       "Recuerde que es responsable de los daños causados por familiares o visitantes.</p></div>"

            Case "PAGO_ATRASADO"
                Return "<div style='background-color: #e2d9f3; padding: 10px; border-radius: 5px; margin: 15px 0;'>" &
                       "<p style='margin: 0; color: #6f42c1;'><strong>⏰ Pago Atrasado:</strong> Le recordamos la importancia de realizar los pagos en las fechas establecidas " &
                       "para evitar inconvenientes administrativos.</p></div>"

            Case Else
                Return "<div style='background-color: #e9ecef; padding: 10px; border-radius: 5px; margin: 15px 0;'>" &
                       "<p style='margin: 0; color: #495057;'><strong>💡 Información:</strong> Este es un pago especial autorizado por la administración del conjunto.</p></div>"
        End Select
    End Function

    ''' <summary>
    ''' Verifica si un archivo es accesible
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

End Class