Imports System.Net
Imports System.Net.Mail
Imports System.IO

Public Class EmailService
    ' Configuración SMTP (ajustar según tu servidor)
    Private Const SMTP_HOST As String = "smtp.gmail.com"
    Private Const SMTP_PORT As Integer = 587
    Private Const SMTP_USER As String = "conjuntocoopdiasam@gmail.com" ' Cambiar por correo real
    Private Const SMTP_PASSWORD As String = "tu_contraseña_aqui" ' Cambiar por contraseña real
    Private Const SMTP_SSL As Boolean = True

    ' Correos del conjunto
    Private Const EMAIL_FROM As String = "conjuntocoopdiasam@yahoo.es"
    Private Const EMAIL_FROM_NAME As String = "Conjunto Residencial COOPDIASAM"
    Private Const EMAIL_COPY As String = "apolo20136@gmail.com"

    Public Shared Function EnviarReciboPorCorreo(pago As PagoModel, apartamento As Apartamento, rutaPDF As String) As Boolean
        Try
            ' Validar correo del destinatario
            If String.IsNullOrEmpty(apartamento.Correo) OrElse Not IsValidEmail(apartamento.Correo) Then
                Return False
            End If

            ' Crear mensaje
            Using mensaje As New MailMessage()
                ' Configurar remitente
                mensaje.From = New MailAddress(EMAIL_FROM, EMAIL_FROM_NAME)

                ' Configurar destinatario
                mensaje.To.Add(New MailAddress(apartamento.Correo))

                ' Agregar copia si está configurada
                If Not String.IsNullOrEmpty(EMAIL_COPY) Then
                    mensaje.CC.Add(New MailAddress(EMAIL_COPY))
                End If

                ' Asunto del correo
                mensaje.Subject = $"Recibo de Pago - Apartamento {apartamento.Torre}{apartamento.NumeroApartamento} - {pago.FechaPago:MMMM yyyy}"

                ' Cuerpo del mensaje en HTML
                mensaje.IsBodyHtml = True
                mensaje.Body = GenerarCuerpoCorreo(pago, apartamento)

                ' Adjuntar PDF si existe
                If File.Exists(rutaPDF) Then
                    Dim adjunto As New Attachment(rutaPDF)
                    mensaje.Attachments.Add(adjunto)
                End If

                ' Configurar cliente SMTP
                Using cliente As New SmtpClient(SMTP_HOST, SMTP_PORT)
                    cliente.EnableSsl = SMTP_SSL
                    cliente.Credentials = New NetworkCredential(SMTP_USER, SMTP_PASSWORD)
                    cliente.DeliveryMethod = SmtpDeliveryMethod.Network
                    cliente.Timeout = 30000 ' 30 segundos

                    ' Enviar correo
                    cliente.Send(mensaje)
                    Return True
                End Using
            End Using

        Catch ex As Exception
            ' Log del error (opcional)
            Console.WriteLine($"Error al enviar correo: {ex.Message}")
            Return False
        End Try
    End Function

    Private Shared Function GenerarCuerpoCorreo(pago As PagoModel, apartamento As Apartamento) As String
        Dim html As New System.Text.StringBuilder()

        html.AppendLine("<!DOCTYPE html>")
        html.AppendLine("<html>")
        html.AppendLine("<head>")
        html.AppendLine("<style>")
        html.AppendLine("body { font-family: Arial, sans-serif; color: #333; }")
        html.AppendLine(".container { max-width: 600px; margin: 0 auto; padding: 20px; }")
        html.AppendLine(".header { background-color: #2980b9; color: white; padding: 20px; text-align: center; }")
        html.AppendLine(".content { background-color: #f8f8f8; padding: 20px; margin-top: 20px; }")
        html.AppendLine(".footer { text-align: center; margin-top: 20px; font-size: 12px; color: #666; }")
        html.AppendLine("table { width: 100%; border-collapse: collapse; }")
        html.AppendLine("td { padding: 8px; }")
        html.AppendLine(".label { font-weight: bold; color: #2980b9; }")
        html.AppendLine(".value { text-align: right; }")
        html.AppendLine(".total { font-size: 18px; font-weight: bold; color: #27ae60; }")
        html.AppendLine("</style>")
        html.AppendLine("</head>")
        html.AppendLine("<body>")
        html.AppendLine("<div class='container'>")

        ' Encabezado
        html.AppendLine("<div class='header'>")
        html.AppendLine("<h1>CONJUNTO RESIDENCIAL COOPDIASAM</h1>")
        html.AppendLine("<p>Recibo de Pago</p>")
        html.AppendLine("</div>")

        ' Contenido
        html.AppendLine("<div class='content'>")
        html.AppendLine($"<h2>Estimado(a) {apartamento.NombreResidente}:</h2>")
        html.AppendLine($"<p>Le informamos que hemos recibido su pago correspondiente al apartamento <strong>{apartamento.Torre}{apartamento.NumeroApartamento}</strong>.</p>")

        ' Detalles del pago
        html.AppendLine("<h3>Detalles del Pago:</h3>")
        html.AppendLine("<table>")
        html.AppendLine($"<tr><td class='label'>Número de Recibo:</td><td class='value'>{pago.NumeroRecibo}</td></tr>")
        html.AppendLine($"<tr><td class='label'>Fecha de Pago:</td><td class='value'>{pago.FechaPago:dd/MM/yyyy}</td></tr>")
        html.AppendLine($"<tr><td class='label'>Saldo Anterior:</td><td class='value'>${pago.SaldoAnterior:N0}</td></tr>")
        html.AppendLine($"<tr><td class='label'>Pago Administración:</td><td class='value'>${pago.PagoAdministracion:N0}</td></tr>")
        html.AppendLine($"<tr><td class='label'>Pago Intereses:</td><td class='value'>${pago.PagoIntereses:N0}</td></tr>")
        html.AppendLine($"<tr><td class='label total'>TOTAL PAGADO:</td><td class='value total'>${pago.TotalPagado:N0}</td></tr>")
        html.AppendLine($"<tr><td class='label'>Saldo Actual:</td><td class='value'>${pago.SaldoActual:N0}</td></tr>")
        html.AppendLine("</table>")

        If Not String.IsNullOrEmpty(pago.Observaciones) Then
            html.AppendLine($"<p><strong>Observaciones:</strong> {pago.Observaciones}</p>")
        End If

        html.AppendLine("<p>Adjunto encontrará el recibo de pago en formato PDF para sus registros.</p>")
        html.AppendLine("<p>Agradecemos su puntual pago.</p>")
        html.AppendLine("</div>")

        ' Pie de página
        html.AppendLine("<div class='footer'>")
        html.AppendLine("<p>Fernando Gamba<br/>")
        html.AppendLine("Administrador<br/>")
        html.AppendLine("Tel: +57 321-9597100<br/>")
        html.AppendLine("Barrio Villa Café - Ibagué, Tolima</p>")
        html.AppendLine("</div>")

        html.AppendLine("</div>")
        html.AppendLine("</body>")
        html.AppendLine("</html>")

        Return html.ToString()
    End Function

    Private Shared Function IsValidEmail(email As String) As Boolean
        Try
            Dim addr As New MailAddress(email)
            Return addr.Address = email
        Catch
            Return False
        End Try
    End Function

    ' Método para enviar correo de prueba
    Public Shared Function ProbarConfiguracionSMTP() As String
        Try
            Using cliente As New SmtpClient(SMTP_HOST, SMTP_PORT)
                cliente.EnableSsl = SMTP_SSL
                cliente.Credentials = New NetworkCredential(SMTP_USER, SMTP_PASSWORD)
                cliente.DeliveryMethod = SmtpDeliveryMethod.Network

                Using mensaje As New MailMessage()
                    mensaje.From = New MailAddress(EMAIL_FROM)
                    mensaje.To.Add(EMAIL_FROM)
                    mensaje.Subject = "Prueba de configuración SMTP"
                    mensaje.Body = "Este es un mensaje de prueba para verificar la configuración SMTP."

                    cliente.Send(mensaje)
                    Return "Configuración SMTP correcta"
                End Using
            End Using
        Catch ex As Exception
            Return $"Error: {ex.Message}"
        End Try
    End Function
End Class