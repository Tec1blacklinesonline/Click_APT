Imports System.Net
Imports System.Net.Mail
Imports System.IO

Public Class EmailService

    ' Configuración SMTP se carga desde archivo
    Private Shared _configuracion As ConfiguracionSMTP = Nothing

    ' Obtener configuración actual
    Private Shared Function ObtenerConfiguracion() As ConfiguracionSMTP
        If _configuracion Is Nothing Then
            _configuracion = ConfiguracionSMTP.Cargar()
        End If
        Return _configuracion
    End Function

    ' Recargar configuración
    Public Shared Sub RecargarConfiguracion()
        _configuracion = ConfiguracionSMTP.Cargar()
    End Sub

    ' Plantilla del correo
    Private Const ASUNTO_RECIBO As String = "Recibo de Pago - Apartamento {0} - {1}"

    Public Shared Function EnviarReciboPorCorreo(pago As PagoModel, apartamento As Apartamento, rutaPDF As String) As Boolean
        Try
            ' Obtener configuración
            Dim config As ConfiguracionSMTP = ObtenerConfiguracion()

            ' Validar configuración
            If Not config.EsValida() Then
                Throw New Exception("La configuración SMTP no es válida. Configure el servicio de correo en Configuración.")
            End If

            ' Validar que existe el correo del destinatario
            If String.IsNullOrEmpty(apartamento.Correo) Then
                Throw New Exception("El apartamento no tiene correo registrado")
            End If

            ' Crear el mensaje
            Using mensaje As New MailMessage()
                ' Configurar remitente
                mensaje.From = New MailAddress(config.CorreoRemitente, config.NombreRemitente)

                ' Configurar destinatario
                mensaje.To.Add(New MailAddress(apartamento.Correo, apartamento.NombreResidente))

                ' Configurar copia (según el documento, siempre envía copia a apolo20136@gmail.com)
                mensaje.CC.Add("apolo20136@gmail.com")

                ' Configurar asunto
                mensaje.Subject = String.Format(ASUNTO_RECIBO,
                                              $"{apartamento.Torre}{apartamento.NumeroApartamento}",
                                              pago.FechaPago.ToString("MMMM yyyy"))

                ' Configurar cuerpo del mensaje
                mensaje.Body = GenerarCuerpoCorreo(pago, apartamento)
                mensaje.IsBodyHtml = True

                ' Adjuntar el PDF
                If File.Exists(rutaPDF) Then
                    Dim adjunto As New Attachment(rutaPDF)
                    mensaje.Attachments.Add(adjunto)
                End If

                ' Configurar cliente SMTP
                Using cliente As New SmtpClient()
                    cliente.Host = config.ServidorSMTP
                    cliente.Port = config.Puerto
                    cliente.EnableSsl = config.UsarSSL
                    cliente.Credentials = New NetworkCredential(config.UsuarioSMTP, config.ContrasenaSMTP)
                    cliente.DeliveryMethod = SmtpDeliveryMethod.Network
                    cliente.Timeout = config.TimeoutSegundos * 1000 ' Convertir a milisegundos

                    ' Enviar el correo
                    cliente.Send(mensaje)
                End Using

                Return True
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al enviar correo: {ex.Message}")
        End Try
    End Function

    Private Shared Function GenerarCuerpoCorreo(pago As PagoModel, apartamento As Apartamento) As String
        Dim html As New System.Text.StringBuilder()

        html.AppendLine("<!DOCTYPE html>")
        html.AppendLine("<html>")
        html.AppendLine("<head>")
        html.AppendLine("<meta charset='utf-8'>")
        html.AppendLine("<style>")
        html.AppendLine("body { font-family: Arial, sans-serif; color: #333; }")
        html.AppendLine(".container { max-width: 600px; margin: 0 auto; padding: 20px; }")
        html.AppendLine(".header { background-color: #2c3e50; color: white; padding: 20px; text-align: center; }")
        html.AppendLine(".content { padding: 20px; background-color: #f8f9fa; }")
        html.AppendLine(".footer { background-color: #34495e; color: white; padding: 15px; text-align: center; font-size: 12px; }")
        html.AppendLine(".info-box { background-color: white; padding: 15px; margin: 10px 0; border-radius: 5px; }")
        html.AppendLine(".highlight { color: #2980b9; font-weight: bold; }")
        html.AppendLine("</style>")
        html.AppendLine("</head>")
        html.AppendLine("<body>")
        html.AppendLine("<div class='container'>")

        ' Encabezado
        html.AppendLine("<div class='header'>")
        html.AppendLine("<h1>CONJUNTO RESIDENCIAL COOPDIASAM</h1>")
        html.AppendLine("<p>Recibo de Pago Mensual</p>")
        html.AppendLine("</div>")

        ' Contenido
        html.AppendLine("<div class='content'>")
        html.AppendLine($"<h2>Estimado(a) {apartamento.NombreResidente}</h2>")
        html.AppendLine("<p>Adjunto encontrará el recibo de pago correspondiente a su apartamento.</p>")

        ' Información del pago
        html.AppendLine("<div class='info-box'>")
        html.AppendLine("<h3>Detalles del Pago:</h3>")
        html.AppendLine($"<p><strong>Apartamento:</strong> Torre {apartamento.Torre} - Apto {apartamento.NumeroApartamento}</p>")
        html.AppendLine($"<p><strong>Número de Recibo:</strong> <span class='highlight'>{pago.NumeroRecibo}</span></p>")
        html.AppendLine($"<p><strong>Fecha de Pago:</strong> {pago.FechaPago:dd/MM/yyyy}</p>")
        html.AppendLine($"<p><strong>Valor Pagado:</strong> <span class='highlight'>${pago.TotalPagado:N0}</span></p>")
        html.AppendLine($"<p><strong>Saldo Actual:</strong> ${pago.SaldoActual:N0}</p>")
        html.AppendLine("</div>")

        ' Mensaje adicional
        html.AppendLine("<div class='info-box'>")
        html.AppendLine("<p><strong>Importante:</strong> Este es un comprobante electrónico de su pago. ")
        html.AppendLine("Por favor conserve este recibo para futuras referencias.</p>")
        html.AppendLine("</div>")

        html.AppendLine("<p>Si tiene alguna pregunta sobre su pago, no dude en contactarnos.</p>")
        html.AppendLine("</div>")

        ' Pie de página
        html.AppendLine("<div class='footer'>")
        html.AppendLine("<p>Administración: Fernando Gamba - Tel: +57 321-9597100</p>")
        html.AppendLine("<p>Barrio Villa Café - Ibagué, Tolima</p>")
        html.AppendLine($"<p>Este correo fue enviado automáticamente el {DateTime.Now:dd/MM/yyyy HH:mm}</p>")
        html.AppendLine("</div>")

        html.AppendLine("</div>")
        html.AppendLine("</body>")
        html.AppendLine("</html>")

        Return html.ToString()
    End Function

    ' Método para enviar facturas masivas (primer día del mes)
    Public Shared Function EnviarFacturasMasivas(torre As Integer, mes As Integer, año As Integer) As Integer
        Dim enviados As Integer = 0

        Try
            ' Obtener todos los apartamentos de la torre
            Dim apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(torre)

            For Each apartamento In apartamentos
                ' Solo enviar si tiene correo registrado
                If Not String.IsNullOrEmpty(apartamento.Correo) Then
                    Try
                        ' Aquí se generaría la factura del mes
                        ' Por ahora solo contamos los enviados
                        enviados += 1
                    Catch ex As Exception
                        ' Registrar error pero continuar con los demás
                        Continue For
                    End Try
                End If
            Next

        Catch ex As Exception
            Throw New Exception($"Error en envío masivo: {ex.Message}")
        End Try

        Return enviados
    End Function

    ' Método para probar la conexión SMTP
    Public Shared Function ProbarConexionSMTP() As Boolean
        Try
            Dim config As ConfiguracionSMTP = ObtenerConfiguracion()

            If Not config.EsValida() Then
                Return False
            End If

            Using cliente As New SmtpClient()
                cliente.Host = config.ServidorSMTP
                cliente.Port = config.Puerto
                cliente.EnableSsl = config.UsarSSL
                cliente.Credentials = New NetworkCredential(config.UsuarioSMTP, config.ContrasenaSMTP)
                cliente.Timeout = 5000 ' 5 segundos

                ' Enviar un correo de prueba
                Using mensaje As New MailMessage()
                    mensaje.From = New MailAddress(config.CorreoRemitente)
                    mensaje.To.Add(config.CorreoRemitente) ' Enviarse a sí mismo
                    mensaje.Subject = "Prueba de conexión SMTP - COOPDIASAM"
                    mensaje.Body = "Este es un correo de prueba del sistema."

                    cliente.Send(mensaje)
                End Using

                Return True
            End Using

        Catch ex As Exception
            Return False
        End Try
    End Function

End Class