' ============================================================================
' SERVICIO DE NOTIFICACIONES AUTOMÁTICAS
' Funcionalidad faltante: Sistema de notificaciones para pagos vencidos
' ============================================================================

Imports System.Data.SQLite
Imports System.Net.Mail
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class ServicioNotificaciones

    ''' <summary>
    ''' Envía notificaciones automáticas de pagos vencidos
    ''' </summary>
    Public Shared Async Function EnviarNotificacionesPagosVencidos() As Task(Of ResultadoNotificaciones)
        Dim resultado As New ResultadoNotificaciones()

        Try
            ' Obtener apartamentos con pagos vencidos
            Dim apartamentosVencidos = ObtenerApartamentosConPagosVencidos()

            If apartamentosVencidos.Count = 0 Then
                resultado.Mensaje = "No hay apartamentos con pagos vencidos"
                resultado.EsExitoso = True
                Return resultado
            End If

            ' Enviar notificaciones en paralelo
            Dim tareas As New List(Of Task(Of Boolean))

            For Each apartamento In apartamentosVencidos
                If Not String.IsNullOrEmpty(apartamento.Correo) Then
                    Dim tarea = Task.Run(Function() EnviarNotificacionIndividual(apartamento))
                    tareas.Add(tarea)
                End If
            Next

            ' Esperar a que terminen todas las tareas
            Dim resultados() As Boolean = Await Task.WhenAll(tareas)

            ' Contar éxitos y fallos
            resultado.NotificacionesEnviadas = resultados.Count(Function(r) r)
            resultado.NotificacionesFallidas = resultados.Count(Function(r) Not r)
            resultado.EsExitoso = True
            resultado.Mensaje = $"Procesadas {apartamentosVencidos.Count} notificaciones. Enviadas: {resultado.NotificacionesEnviadas}, Fallidas: {resultado.NotificacionesFallidas}"

        Catch ex As Exception
            resultado.EsExitoso = False
            resultado.Mensaje = $"Error al enviar notificaciones: {ex.Message}"
        End Try

        Return resultado
    End Function

    ''' <summary>
    ''' Obtiene apartamentos con pagos vencidos
    ''' </summary>
    Private Shared Function ObtenerApartamentosConPagosVencidos() As List(Of ApartamentoConDeuda)
        Dim apartamentos As New List(Of ApartamentoConDeuda)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT DISTINCT 
                        a.id_apartamentos,
                        a.id_torre,
                        a.numero_apartamento,
                        a.nombre_residente,
                        a.correo,
                        a.telefono,
                        COUNT(c.id_cuota) as cuotas_vencidas,
                        SUM(c.valor_cuota) as total_deuda,
                        MIN(c.fecha_vencimiento) as fecha_mas_antigua,
                        MAX(c.fecha_vencimiento) as fecha_mas_reciente
                    FROM Apartamentos a
                    INNER JOIN cuotas_generadas_apartamento c ON a.id_apartamentos = c.id_apartamentos
                    WHERE c.estado = 'pendiente' 
                        AND date(c.fecha_vencimiento) < date('now')
                        AND a.correo IS NOT NULL 
                        AND a.correo != ''
                    GROUP BY a.id_apartamentos
                    ORDER BY MIN(c.fecha_vencimiento) ASC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim apartamento As New ApartamentoConDeuda With {
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .Torre = Convert.ToInt32(reader("id_torre")),
                                .NumeroApartamento = reader("numero_apartamento").ToString(),
                                .NombreResidente = If(IsDBNull(reader("nombre_residente")), "", reader("nombre_residente").ToString()),
                                .Correo = reader("correo").ToString(),
                                .Telefono = If(IsDBNull(reader("telefono")), "", reader("telefono").ToString()),
                                .CuotasVencidas = Convert.ToInt32(reader("cuotas_vencidas")),
                                .TotalDeuda = Convert.ToDecimal(reader("total_deuda")),
                                .FechaMasAntigua = Convert.ToDateTime(reader("fecha_mas_antigua")),
                                .FechaMasReciente = Convert.ToDateTime(reader("fecha_mas_reciente"))
                            }

                            ' Calcular días de mora
                            apartamento.DiasEnMora = (DateTime.Now.Date - apartamento.FechaMasAntigua.Date).Days

                            apartamentos.Add(apartamento)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al obtener apartamentos con pagos vencidos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return apartamentos
    End Function

    ''' <summary>
    ''' Envía notificación individual a un apartamento
    ''' </summary>
    Private Shared Function EnviarNotificacionIndividual(apartamento As ApartamentoConDeuda) As Boolean
        Try
            ' Determinar tipo de notificación según días de mora
            Dim tipoNotificacion As TipoNotificacion
            If apartamento.DiasEnMora <= 7 Then
                tipoNotificacion = TipoNotificacion.Recordatorio
            ElseIf apartamento.DiasEnMora <= 30 Then
                tipoNotificacion = TipoNotificacion.PrimerAviso
            ElseIf apartamento.DiasEnMora <= 60 Then
                tipoNotificacion = TipoNotificacion.SegundoAviso
            Else
                tipoNotificacion = TipoNotificacion.AvisoUrgente
            End If

            ' Generar contenido del correo
            Dim asunto As String = GenerarAsuntoNotificacion(tipoNotificacion, apartamento)
            Dim cuerpo As String = GenerarCuerpoNotificacion(tipoNotificacion, apartamento)

            ' Enviar correo
            Return EnviarCorreoNotificacion(apartamento.Correo, apartamento.NombreResidente, asunto, cuerpo)

        Catch ex As Exception
            ' Log del error
            System.Diagnostics.Debug.WriteLine($"Error al enviar notificación a {apartamento.Correo}: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Genera el asunto del correo según el tipo de notificación
    ''' </summary>
    Private Shared Function GenerarAsuntoNotificacion(tipo As TipoNotificacion, apartamento As ApartamentoConDeuda) As String
        Select Case tipo
            Case TipoNotificacion.Recordatorio
                Return $"Recordatorio de Pago - Torre {apartamento.Torre} Apt {apartamento.NumeroApartamento}"
            Case TipoNotificacion.PrimerAviso
                Return $"PRIMER AVISO - Pago Vencido - Torre {apartamento.Torre} Apt {apartamento.NumeroApartamento}"
            Case TipoNotificacion.SegundoAviso
                Return $"SEGUNDO AVISO - Pago Vencido - Torre {apartamento.Torre} Apt {apartamento.NumeroApartamento}"
            Case TipoNotificacion.AvisoUrgente
                Return $"AVISO URGENTE - Pago en Mora - Torre {apartamento.Torre} Apt {apartamento.NumeroApartamento}"
            Case Else
                Return $"Notificación de Pago - Torre {apartamento.Torre} Apt {apartamento.NumeroApartamento}"
        End Select
    End Function

    ''' <summary>
    ''' Genera el cuerpo HTML del correo
    ''' </summary>
    Private Shared Function GenerarCuerpoNotificacion(tipo As TipoNotificacion, apartamento As ApartamentoConDeuda) As String
        ' CORRECCIÓN: Inicializar variables con valores por defecto
        Dim colorTipo As String = "#3498db"      ' Azul por defecto
        Dim iconoTipo As String = "📋"           ' Icono por defecto
        Dim mensajeTipo As String = "informarle sobre el estado de sus cuotas de administración" ' Mensaje por defecto

        Select Case tipo
            Case TipoNotificacion.Recordatorio
                colorTipo = "#3498db"
                iconoTipo = "📋"
                mensajeTipo = "recordarle que tiene cuotas de administración pendientes"
            Case TipoNotificacion.PrimerAviso
                colorTipo = "#f39c12"
                iconoTipo = "⚠️"
                mensajeTipo = "informarle que tiene cuotas vencidas que requieren pago inmediato"
            Case TipoNotificacion.SegundoAviso
                colorTipo = "#e67e22"
                iconoTipo = "🚨"
                mensajeTipo = "notificarle URGENTEMENTE sobre cuotas en mora"
            Case TipoNotificacion.AvisoUrgente
                colorTipo = "#e74c3c"
                iconoTipo = "🚨"
                mensajeTipo = "notificarle sobre el estado CRÍTICO de mora de su apartamento"
            Case Else
                ' Ya tienen valores por defecto arriba, pero se pueden redefinir aquí si es necesario
                colorTipo = "#95a5a6"  ' Gris para casos no definidos
                iconoTipo = "📄"
                mensajeTipo = "informarle sobre el estado de sus cuotas de administración"
        End Select

        ' Calcular intereses de mora
        Dim tasaInteres As Decimal = ParametrosDAL.ObtenerTasaInteresMoraActual()
        Dim interesesCalculados As Decimal = apartamento.TotalDeuda * (tasaInteres / 100D) * (apartamento.DiasEnMora / 365D)

        Return $"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background-color: #f4f4f4; }}
        .container {{ max-width: 600px; margin: 0 auto; background-color: white; }}
        .header {{ background-color: {colorTipo}; color: white; padding: 20px; text-align: center; }}
        .content {{ padding: 30px; }}
        .footer {{ background-color: #34495e; color: white; padding: 20px; text-align: center; font-size: 12px; }}
        .highlight {{ background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; border-radius: 5px; margin: 15px 0; }}
        .amount {{ font-size: 24px; font-weight: bold; color: {colorTipo}; }}
        .urgent {{ color: #e74c3c; font-weight: bold; }}
        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
        th {{ background-color: #f8f9fa; }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h1>{iconoTipo} CONJUNTO RESIDENCIAL COOPDIASAM</h1>
            <h2>Notificación de Pago</h2>
        </div>
        
        <div class='content'>
            <p>Estimado(a) <strong>{apartamento.NombreResidente}</strong>,</p>
            
            <p>Nos dirigimos a usted para {mensajeTipo}:</p>
            
            <div class='highlight'>
                <h3>📍 Información del Apartamento:</h3>
                <p><strong>Torre:</strong> {apartamento.Torre}</p>
                <p><strong>Apartamento:</strong> {apartamento.NumeroApartamento}</p>
                <p><strong>Propietario:</strong> {apartamento.NombreResidente}</p>
            </div>

            <h3>💰 Detalles de la Deuda:</h3>
            <table>
                <tr>
                    <th>Concepto</th>
                    <th>Cantidad</th>
                    <th>Valor</th>
                </tr>
                <tr>
                    <td>Cuotas Vencidas</td>
                    <td>{apartamento.CuotasVencidas}</td>
                    <td class='amount'>{apartamento.TotalDeuda:C}</td>
                </tr>
                <tr>
                    <td>Intereses de Mora ({tasaInteres}% anual)</td>
                    <td>{apartamento.DiasEnMora} días</td>
                    <td class='amount'>{Math.Round(interesesCalculados, 0):C}</td>
                </tr>
                <tr style='background-color: #fee; font-weight: bold;'>
                    <td>TOTAL A PAGAR</td>
                    <td></td>
                    <td class='amount urgent'>{apartamento.TotalDeuda + Math.Round(interesesCalculados, 0):C}</td>
                </tr>
            </table>

            <div class='highlight'>
                <h4>📅 Información Importante:</h4>
                <p><strong>Fecha de vencimiento más antigua:</strong> {apartamento.FechaMasAntigua:dd/MM/yyyy}</p>
                <p><strong>Días en mora:</strong> <span class='urgent'>{apartamento.DiasEnMora} días</span></p>
                <p><strong>Fecha de vencimiento más reciente:</strong> {apartamento.FechaMasReciente:dd/MM/yyyy}</p>
            </div>

            <h3>💳 Formas de Pago:</h3>
            <ul>
                <li><strong>Transferencia Bancaria:</strong> Banco Davivienda - Cuenta Corriente 472-40001054-4</li>
                <li><strong>Consignación:</strong> En cualquier oficina Davivienda</li>
                <li><strong>Pago en Efectivo:</strong> Directamente en la administración</li>
                <li><strong>PSE:</strong> A través de la página web del banco</li>
            </ul>

            <p><strong>⚠️ Importante:</strong> Envíe el comprobante de pago por WhatsApp al 321-9597100 o correo electrónico.</p>

            <div class='highlight'>
                <h4>📞 Contacto:</h4>
                <p><strong>Fernando Gamba - Administrador</strong></p>
                <p><strong>Teléfono:</strong> +57 321-9597100</p>
                <p><strong>Email:</strong> correomensajeriablacklines@gmail.com</p>
                <p><strong>Horario de Atención:</strong> Lunes a Viernes 8:00 AM - 5:00 PM</p>
            </div>
        </div>
        
        <div class='footer'>
            <p>Este es un mensaje automático del Sistema de Gestión COOPDIASAM</p>
            <p>Por favor no responda a este correo electrónico</p>
            <p>Fecha de envío: {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>
        </div>
    </div>
</body>
</html>"
    End Function

    ''' <summary>
    ''' Envía correo de notificación usando EmailService
    ''' </summary>
    Private Shared Function EnviarCorreoNotificacion(destinatario As String, nombreDestinatario As String, asunto As String, cuerpoHtml As String) As Boolean
        Try
            Using client As New SmtpClient("smtp.gmail.com")
                client.Port = 587
                client.EnableSsl = True
                client.UseDefaultCredentials = False
                client.Credentials = New System.Net.NetworkCredential("correomensajeriablacklines@gmail.com", "trwp azlx qehm gtby")
                client.Timeout = 30000

                Using mail As New MailMessage()
                    mail.From = New MailAddress("correomensajeriablacklines@gmail.com", "Administración COOPDIASAM")
                    mail.To.Add(destinatario)
                    mail.Subject = asunto
                    mail.Body = cuerpoHtml
                    mail.IsBodyHtml = True
                    mail.Priority = MailPriority.High

                    client.Send(mail)
                End Using
            End Using

            ' Registrar notificación enviada
            RegistrarNotificacionEnviada(destinatario, asunto)
            Return True

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al enviar notificación: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Registra en base de datos que se envió una notificación
    ''' </summary>
    Private Shared Sub RegistrarNotificacionEnviada(destinatario As String, asunto As String)
        Try
            ConexionBD.RegistrarActividad(
                "Sistema",
                "notificaciones",
                0,
                "INSERT",
                $"Notificación enviada a {destinatario}: {asunto}"
            )
        Catch ex As Exception
            ' Error silencioso
        End Try
    End Sub

    ''' <summary>
    ''' Programa notificaciones automáticas
    ''' </summary>
    Public Shared Sub ProgramarNotificacionesAutomaticas()
        Dim timer As New Timer With {
            .Interval = 24 * 60 * 60 * 1000 ' 24 horas
        }

        AddHandler timer.Tick, Async Sub()
                                   ' Ejecutar solo en días laborales y a las 9:00 AM
                                   If DateTime.Now.DayOfWeek <> DayOfWeek.Saturday AndAlso
                                      DateTime.Now.DayOfWeek <> DayOfWeek.Sunday AndAlso
                                      DateTime.Now.Hour = 9 Then

                                       Dim resultado = Await EnviarNotificacionesPagosVencidos()
                                       System.Diagnostics.Debug.WriteLine($"Notificaciones automáticas: {resultado.Mensaje}")
                                   End If
                               End Sub

        timer.Start()
    End Sub

    ''' <summary>
    ''' Envía notificación de bienvenida a nuevos propietarios
    ''' </summary>
    Public Shared Function EnviarNotificacionBienvenida(apartamento As Apartamento) As Boolean
        If String.IsNullOrEmpty(apartamento.Correo) Then
            Return False
        End If

        Try
            Dim asunto As String = $"Bienvenido a COOPDIASAM - Torre {apartamento.Torre} Apt {apartamento.NumeroApartamento}"
            Dim cuerpo As String = GenerarCorreoBienvenida(apartamento)

            Return EnviarCorreoNotificacion(apartamento.Correo, apartamento.NombreResidente, asunto, cuerpo)

        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Shared Function GenerarCorreoBienvenida(apartamento As Apartamento) As String
        Return $"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background-color: #f4f4f4; }}
        .container {{ max-width: 600px; margin: 0 auto; background-color: white; }}
        .header {{ background-color: #27ae60; color: white; padding: 20px; text-align: center; }}
        .content {{ padding: 30px; }}
        .footer {{ background-color: #34495e; color: white; padding: 20px; text-align: center; font-size: 12px; }}
        .welcome {{ background-color: #d5f4e6; border: 1px solid #27ae60; padding: 20px; border-radius: 5px; margin: 20px 0; }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h1>🏠 BIENVENIDO A COOPDIASAM</h1>
        </div>
        
        <div class='content'>
            <div class='welcome'>
                <h2>¡Bienvenido(a) {apartamento.NombreResidente}!</h2>
                <p>Es un placer tenerle como propietario en nuestro conjunto residencial.</p>
            </div>

            <h3>📍 Su Apartamento:</h3>
            <p><strong>Torre:</strong> {apartamento.Torre}</p>
            <p><strong>Apartamento:</strong> {apartamento.NumeroApartamento}</p>
            <p><strong>Código:</strong> {apartamento.ObtenerCodigoApartamento()}</p>

            <h3>📋 Información Importante:</h3>
            <ul>
                <li>Las cuotas de administración vencen el día 10 de cada mes</li>
                <li>Mantenga actualizada su información de contacto</li>
                <li>Guarde los comprobantes de pago</li>
                <li>Consulte regularmente su estado de cuenta</li>
            </ul>

            <h3>📞 Contacto de Administración:</h3>
            <p><strong>Fernando Gamba - Administrador</strong></p>
            <p><strong>Teléfono:</strong> +57 321-9597100</p>
            <p><strong>Email:</strong> correomensajeriablacklines@gmail.com</p>
        </div>
        
        <div class='footer'>
            <p>Conjunto Residencial COOPDIASAM</p>
            <p>Sistema de Gestión Administrativa</p>
        </div>
    </div>
</body>
</html>"
    End Function

End Class

' ============================================================================
' CLASES DE APOYO
' ============================================================================

Public Enum TipoNotificacion
    Recordatorio
    PrimerAviso
    SegundoAviso
    AvisoUrgente
End Enum

Public Class ApartamentoConDeuda
    Public Property IdApartamento As Integer
    Public Property Torre As Integer
    Public Property NumeroApartamento As String
    Public Property NombreResidente As String
    Public Property Correo As String
    Public Property Telefono As String
    Public Property CuotasVencidas As Integer
    Public Property TotalDeuda As Decimal
    Public Property DiasEnMora As Integer
    Public Property FechaMasAntigua As DateTime
    Public Property FechaMasReciente As DateTime
End Class

Public Class ResultadoNotificaciones
    Public Property EsExitoso As Boolean
    Public Property Mensaje As String
    Public Property NotificacionesEnviadas As Integer
    Public Property NotificacionesFallidas As Integer

    Public Sub New()
        EsExitoso = False
        Mensaje = ""
        NotificacionesEnviadas = 0
        NotificacionesFallidas = 0
    End Sub
End Class