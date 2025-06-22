' ============================================================================
' EMAIL SERVICE EXTRA ADAPTADO - COMPATIBLE CON TU ESTRUCTURA EXISTENTE
' ✅ Mantiene la compatibilidad con tu EmailServiceExtra.vb
' ✅ Agrega funcionalidad de envío masivo mejorado
' ============================================================================

Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.IO
Imports System.Threading.Tasks
Imports System.Threading

Public Class EmailServiceExtraAdaptado
    Inherits EmailServiceExtra

    ' Configuración para envío masivo (misma configuración SMTP que tienes)
    Private Const MaxEmailsPorConexion As Integer = 8
    Private Const DelayEntreEmails As Integer = 2000 ' 2 segundos
    Private Const DelayEntreConexiones As Integer = 6000 ' 6 segundos
    Private Const MaxReintentos As Integer = 3

    ''' <summary>
    ''' Envía recibos de pago extra de forma masiva con control de límites de Gmail
    ''' </summary>
    Public Shared Async Function EnviarRecibosMasivosMejorado(
    recibosData As List(Of DatosEnvioRecibo),
    Optional progreso As IProgress(Of ProgressInfo) = Nothing,
    Optional cancellationToken As CancellationToken = Nothing
) As Task(Of ResultadoEnvioMasivo)

        Dim resultado As New ResultadoEnvioMasivo()

        Try
            ' Validar datos de entrada
            If recibosData Is Nothing OrElse recibosData.Count = 0 Then
                resultado.Mensaje = "No hay recibos para enviar"
                Return resultado
            End If

            ' Filtrar solo recibos válidos
            Dim recibosValidosList As New List(Of DatosEnvioRecibo)
            For Each r In recibosData
                If Not String.IsNullOrWhiteSpace(r.CorreoDestino) AndAlso
               File.Exists(r.RutaPDF) AndAlso
               Not String.IsNullOrWhiteSpace(r.NumeroRecibo) Then

                    recibosValidosList.Add(r)
                End If
            Next

            If recibosValidosList.Count = 0 Then
                resultado.Mensaje = "No hay recibos válidos para enviar (verifique correos y archivos PDF)"
                Return resultado
            End If

            resultado.TotalRecibos = recibosValidosList.Count

            progreso?.Report(New ProgressInfo With {
            .Mensaje = $"🚀 Iniciando envío masivo de {recibosValidosList.Count} recibos de pagos extra...",
            .Progreso = 0
        })

            ' Dividir en lotes para Gmail
            Dim lotes = DividirEnLotes(recibosValidosList, MaxEmailsPorConexion)
            Dim emailsEnviados As Integer = 0

            For i As Integer = 0 To lotes.Count - 1
                If cancellationToken.IsCancellationRequested Then
                    resultado.Mensaje = "❌ Envío cancelado por el usuario"
                    Exit For
                End If

                progreso?.Report(New ProgressInfo With {
                .Mensaje = $"📧 Procesando lote {i + 1} de {lotes.Count} ({lotes(i).Count} correos)...",
                .Progreso = CInt((emailsEnviados / resultado.TotalRecibos) * 100)
            })

                ' Procesar lote con reintentos
                Dim resultadoLote = Await ProcesarLoteConReintentos(
                lotes(i), progreso, emailsEnviados, resultado.TotalRecibos, cancellationToken)

                resultado.EmailsExitosos += resultadoLote.Exitosos
                resultado.EmailsConError += resultadoLote.ConError
                resultado.ErroresDetallados.AddRange(resultadoLote.Errores)

                emailsEnviados += lotes(i).Count

                ' Pausa entre lotes (excepto el último)
                If i < lotes.Count - 1 AndAlso Not cancellationToken.IsCancellationRequested Then
                    progreso?.Report(New ProgressInfo With {
                    .Mensaje = $"⏳ Pausa de seguridad entre lotes... ({emailsEnviados}/{resultado.TotalRecibos})",
                    .Progreso = CInt((emailsEnviados / resultado.TotalRecibos) * 100)
                })

                    Await Task.Delay(DelayEntreConexiones, cancellationToken)
                End If
            Next

            ' Resultado final
            resultado.Exitoso = (resultado.EmailsExitosos > 0)

            If cancellationToken.IsCancellationRequested Then
                resultado.Mensaje = $"⚠️ Proceso cancelado: {resultado.EmailsExitosos} enviados, {resultado.EmailsConError} pendientes"
            Else
                resultado.Mensaje = $"✅ Proceso completado: {resultado.EmailsExitosos} exitosos, {resultado.EmailsConError} con errores"
            End If

        Catch ex As Exception
            resultado.Mensaje = $"❌ Error crítico en envío masivo: {ex.Message}"
            resultado.ErroresDetallados.Add($"Error crítico: {ex.Message}")
        End Try

        Return resultado
    End Function



    ' Procesa un lote de correos con reintentos automáticos
    Private Shared Async Function ProcesarLoteConReintentos(
    lote As List(Of DatosEnvioRecibo),
    progreso As IProgress(Of ProgressInfo),
    emailsEnviadosAntes As Integer,
    totalEmails As Integer,
    cancellationToken As CancellationToken
) As Task(Of ResultadoLote)

        Dim resultado As New ResultadoLote()

        For intento = 1 To MaxReintentos
            If cancellationToken.IsCancellationRequested Then
                Exit For
            End If

            Dim huboErrorSMTP As Boolean = False
            Dim smtpExGuardado As Exception = Nothing

            Try
                ' Crear conexión SMTP
                Using client As New SmtpClient("smtp.gmail.com")
                    ConfigurarClienteSMTP(client)

                    For i = 0 To lote.Count - 1
                        If cancellationToken.IsCancellationRequested Then
                            Exit For
                        End If

                        Try
                            Dim recibo = lote(i)
                            Dim emailsCompletados = emailsEnviadosAntes + i + 1

                            progreso?.Report(New ProgressInfo With {
                            .Mensaje = $"📧 Enviando a {recibo.NombreDestino} ({recibo.Apartamento})... ({emailsCompletados}/{totalEmails})",
                            .Progreso = CInt((emailsCompletados / totalEmails) * 100)
                        })

                            ' Enviar email individual usando tu método existente
                            Dim exitoso = EnviarEmailIndividualAdaptado(client, recibo)

                            If exitoso Then
                                resultado.Exitosos += 1
                                progreso?.Report(New ProgressInfo With {
                                .Mensaje = $"✅ Enviado a {recibo.NombreDestino} - {recibo.NumeroRecibo}",
                                .Progreso = CInt((emailsCompletados / totalEmails) * 100)
                            })
                            Else
                                resultado.ConError += 1
                                resultado.Errores.Add($"❌ {recibo.NombreDestino} ({recibo.Apartamento}): Error en el envío")
                            End If

                            ' Pausa entre emails para evitar límites de Gmail
                            If i < lote.Count - 1 AndAlso Not cancellationToken.IsCancellationRequested Then
                                Await Task.Delay(DelayEntreEmails, cancellationToken)
                            End If

                        Catch emailEx As Exception
                            resultado.ConError += 1
                            resultado.Errores.Add($"❌ {lote(i).NombreDestino}: {emailEx.Message}")
                        End Try
                    Next
                End Using

                ' Si llegamos aquí sin errores SMTP, el lote se procesó correctamente
                Exit For

            Catch smtpEx As Exception
                progreso?.Report(New ProgressInfo With {
                .Mensaje = $"⚠️ Error SMTP en lote, reintento {intento}/{MaxReintentos}... {smtpEx.Message}",
                .Progreso = CInt(((emailsEnviadosAntes) / totalEmails) * 100)
            })

                smtpExGuardado = smtpEx
                huboErrorSMTP = True

                If intento = MaxReintentos Then
                    ' Último intento fallido, marcar pendientes como error
                    For Each recibo In lote
                        If Not resultado.Errores.Any(Function(e) e.Contains(recibo.NombreDestino)) Then
                            resultado.ConError += 1
                            resultado.Errores.Add($"❌ {recibo.NombreDestino}: Error SMTP tras {MaxReintentos} intentos - {smtpEx.Message}")
                        End If
                    Next
                End If
            End Try

            ' Esperar antes del siguiente intento (fuera del Catch)
            If huboErrorSMTP AndAlso intento < MaxReintentos AndAlso Not cancellationToken.IsCancellationRequested Then
                Await Task.Delay(DelayEntreConexiones * intento, cancellationToken)
            End If
        Next

        Return resultado
    End Function


    ''' <summary>
    ''' Envía un email individual usando tu método existente adaptado
    ''' </summary>
    Private Shared Function EnviarEmailIndividualAdaptado(client As SmtpClient, datosRecibo As DatosEnvioRecibo) As Boolean
        Try
            ' Usar tu método existente de EmailServiceExtra con algunas mejoras
            Return EmailServiceExtra.EnviarReciboPagoExtra(
                datosRecibo.CorreoDestino,
                datosRecibo.NombreDestino,
                datosRecibo.NumeroRecibo,
                datosRecibo.TipoPago,
                datosRecibo.RutaPDF
            )

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error enviando email individual: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Configura el cliente SMTP (usar la misma configuración que tienes)
    ''' </summary>
    Private Shared Sub ConfigurarClienteSMTP(client As SmtpClient)
        client.Port = 587
        client.EnableSsl = True
        client.DeliveryMethod = SmtpDeliveryMethod.Network
        client.UseDefaultCredentials = False
        client.Credentials = New System.Net.NetworkCredential("correomensajeriablacklines@gmail.com", "trwp azlx qehm gtby")
        client.Timeout = 45000 ' Timeout más largo para lotes
    End Sub

    ''' <summary>
    ''' Divide la lista de recibos en lotes más pequeños
    ''' </summary>
    Private Shared Function DividirEnLotes(recibos As List(Of DatosEnvioRecibo), tamanoLote As Integer) As List(Of List(Of DatosEnvioRecibo))
        Dim lotes As New List(Of List(Of DatosEnvioRecibo))

        For i = 0 To recibos.Count - 1 Step tamanoLote
            Dim lote = recibos.Skip(i).Take(tamanoLote).ToList()
            lotes.Add(lote)
        Next

        Return lotes
    End Function

End Class

' ============================================================================
' CLASES DE SOPORTE ADAPTADAS A TU ESTRUCTURA
' ============================================================================

''' <summary>
''' Datos para envío de recibo individual - Adaptado a tu estructura
''' </summary>
Public Class DatosEnvioRecibo
    Public Property CorreoDestino As String
    Public Property NombreDestino As String
    Public Property NumeroRecibo As String
    Public Property TipoPago As String
    Public Property RutaPDF As String
    Public Property Apartamento As String
    Public Property IdApartamento As Integer
End Class

''' <summary>
''' Resultado del envío masivo
''' </summary>
Public Class ResultadoEnvioMasivo
    Public Property Exitoso As Boolean = False
    Public Property TotalRecibos As Integer = 0
    Public Property EmailsExitosos As Integer = 0
    Public Property EmailsConError As Integer = 0
    Public Property Mensaje As String = ""
    Public Property ErroresDetallados As New List(Of String)

    ''' <summary>
    ''' Obtiene un resumen formateado del resultado
    ''' </summary>
    Public ReadOnly Property ResumenFormateado As String
        Get
            Dim resumen As New Text.StringBuilder()
            resumen.AppendLine($"📊 RESUMEN DEL ENVÍO MASIVO")
            resumen.AppendLine($"✅ Exitosos: {EmailsExitosos}")
            resumen.AppendLine($"❌ Con errores: {EmailsConError}")
            resumen.AppendLine($"📄 Total procesados: {TotalRecibos}")

            If EmailsConError > 0 AndAlso ErroresDetallados.Count > 0 Then
                resumen.AppendLine()
                resumen.AppendLine("🔍 ERRORES DETALLADOS:")
                For i = 0 To Math.Min(4, ErroresDetallados.Count - 1)
                    resumen.AppendLine($"• {ErroresDetallados(i)}")
                Next

                If ErroresDetallados.Count > 5 Then
                    resumen.AppendLine($"... y {ErroresDetallados.Count - 5} errores más")
                End If
            End If

            Return resumen.ToString()
        End Get
    End Property
End Class

''' <summary>
''' Resultado de un lote individual
''' </summary>
Public Class ResultadoLote
    Public Property Exitosos As Integer = 0
    Public Property ConError As Integer = 0
    Public Property Errores As New List(Of String)
End Class

''' <summary>
''' Información de progreso para la interfaz
''' </summary>
