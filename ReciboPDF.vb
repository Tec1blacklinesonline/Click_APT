Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports iTextSharp.text.pdf.draw

Public Class ReciboPDF

    ' Constantes del conjunto según el documento de requerimientos
    Private Const NOMBRE_CONJUNTO As String = "COOPDIASAM"
    Private Const DESCRIPCION_CONJUNTO As String = "Conjunto Habitacional"
    Private Const NIT As String = "900225635-8"
    Private Const DIRECCION As String = "Barrio Villa Café - Ibagué Tolima"
    Private Const EMAIL_CONJUNTO As String = "conjuntocoopdiasam@yahoo.es"
    Private Const ADMINISTRADOR_NOMBRE As String = "Fernando Gamba"
    Private Const ADMINISTRADOR_TELEFONO As String = "+57 321 9597100"
    Private Const ADMINISTRADOR_CARGO As String = "Administrador CONJUNTO RESIDENCIAL COOPDIASAM"

    ' Método principal para generar el PDF del recibo
    Public Shared Function GenerarReciboPDF(pago As PagoModel, apartamento As Apartamento, rutaArchivo As String) As Boolean
        Try
            ' Crear el documento PDF
            Dim documento As New Document(PageSize.A4, 40, 40, 40, 40)
            Dim writer As PdfWriter = PdfWriter.GetInstance(documento, New FileStream(rutaArchivo, FileMode.Create))

            documento.Open()

            ' Agregar contenido siguiendo el formato exacto del documento
            AgregarEncabezadoYDatos(documento, apartamento, pago)
            AgregarNumeroRecibo(documento, pago)
            AgregarDetallePagos(documento, pago)
            AgregarObservaciones(documento, pago)
            AgregarFirmaAdministrador(documento)

            documento.Close()
            Return True

        Catch ex As Exception
            Throw New Exception($"Error al generar PDF: {ex.Message}")
        End Try
    End Function

    Private Shared Sub AgregarEncabezadoYDatos(documento As Document, apartamento As Apartamento, pago As PagoModel)
        ' Fuentes
        Dim fontTitulo As New Font(Font.FontFamily.HELVETICA, 14, Font.BOLD)
        Dim fontSubtitulo As New Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD)
        Dim fontPequeña As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL)

        ' Línea separadora superior
        Dim lineaSeparadora As New LineSeparator(1, 100, BaseColor.BLACK, Element.ALIGN_CENTER, 0)
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))

        ' 1. Encabezado y Datos de Identificación
        Dim parrafoEncabezado As New Paragraph()
        parrafoEncabezado.Add(New Chunk($"Nombre de la organización: {NOMBRE_CONJUNTO} - {DESCRIPCION_CONJUNTO}", fontBold))
        parrafoEncabezado.Add(Chunk.NEWLINE)
        parrafoEncabezado.Add(New Chunk($"NIT: {FormatearNIT(NIT)}", fontNormal))
        parrafoEncabezado.Add(Chunk.NEWLINE)
        parrafoEncabezado.Add(New Chunk($"Fecha de impresión: {DateTime.Now:dd-MMM-yyyy HH:mm}", fontNormal))
        parrafoEncabezado.Add(Chunk.NEWLINE)
        parrafoEncabezado.Add(New Chunk($"Correo electrónico para copia del recibo: {apartamento.Correo}", fontNormal))

        documento.Add(parrafoEncabezado)
        documento.Add(New Paragraph(" "))
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))

        ' 2. Datos del Residente
        Dim parrafoResidente As New Paragraph()
        parrafoResidente.Add(New Chunk("2. Datos del Residente", fontBold))
        parrafoResidente.Add(Chunk.NEWLINE)
        parrafoResidente.Add(Chunk.NEWLINE)
        parrafoResidente.Add(New Chunk($"Nombre: {apartamento.NombreResidente.ToUpper()}", fontNormal))
        parrafoResidente.Add(Chunk.NEWLINE)
        parrafoResidente.Add(New Chunk($"Apartamento: APT - {apartamento.Torre}{apartamento.NumeroApartamento}", fontNormal))
        parrafoResidente.Add(Chunk.NEWLINE)
        parrafoResidente.Add(New Chunk($"Ficha catastral: {ObtenerFichaCatastral(apartamento, pago)}", fontNormal))
        parrafoResidente.Add(Chunk.NEWLINE)
        parrafoResidente.Add(New Chunk($"Celular: {apartamento.Telefono}", fontNormal))
        parrafoResidente.Add(Chunk.NEWLINE)
        parrafoResidente.Add(New Chunk($"Correo electrónico: {apartamento.Correo}", fontNormal))

        documento.Add(parrafoResidente)
        documento.Add(New Paragraph(" "))
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))

        ' 3. Datos del Recibo
        Dim parrafoRecibo As New Paragraph()
        parrafoRecibo.Add(New Chunk("3. Datos del Recibo", fontBold))
        parrafoRecibo.Add(Chunk.NEWLINE)
        parrafoRecibo.Add(Chunk.NEWLINE)
        parrafoRecibo.Add(New Chunk($"Número del recibo: {pago.NumeroRecibo}", fontNormal))
        parrafoRecibo.Add(Chunk.NEWLINE)
        parrafoRecibo.Add(New Chunk($"Estado final de cuenta: {FormatearMoneda(pago.SaldoActual)}", fontNormal))
        parrafoRecibo.Add(Chunk.NEWLINE)
        parrafoRecibo.Add(New Chunk($"Fecha del último pago: {pago.FechaPago:dd-MMM-yyyy}", fontNormal))

        documento.Add(parrafoRecibo)
        documento.Add(New Paragraph(" "))
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarNumeroRecibo(documento As Document, pago As PagoModel)
        Dim fontGrande As New Font(Font.FontFamily.HELVETICA, 16, Font.BOLD)

        ' Agregar RECIBO DE CAJA No. centrado
        Dim parrafoNumero As New Paragraph()
        parrafoNumero.Add(New Chunk($"RECIBO DE CAJA No. {pago.NumeroRecibo}", fontGrande))
        parrafoNumero.Alignment = Element.ALIGN_CENTER

        documento.Add(parrafoNumero)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarDetallePagos(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)
        Dim lineaSeparadora As New LineSeparator(1, 100, BaseColor.BLACK, Element.ALIGN_CENTER, 0)

        ' 4. Detalle de Pagos
        Dim tablaDetalle As New PdfPTable(2)
        tablaDetalle.WidthPercentage = 100
        tablaDetalle.SetWidths({50, 50})

        ' Encabezados
        Dim celdaConcepto As New PdfPCell(New Phrase("Concepto", fontBold))
        celdaConcepto.HorizontalAlignment = Element.ALIGN_LEFT
        celdaConcepto.Border = Rectangle.NO_BORDER
        celdaConcepto.PaddingBottom = 5

        Dim celdaValor As New PdfPCell(New Phrase("Valor", fontBold))
        celdaValor.HorizontalAlignment = Element.ALIGN_RIGHT
        celdaValor.Border = Rectangle.NO_BORDER
        celdaValor.PaddingBottom = 5

        tablaDetalle.AddCell(celdaConcepto)
        tablaDetalle.AddCell(celdaValor)

        ' Datos
        AgregarFilaTabla(tablaDetalle, "Saldo Anterior", FormatearMoneda(pago.SaldoAnterior), fontNormal)
        AgregarFilaTabla(tablaDetalle, "V/R Pagado Administración", FormatearMoneda(pago.PagoAdministracion), fontNormal)
        AgregarFilaTabla(tablaDetalle, "V/R Pagado Intereses", FormatearMoneda(pago.PagoIntereses), fontNormal)
        AgregarFilaTabla(tablaDetalle, "Total Pagado", FormatearMoneda(pago.TotalPagado), fontBold)

        documento.Add(New Paragraph("4. Detalle de Pagos", fontBold))
        documento.Add(New Paragraph(" "))
        documento.Add(tablaDetalle)
        documento.Add(New Paragraph(" "))
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))

        ' 5. Saldos
        Dim tablaSaldos As New PdfPTable(2)
        tablaSaldos.WidthPercentage = 100
        tablaSaldos.SetWidths({50, 50})

        ' Encabezados
        tablaSaldos.AddCell(celdaConcepto)
        tablaSaldos.AddCell(celdaValor)

        ' Datos
        AgregarFilaTabla(tablaSaldos, "Interés", FormatearMoneda(pago.PagoIntereses), fontNormal)
        AgregarFilaTabla(tablaSaldos, "Cuota", FormatearMoneda(pago.CuotaActual), fontNormal)
        AgregarFilaTabla(tablaSaldos, "Total", FormatearMoneda(pago.CuotaActual + pago.PagoIntereses), fontBold)

        documento.Add(New Paragraph("5. Saldos", fontBold))
        documento.Add(New Paragraph(" "))
        documento.Add(tablaSaldos)
        documento.Add(New Paragraph(" "))
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarFilaTabla(tabla As PdfPTable, concepto As String, valor As String, fuente As Font)
        Dim celdaConcepto As New PdfPCell(New Phrase(concepto, fuente))
        celdaConcepto.Border = Rectangle.NO_BORDER
        celdaConcepto.HorizontalAlignment = Element.ALIGN_LEFT

        Dim celdaValor As New PdfPCell(New Phrase(valor, fuente))
        celdaValor.Border = Rectangle.NO_BORDER
        celdaValor.HorizontalAlignment = Element.ALIGN_RIGHT

        tabla.AddCell(celdaConcepto)
        tabla.AddCell(celdaValor)
    End Sub

    Private Shared Sub AgregarObservaciones(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)
        Dim lineaSeparadora As New LineSeparator(1, 100, BaseColor.BLACK, Element.ALIGN_CENTER, 0)

        ' 6. Observación
        Dim parrafoObs As New Paragraph()
        parrafoObs.Add(New Chunk("6. Observación", fontBold))
        parrafoObs.Add(Chunk.NEWLINE)
        parrafoObs.Add(Chunk.NEWLINE)

        If Not String.IsNullOrEmpty(pago.Observaciones) Then
            parrafoObs.Add(New Chunk(pago.Observaciones, fontNormal))
        Else
            parrafoObs.Add(New Chunk("(Campo presente pero sin contenido en este caso)", New Font(Font.FontFamily.HELVETICA, 9, Font.ITALIC)))
        End If

        documento.Add(parrafoObs)
        documento.Add(New Paragraph(" "))
        documento.Add(lineaSeparadora)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarFirmaAdministrador(documento As Document)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)
        Dim fontFirma As New Font(Font.FontFamily.HELVETICA, 12, Font.BOLDITALIC)

        ' 7. Firma y Responsable
        Dim parrafoFirma As New Paragraph()
        parrafoFirma.Add(New Chunk("7. Firma y Responsable", fontBold))
        parrafoFirma.Add(Chunk.NEWLINE)
        parrafoFirma.Add(Chunk.NEWLINE)

        ' Espacio para la firma
        parrafoFirma.Add(Chunk.NEWLINE)
        parrafoFirma.Add(Chunk.NEWLINE)
        parrafoFirma.Add(New Chunk("_____________________________", fontNormal))
        parrafoFirma.Add(Chunk.NEWLINE)

        parrafoFirma.Add(New Chunk($"Firmado por: {ADMINISTRADOR_NOMBRE}", fontNormal))
        parrafoFirma.Add(Chunk.NEWLINE)
        parrafoFirma.Add(New Chunk($"Teléfono: {ADMINISTRADOR_TELEFONO}", fontNormal))
        parrafoFirma.Add(Chunk.NEWLINE)
        parrafoFirma.Add(New Chunk($"Cargo: {ADMINISTRADOR_CARGO}", fontNormal))
        parrafoFirma.Add(Chunk.NEWLINE)
        parrafoFirma.Add(New Chunk("Firma escaneada incluida", New Font(Font.FontFamily.HELVETICA, 8, Font.ITALIC)))

        documento.Add(parrafoFirma)
    End Sub

    ' Métodos auxiliares
    Private Shared Function FormatearMoneda(valor As Decimal) As String
        Return $"${valor:N0}"
    End Function

    Private Shared Function FormatearNIT(nit As String) As String
        ' Formatear NIT con puntos y guión
        If nit.Length >= 10 Then
            Return $"{nit.Substring(0, 3)} {nit.Substring(3, 3)} {nit.Substring(6, 3)} - {nit.Substring(9)}"
        End If
        Return nit
    End Function

    Private Shared Function ObtenerFichaCatastral(apartamento As Apartamento, pago As PagoModel) As String
        ' Según el ejemplo del documento: 20129-2192102
        ' Si no hay matrícula, generar una basada en torre y apartamento
        If Not String.IsNullOrEmpty(pago.MatriculaInmobiliaria) Then
            Return pago.MatriculaInmobiliaria
        End If
        Return $"20129-21921{apartamento.Torre:00}{apartamento.NumeroApartamento}"
    End Function

    ' Método para generar y guardar el PDF
    Public Shared Function GenerarYGuardarRecibo(pago As PagoModel, apartamento As Apartamento, carpetaDestino As String) As String
        Try
            ' Validar que la carpeta existe
            If Not Directory.Exists(carpetaDestino) Then
                Directory.CreateDirectory(carpetaDestino)
            End If

            ' Crear nombre de archivo único
            Dim nombreArchivo As String = $"Recibo_{pago.NumeroRecibo}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
            Dim rutaCompleta As String = Path.Combine(carpetaDestino, nombreArchivo)

            ' Generar el PDF
            If GenerarReciboPDF(pago, apartamento, rutaCompleta) Then
                Return rutaCompleta
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Throw New Exception($"Error al generar y guardar recibo: {ex.Message}")
        End Try
    End Function

    ' Método para enviar el recibo por correo
    Public Shared Function EnviarReciboPorCorreo(pago As PagoModel, apartamento As Apartamento, rutaArchivoPDF As String) As Boolean
        Try
            ' Validar que el apartamento tenga correo
            If String.IsNullOrEmpty(apartamento.Correo) Then
                Throw New Exception("No se encuentra el correo en la base de datos")
            End If

            ' Configurar el mensaje
            Dim mensaje As New MailMessage()
            mensaje.From = New MailAddress(EMAIL_CONJUNTO, NOMBRE_CONJUNTO)
            mensaje.To.Add(New MailAddress(apartamento.Correo))
            mensaje.Subject = $"Recibo de Pago - {NOMBRE_CONJUNTO} - Apartamento {apartamento.Torre}{apartamento.NumeroApartamento}"

            ' Cuerpo del mensaje
            Dim cuerpoHtml As String = $"
                <html>
                <body style='font-family: Arial, sans-serif;'>
                    <h2>Estimado(a) {apartamento.NombreResidente}</h2>
                    <p>Adjunto encontrará el recibo de pago correspondiente a su apartamento <strong>{apartamento.Torre}{apartamento.NumeroApartamento}</strong>.</p>
                    <br>
                    <p><strong>Detalles del pago:</strong></p>
                    <ul>
                        <li>Número de recibo: {pago.NumeroRecibo}</li>
                        <li>Fecha de pago: {pago.FechaPago:dd/MM/yyyy}</li>
                        <li>Valor pagado: {FormatearMoneda(pago.TotalPagado)}</li>
                        <li>Saldo actual: {FormatearMoneda(pago.SaldoActual)}</li>
                    </ul>
                    <br>
                    <p>Agradecemos su puntualidad en el pago.</p>
                    <br>
                    <p>Cordialmente,</p>
                    <p><strong>{ADMINISTRADOR_NOMBRE}</strong><br>
                    {ADMINISTRADOR_CARGO}<br>
                    Tel: {ADMINISTRADOR_TELEFONO}</p>
                </body>
                </html>"

            mensaje.Body = cuerpoHtml
            mensaje.IsBodyHtml = True

            ' Adjuntar el PDF
            If File.Exists(rutaArchivoPDF) Then
                Dim adjunto As New Attachment(rutaArchivoPDF)
                mensaje.Attachments.Add(adjunto)
            End If

            ' Configurar el cliente SMTP (ejemplo con Gmail, ajustar según necesidad)
            Dim cliente As New SmtpClient("smtp.gmail.com", 587)
            cliente.EnableSsl = True
            cliente.Credentials = New NetworkCredential("tu_correo@gmail.com", "tu_contraseña_aplicacion")

            ' Enviar el mensaje
            cliente.Send(mensaje)

            Return True

        Catch ex As Exception
            Throw New Exception($"Error al enviar correo: {ex.Message}")
        End Try
    End Function

End Class