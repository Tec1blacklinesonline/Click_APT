Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class ReciboPDF
    Public Shared Function GenerarReciboDesdeFormulario(pago As PagoModel, apartamento As Apartamento) As String
        Try
            ' Crear directorio si no existe
            Dim directorio As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Recibos")
            If Not Directory.Exists(directorio) Then
                Directory.CreateDirectory(directorio)
            End If

            ' Nombre del archivo
            Dim nombreArchivo As String = $"Recibo_{pago.NumeroRecibo}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
            Dim rutaCompleta As String = Path.Combine(directorio, nombreArchivo)

            ' Crear el documento PDF
            Using fs As New FileStream(rutaCompleta, FileMode.Create)
                Dim document As New Document(PageSize.LETTER, 40, 40, 40, 40)
                Dim writer As PdfWriter = PdfWriter.GetInstance(document, fs)

                document.Open()

                ' Fuentes
                Dim fuenteTitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.BLACK)
                Dim fuenteSubtitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.BLACK)
                Dim fuenteNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.BLACK)
                Dim fuentePequena As Font = FontFactory.GetFont(FontFactory.HELVETICA, 8, BaseColor.GRAY)

                ' Encabezado
                Dim titulo As New Paragraph("CONJUNTO RESIDENCIAL", fuenteTitulo)
                titulo.Alignment = Element.ALIGN_CENTER
                document.Add(titulo)

                Dim subtitulo As New Paragraph("RECIBO DE PAGO", fuenteSubtitulo)
                subtitulo.Alignment = Element.ALIGN_CENTER
                subtitulo.SpacingAfter = 20
                document.Add(subtitulo)

                ' Información del recibo
                Dim tablaInfo As New PdfPTable(2)
                tablaInfo.WidthPercentage = 100
                tablaInfo.SetWidths(New Single() {1, 1})

                ' Celda izquierda
                Dim celdaIzq As New PdfPCell()
                celdaIzq.Border = Rectangle.NO_BORDER
                celdaIzq.AddElement(New Paragraph($"Recibo No: {pago.NumeroRecibo}", fuenteNormal))
                celdaIzq.AddElement(New Paragraph($"Fecha: {pago.FechaPago:dd/MM/yyyy}", fuenteNormal))
                tablaInfo.AddCell(celdaIzq)

                ' Celda derecha
                Dim celdaDer As New PdfPCell()
                celdaDer.Border = Rectangle.NO_BORDER
                celdaDer.HorizontalAlignment = Element.ALIGN_RIGHT
                celdaDer.AddElement(New Paragraph($"Apartamento: {apartamento.NumeroApartamento}", fuenteNormal))
                celdaDer.AddElement(New Paragraph($"Torre: {apartamento.Torre}", fuenteNormal))
                tablaInfo.AddCell(celdaDer)

                document.Add(tablaInfo)
                document.Add(New Paragraph(" "))

                ' Información del propietario
                Dim tablaProp As New PdfPTable(1)
                tablaProp.WidthPercentage = 100

                Dim celdaProp As New PdfPCell()
                celdaProp.BackgroundColor = New BaseColor(240, 240, 240)
                celdaProp.Padding = 10
                celdaProp.AddElement(New Paragraph("INFORMACIÓN DEL PROPIETARIO", fuenteSubtitulo))
                celdaProp.AddElement(New Paragraph($"Nombre: {apartamento.Propietario}", fuenteNormal))
                celdaProp.AddElement(New Paragraph($"Matrícula Inmobiliaria: {pago.MatriculaInmobiliaria}", fuenteNormal))
                tablaProp.AddCell(celdaProp)

                document.Add(tablaProp)
                document.Add(New Paragraph(" "))

                ' Detalle del pago
                Dim tablaDetalle As New PdfPTable(2)
                tablaDetalle.WidthPercentage = 100
                tablaDetalle.SetWidths(New Single() {3, 1})

                ' Encabezados
                Dim celdaConcepto As New PdfPCell(New Phrase("CONCEPTO", fuenteSubtitulo))
                celdaConcepto.BackgroundColor = New BaseColor(52, 152, 219)
                celdaConcepto.HorizontalAlignment = Element.ALIGN_CENTER
                celdaConcepto.Padding = 8
                tablaDetalle.AddCell(celdaConcepto)

                Dim celdaValor As New PdfPCell(New Phrase("VALOR", fuenteSubtitulo))
                celdaValor.BackgroundColor = New BaseColor(52, 152, 219)
                celdaValor.HorizontalAlignment = Element.ALIGN_CENTER
                celdaValor.Padding = 8
                tablaDetalle.AddCell(celdaValor)

                ' Filas de detalle
                AgregarFilaDetalle(tablaDetalle, "Saldo Anterior", pago.SaldoAnterior, fuenteNormal)
                AgregarFilaDetalle(tablaDetalle, "Cuota Administración", pago.PagoAdministracion, fuenteNormal)
                AgregarFilaDetalle(tablaDetalle, "Intereses de Mora", pago.PagoIntereses, fuenteNormal)
                AgregarFilaDetalle(tablaDetalle, "Total Pagado", pago.TotalPagado, fuenteSubtitulo, True)
                AgregarFilaDetalle(tablaDetalle, "Nuevo Saldo", pago.SaldoActual, fuenteSubtitulo, True)

                document.Add(tablaDetalle)

                ' Observaciones
                If Not String.IsNullOrEmpty(pago.Observaciones) Then
                    document.Add(New Paragraph(" "))
                    Dim tablaObs As New PdfPTable(1)
                    tablaObs.WidthPercentage = 100

                    Dim celdaObs As New PdfPCell()
                    celdaObs.BackgroundColor = New BaseColor(255, 255, 230)
                    celdaObs.Padding = 10
                    celdaObs.AddElement(New Paragraph("OBSERVACIONES", fuenteSubtitulo))
                    celdaObs.AddElement(New Paragraph(pago.Observaciones, fuenteNormal))
                    tablaObs.AddCell(celdaObs)

                    document.Add(tablaObs)
                End If

                ' Pie de página
                document.Add(New Paragraph(" "))
                document.Add(New Paragraph(" "))

                Dim piePagina As New Paragraph("Este documento es un comprobante de pago oficial del Conjunto Residencial", fuentePequena)
                piePagina.Alignment = Element.ALIGN_CENTER
                document.Add(piePagina)

                Dim fecha As New Paragraph($"Generado el {DateTime.Now:dd/MM/yyyy HH:mm:ss}", fuentePequena)
                fecha.Alignment = Element.ALIGN_CENTER
                document.Add(fecha)

                document.Close()
            End Using

            Return rutaCompleta

        Catch ex As Exception
            Throw New Exception($"Error al generar PDF: {ex.Message}")
        End Try
    End Function

    Private Shared Sub AgregarFilaDetalle(tabla As PdfPTable, concepto As String, valor As Decimal, fuente As Font, Optional esTotal As Boolean = False)
        Dim celdaConcepto As New PdfPCell(New Phrase(concepto, fuente))
        Dim celdaValor As New PdfPCell(New Phrase($"${valor:N0}", fuente))

        If esTotal Then
            celdaConcepto.BackgroundColor = New BaseColor(230, 230, 230)
            celdaValor.BackgroundColor = New BaseColor(230, 230, 230)
        End If

        celdaConcepto.Padding = 5
        celdaValor.Padding = 5
        celdaValor.HorizontalAlignment = Element.ALIGN_RIGHT

        tabla.AddCell(celdaConcepto)
        tabla.AddCell(celdaValor)
    End Sub

    Public Shared Function GenerarRecibosMasivos(pagos As List(Of PagoModel), apartamentos As List(Of Apartamento)) As Dictionary(Of Integer, String)
        Dim rutas As New Dictionary(Of Integer, String)

        For Each pago In pagos
            Dim apartamento = apartamentos.FirstOrDefault(Function(a) a.IdApartamento = pago.IdApartamento)
            If apartamento IsNot Nothing Then
                Try
                    Dim ruta = GenerarReciboDesdeFormulario(pago, apartamento)
                    rutas.Add(pago.IdApartamento, ruta)
                Catch ex As Exception
                    ' Registrar el error pero continuar con los demás
                End Try
            End If
        Next

        Return rutas
    End Function

    Public Shared Function GenerarReciboPDF(pago As PagoModel, apartamento As Apartamento, rutaDestino As String) As Boolean
        Try
            ' Crear el documento PDF
            Using fs As New FileStream(rutaDestino, FileMode.Create)
                Dim document As New Document(PageSize.LETTER, 40, 40, 40, 40)
                Dim writer As PdfWriter = PdfWriter.GetInstance(document, fs)

                document.Open()

                ' Fuentes
                Dim fuenteTitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.BLACK)
                Dim fuenteSubtitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.BLACK)
                Dim fuenteNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.BLACK)
                Dim fuentePequena As Font = FontFactory.GetFont(FontFactory.HELVETICA, 8, BaseColor.GRAY)

                ' Encabezado
                Dim titulo As New Paragraph("CONJUNTO RESIDENCIAL", fuenteTitulo)
                titulo.Alignment = Element.ALIGN_CENTER
                document.Add(titulo)

                Dim subtitulo As New Paragraph("RECIBO DE PAGO", fuenteSubtitulo)
                subtitulo.Alignment = Element.ALIGN_CENTER
                subtitulo.SpacingAfter = 20
                document.Add(subtitulo)

                ' Información del recibo
                Dim tablaInfo As New PdfPTable(2)
                tablaInfo.WidthPercentage = 100
                tablaInfo.SetWidths(New Single() {1, 1})

                ' Celda izquierda
                Dim celdaIzq As New PdfPCell()
                celdaIzq.Border = Rectangle.NO_BORDER
                celdaIzq.AddElement(New Paragraph($"Recibo No: {pago.NumeroRecibo}", fuenteNormal))
                celdaIzq.AddElement(New Paragraph($"Fecha: {pago.FechaPago:dd/MM/yyyy}", fuenteNormal))
                tablaInfo.AddCell(celdaIzq)

                ' Celda derecha
                Dim celdaDer As New PdfPCell()
                celdaDer.Border = Rectangle.NO_BORDER
                celdaDer.HorizontalAlignment = Element.ALIGN_RIGHT
                celdaDer.AddElement(New Paragraph($"Apartamento: {apartamento.NumeroApartamento}", fuenteNormal))
                celdaDer.AddElement(New Paragraph($"Torre: {apartamento.Torre}", fuenteNormal))
                tablaInfo.AddCell(celdaDer)

                document.Add(tablaInfo)
                document.Add(New Paragraph(" "))

                ' Información del propietario
                Dim tablaProp As New PdfPTable(1)
                tablaProp.WidthPercentage = 100

                Dim celdaProp As New PdfPCell()
                celdaProp.BackgroundColor = New BaseColor(240, 240, 240)
                celdaProp.Padding = 10
                celdaProp.AddElement(New Paragraph("INFORMACIÓN DEL PROPIETARIO", fuenteSubtitulo))
                celdaProp.AddElement(New Paragraph($"Nombre: {If(String.IsNullOrEmpty(apartamento.Propietario), apartamento.NombreResidente, apartamento.Propietario)}", fuenteNormal))
                celdaProp.AddElement(New Paragraph($"Matrícula Inmobiliaria: {pago.MatriculaInmobiliaria}", fuenteNormal))
                tablaProp.AddCell(celdaProp)

                document.Add(tablaProp)
                document.Add(New Paragraph(" "))

                ' Detalle del pago
                Dim tablaDetalle As New PdfPTable(2)
                tablaDetalle.WidthPercentage = 100
                tablaDetalle.SetWidths(New Single() {3, 1})

                ' Encabezados
                Dim celdaConcepto As New PdfPCell(New Phrase("CONCEPTO", fuenteSubtitulo))
                celdaConcepto.BackgroundColor = New BaseColor(52, 152, 219)
                celdaConcepto.HorizontalAlignment = Element.ALIGN_CENTER
                celdaConcepto.Padding = 8
                tablaDetalle.AddCell(celdaConcepto)

                Dim celdaValor As New PdfPCell(New Phrase("VALOR", fuenteSubtitulo))
                celdaValor.BackgroundColor = New BaseColor(52, 152, 219)
                celdaValor.HorizontalAlignment = Element.ALIGN_CENTER
                celdaValor.Padding = 8
                tablaDetalle.AddCell(celdaValor)

                ' Filas de detalle
                AgregarFilaDetalle(tablaDetalle, "Saldo Anterior", pago.SaldoAnterior, fuenteNormal)
                AgregarFilaDetalle(tablaDetalle, "Cuota Administración", pago.PagoAdministracion, fuenteNormal)
                AgregarFilaDetalle(tablaDetalle, "Intereses de Mora", pago.PagoIntereses, fuenteNormal)
                AgregarFilaDetalle(tablaDetalle, "Total Pagado", pago.TotalPagado, fuenteSubtitulo, True)
                AgregarFilaDetalle(tablaDetalle, "Nuevo Saldo", pago.SaldoActual, fuenteSubtitulo, True)

                document.Add(tablaDetalle)

                ' Observaciones
                If Not String.IsNullOrEmpty(pago.Observaciones) Then
                    document.Add(New Paragraph(" "))
                    Dim tablaObs As New PdfPTable(1)
                    tablaObs.WidthPercentage = 100

                    Dim celdaObs As New PdfPCell()
                    celdaObs.BackgroundColor = New BaseColor(255, 255, 230)
                    celdaObs.Padding = 10
                    celdaObs.AddElement(New Paragraph("OBSERVACIONES", fuenteSubtitulo))
                    celdaObs.AddElement(New Paragraph(pago.Observaciones, fuenteNormal))
                    tablaObs.AddCell(celdaObs)

                    document.Add(tablaObs)
                End If

                ' Pie de página
                document.Add(New Paragraph(" "))
                document.Add(New Paragraph(" "))

                Dim piePagina As New Paragraph("Este documento es un comprobante de pago oficial del Conjunto Residencial", fuentePequena)
                piePagina.Alignment = Element.ALIGN_CENTER
                document.Add(piePagina)

                Dim fecha As New Paragraph($"Generado el {DateTime.Now:dd/MM/yyyy HH:mm:ss}", fuentePequena)
                fecha.Alignment = Element.ALIGN_CENTER
                document.Add(fecha)

                document.Close()
            End Using

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function
End Class