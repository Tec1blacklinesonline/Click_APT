' ============================================================================
' RECIBO PDF - VERSIÓN FINAL CORREGIDA SIN ERRORES DE COMPILACIÓN
' ✅ Resuelve conflictos de espacios de nombres
' ✅ Incluye métodos faltantes que usa FormPagos
' ============================================================================

Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
' ✅ REMOVIDO: System.Drawing para evitar conflictos con iTextSharp.text

Public Class ReciboPDF

    ' ✅ AGREGADO: Método faltante GenerarReciboDePagoTemporal
    Public Shared Function GenerarReciboDePagoTemporal(pago As PagoModel, apartamento As Apartamento) As String
        ' Usar el mismo método principal pero con nombre temporal
        Return GenerarReciboDePagoSeguro(pago, apartamento)
    End Function

    ' ✅ AGREGADO: Método faltante GenerarReciboDePagoEspecifico con sobrecarga para 3 parámetros
    Public Shared Function GenerarReciboDePagoEspecifico(pago As PagoModel, apartamento As Apartamento) As String
        ' Usar el mismo método principal pero con nombre específico
        Return GenerarReciboDePagoSeguro(pago, apartamento)
    End Function

    ' ✅ AGREGADO: Sobrecarga para GenerarReciboDePagoEspecifico con ruta personalizada
    Public Shared Function GenerarReciboDePagoEspecifico(pago As PagoModel, apartamento As Apartamento, rutaPersonalizada As String) As String
        Try
            ' Usar la ruta personalizada proporcionada
            Dim rutaCarpeta As String = Path.GetDirectoryName(rutaPersonalizada)
            If Not Directory.Exists(rutaCarpeta) Then
                Directory.CreateDirectory(rutaCarpeta)
            End If

            ' ✅ CORREGIDO: Cálculos matemáticos correctos
            Dim saldoAnterior As Decimal = Math.Abs(pago.SaldoAnterior)
            Dim valorPagado As Decimal = pago.PagoAdministracion + pago.PagoIntereses
            Dim saldoActual As Decimal = saldoAnterior - valorPagado

            If saldoActual < 0 Then
                saldoActual = 0
            End If

            ' Crear documento PDF en la ruta específica
            Dim documento As New Document(PageSize.Letter, 50, 50, 50, 50)
            Dim writer As PdfWriter = PdfWriter.GetInstance(documento, New FileStream(rutaPersonalizada, FileMode.Create))

            documento.Open()

            ' ✅ CORREGIDO: Fuentes especificando explícitamente iTextSharp.text.Font
            Dim fuenteTitulo As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18, BaseColor.Gray)
            Dim fuenteSubtitulo As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.Black)
            Dim fuenteNormal As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black)
            Dim fuenteNegrita As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.Black)
            Dim fuenteTotalPagado As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Blue)
            Dim fuenteSaldoPendiente As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Red)

            ' Header con información de la empresa
            AgregarHeader(documento, fuenteTitulo, fuenteNormal)

            ' Información del recibo
            AgregarInformacionRecibo(documento, pago, fuenteSubtitulo, fuenteNormal)

            ' Información del propietario
            AgregarInformacionPropietario(documento, apartamento, fuenteSubtitulo, fuenteNormal)

            ' ✅ CORREGIDO: Detalles del pago con cálculos correctos
            AgregarDetallesPagoCORREGIDO(documento, pago, saldoAnterior, valorPagado, saldoActual, fuenteSubtitulo, fuenteNormal, fuenteTotalPagado)

            ' ✅ CORREGIDO: Saldo actual con cálculo correcto
            AgregarSaldoActualCORREGIDO(documento, saldoActual, fuenteSubtitulo, fuenteNormal, fuenteSaldoPendiente)

            ' Observaciones
            AgregarObservaciones(documento, pago, fuenteSubtitulo, fuenteNormal)

            ' Footer
            AgregarFooter(documento, fuenteNormal)

            documento.Close()

            Return rutaPersonalizada

        Catch ex As Exception
            Throw New Exception($"Error al generar PDF específico: {ex.Message}")
        End Try
    End Function

    Public Shared Function GenerarReciboDePagoSeguro(pago As PagoModel, apartamento As Apartamento) As String
        Try
            ' Crear directorio si no existe
            Dim rutaCarpeta As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM", "Recibos")
            If Not Directory.Exists(rutaCarpeta) Then
                Directory.CreateDirectory(rutaCarpeta)
            End If

            ' Nombre del archivo
            Dim nombreArchivo As String = $"Recibo_{pago.NumeroRecibo}_{apartamento.Torre}-{apartamento.NumeroApartamento}_{DateTime.Now:yyyyMMddHHmmss}.pdf"
            Dim rutaCompleta As String = Path.Combine(rutaCarpeta, nombreArchivo)

            ' ✅ CORREGIDO: Cálculos matemáticos correctos
            Dim saldoAnterior As Decimal = Math.Abs(pago.SaldoAnterior) ' Valor absoluto para manejar negativos
            Dim valorPagado As Decimal = pago.PagoAdministracion + pago.PagoIntereses ' Suma real de lo pagado
            Dim saldoActual As Decimal = saldoAnterior - valorPagado ' Restar lo pagado del saldo anterior

            ' Si el saldo queda negativo, significa que está a favor
            If saldoActual < 0 Then
                saldoActual = 0 ' O manejar como saldo a favor
            End If

            ' Crear documento PDF
            Dim documento As New Document(PageSize.Letter, 50, 50, 50, 50)
            Dim writer As PdfWriter = PdfWriter.GetInstance(documento, New FileStream(rutaCompleta, FileMode.Create))

            documento.Open()

            ' ✅ CORREGIDO: Fuentes especificando explícitamente iTextSharp.text.Font
            Dim fuenteTitulo As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18, BaseColor.Gray)
            Dim fuenteSubtitulo As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.Black)
            Dim fuenteNormal As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black)
            Dim fuenteNegrita As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.Black)
            Dim fuenteTotalPagado As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Blue)
            Dim fuenteSaldoPendiente As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Red)

            ' Header con información de la empresa
            AgregarHeader(documento, fuenteTitulo, fuenteNormal)

            ' Información del recibo
            AgregarInformacionRecibo(documento, pago, fuenteSubtitulo, fuenteNormal)

            ' Información del propietario
            AgregarInformacionPropietario(documento, apartamento, fuenteSubtitulo, fuenteNormal)

            ' ✅ CORREGIDO: Detalles del pago con cálculos correctos
            AgregarDetallesPagoCORREGIDO(documento, pago, saldoAnterior, valorPagado, saldoActual, fuenteSubtitulo, fuenteNormal, fuenteTotalPagado)

            ' ✅ CORREGIDO: Saldo actual con cálculo correcto
            AgregarSaldoActualCORREGIDO(documento, saldoActual, fuenteSubtitulo, fuenteNormal, fuenteSaldoPendiente)

            ' Observaciones
            AgregarObservaciones(documento, pago, fuenteSubtitulo, fuenteNormal)

            ' Footer
            AgregarFooter(documento, fuenteNormal)

            documento.Close()

            Return rutaCompleta

        Catch ex As Exception
            Throw New Exception($"Error al generar PDF: {ex.Message}")
        End Try
    End Function

    ' ✅ CORREGIDO: Método para detalles del pago con tipos explícitos
    Private Shared Sub AgregarDetallesPagoCORREGIDO(documento As Document, pago As PagoModel, saldoAnterior As Decimal, valorPagado As Decimal, saldoActual As Decimal, fuenteSubtitulo As iTextSharp.text.Font, fuenteNormal As iTextSharp.text.Font, fuenteTotalPagado As iTextSharp.text.Font)



        Dim tituloDetalles As New Paragraph("DETALLES DEL PAGO", fuenteSubtitulo)
        tituloDetalles.Alignment = Element.ALIGN_LEFT
        documento.Add(tituloDetalles)
        documento.Add(New Paragraph(" "))

        ' Crear tabla para detalles
        Dim tablaDetalles As New PdfPTable(2)
        tablaDetalles.WidthPercentage = 100
        tablaDetalles.SetWidths({70.0F, 30.0F})

        ' Header de la tabla
        Dim celdaConcepto As New PdfPCell(New Phrase("Concepto", fuenteSubtitulo))
        celdaConcepto.BackgroundColor = New BaseColor(52, 152, 219)
        celdaConcepto.HorizontalAlignment = Element.ALIGN_CENTER
        celdaConcepto.Padding = 8
        celdaConcepto.Border = iTextSharp.text.Rectangle.NO_BORDER

        Dim celdaValor As New PdfPCell(New Phrase("Valor", fuenteSubtitulo))
        celdaValor.BackgroundColor = New BaseColor(52, 152, 219)
        celdaValor.HorizontalAlignment = Element.ALIGN_CENTER
        celdaValor.Padding = 8
        celdaValor.Border = iTextSharp.text.Rectangle.NO_BORDER

        tablaDetalles.AddCell(celdaConcepto)
        tablaDetalles.AddCell(celdaValor)

        ' ✅ CORREGIDO: Mostrar saldo anterior como deuda previa
        AgregarFilaTabla(tablaDetalles, "Saldo Anterior", If(saldoAnterior > 0, "-$ " & saldoAnterior.ToString("N2"), "$ 0,00"), fuenteNormal)

        ' ✅ CORREGIDO: Mostrar valor realmente pagado
        If pago.PagoAdministracion > 0 Then
            AgregarFilaTabla(tablaDetalles, "Valor Pagado Administración", "$ " & pago.PagoAdministracion.ToString("N2"), fuenteNormal)
        End If

        If pago.PagoIntereses > 0 Then
            AgregarFilaTabla(tablaDetalles, "Valor Pagado Intereses", "$ " & pago.PagoIntereses.ToString("N2"), fuenteNormal)
        End If

        ' Fila vacía para separar
        AgregarFilaTabla(tablaDetalles, "", "", fuenteNormal)

        ' ✅ CORREGIDO: Total pagado = suma de administración + intereses
        Dim celdaTotalConcepto As New PdfPCell(New Phrase("TOTAL PAGADO", fuenteTotalPagado))
        celdaTotalConcepto.BackgroundColor = New BaseColor(211, 211, 211) ' ✅ CORREGIDO: Equivalente a LIGHT_GRAY
        celdaTotalConcepto.HorizontalAlignment = Element.ALIGN_LEFT
        celdaTotalConcepto.Padding = 8
        celdaTotalConcepto.Border = iTextSharp.text.Rectangle.BOX

        Dim celdaTotalValor As New PdfPCell(New Phrase("$ " & valorPagado.ToString("N2"), fuenteTotalPagado))
        celdaTotalValor.BackgroundColor = New BaseColor(211, 211, 211) ' ✅ CORREGIDO: Equivalente a LIGHT_GRAY
        celdaTotalValor.HorizontalAlignment = Element.ALIGN_RIGHT
        celdaTotalValor.Padding = 8
        celdaTotalValor.Border = iTextSharp.text.Rectangle.BOX

        tablaDetalles.AddCell(celdaTotalConcepto)
        tablaDetalles.AddCell(celdaTotalValor)

        documento.Add(tablaDetalles)
    End Sub

    ' ✅ CORREGIDO: Método para saldo actual con tipos explícitos
    Private Shared Sub AgregarSaldoActualCORREGIDO(documento As Document, saldoActual As Decimal, fuenteSubtitulo As iTextSharp.text.Font, fuenteNormal As iTextSharp.text.Font, fuenteSaldoPendiente As iTextSharp.text.Font)

        Dim tituloSaldo As New Paragraph("SALDO ACTUAL", fuenteSubtitulo)
        tituloSaldo.Alignment = Element.ALIGN_LEFT
        documento.Add(tituloSaldo)
        documento.Add(New Paragraph(" "))

        Dim tablaSaldo As New PdfPTable(2)
        tablaSaldo.WidthPercentage = 100
        tablaSaldo.SetWidths({70.0F, 30.0F})

        ' Header
        Dim celdaConceptoSaldo As New PdfPCell(New Phrase("Concepto", fuenteSubtitulo))
        celdaConceptoSaldo.BackgroundColor = New BaseColor(52, 152, 219)
        celdaConceptoSaldo.HorizontalAlignment = Element.ALIGN_CENTER
        celdaConceptoSaldo.Padding = 8

        Dim celdaValorSaldo As New PdfPCell(New Phrase("Valor", fuenteSubtitulo))
        celdaValorSaldo.BackgroundColor = New BaseColor(52, 152, 219)
        celdaValorSaldo.HorizontalAlignment = Element.ALIGN_CENTER
        celdaValorSaldo.Padding = 14

        tablaSaldo.AddCell(celdaConceptoSaldo)
        tablaSaldo.AddCell(celdaValorSaldo)

        ' ✅ CORREGIDO: Saldo cuotas pendientes = saldo anterior - valor pagado
        AgregarFilaTabla(tablaSaldo, "Saldo Cuotas Pendientes", "$ " & saldoActual.ToString("N2"), fuenteNormal)

        documento.Add(tablaSaldo)

        documento.Add(New Paragraph(" "))

        ' ✅ CORREGIDO: Total saldo pendiente
        Dim tablaTotalSaldo As New PdfPTable(2)
        tablaTotalSaldo.WidthPercentage = 100
        tablaTotalSaldo.SetWidths({70.0F, 30.0F})

        Dim celdaTotalSaldoConcepto As New PdfPCell(New Phrase("TOTAL SALDO PENDIENTE", fuenteSaldoPendiente))
        celdaTotalSaldoConcepto.BackgroundColor = New BaseColor(211, 211, 211) ' ✅ CORREGIDO: Equivalente a LIGHT_GRAY
        celdaTotalSaldoConcepto.HorizontalAlignment = Element.ALIGN_LEFT
        celdaTotalSaldoConcepto.Padding = 8
        celdaTotalSaldoConcepto.Border = iTextSharp.text.Rectangle.BOX

        Dim celdaTotalSaldoValor As New PdfPCell(New Phrase("$ " & saldoActual.ToString("N2"), fuenteSaldoPendiente))
        celdaTotalSaldoValor.BackgroundColor = New BaseColor(211, 211, 211) ' ✅ CORREGIDO: Equivalente a LIGHT_GRAY
        celdaTotalSaldoValor.HorizontalAlignment = Element.ALIGN_RIGHT
        celdaTotalSaldoValor.Padding = 8
        celdaTotalSaldoValor.Border = iTextSharp.text.Rectangle.BOX

        tablaTotalSaldo.AddCell(celdaTotalSaldoConcepto)
        tablaTotalSaldo.AddCell(celdaTotalSaldoValor)

        documento.Add(tablaTotalSaldo)
    End Sub

    ' ✅ CORREGIDO: Métodos auxiliares con tipos explícitos
    Private Shared Sub AgregarHeader(documento As Document, fuenteTitulo As iTextSharp.text.Font, fuenteNormal As iTextSharp.text.Font)
        Dim titulo As New Paragraph("CONJUNTO RESIDENCIAL COOPDIASAM", fuenteTitulo)
        titulo.Alignment = Element.ALIGN_CENTER
        documento.Add(titulo)

        Dim subtitulo As New Paragraph("RECIBO DE PAGO ADMINISTRACIÓN", fuenteTitulo)
        subtitulo.Alignment = Element.ALIGN_CENTER
        documento.Add(subtitulo)

        Dim infoEmpresa As New Paragraph("NIT: 830.123.456-7" & vbLf & "Dirección: Calle 123 #45-67, Bogotá" & vbLf & "Teléfono: (601) 234-5678", fuenteNormal)
        infoEmpresa.Alignment = Element.ALIGN_CENTER
        documento.Add(infoEmpresa)
        documento.Add(New Paragraph(" ", fuenteNormal))
    End Sub

    Private Shared Sub AgregarInformacionRecibo(documento As Document, pago As PagoModel, fuenteSubtitulo As iTextSharp.text.Font, fuenteNormal As iTextSharp.text.Font)
        Dim tablaRecibo As New PdfPTable(4)
        tablaRecibo.WidthPercentage = 100
        tablaRecibo.SetWidths({25.0F, 25.0F, 25.0F, 25.0F})

        AgregarCeldaInfo(tablaRecibo, "Recibo No:", pago.NumeroRecibo, fuenteNormal)
        AgregarCeldaInfo(tablaRecibo, "Fecha Pago:", pago.FechaPago.ToString("dd/MM/yyyy"), fuenteNormal)
        AgregarCeldaInfo(tablaRecibo, "Fecha Registro:", pago.FechaRegistro.ToString("dd/MM/yyyy"), fuenteNormal)
        AgregarCeldaInfo(tablaRecibo, "Usuario:", If(String.IsNullOrEmpty(pago.UsuarioRegistro), "Fernando_Gamba", pago.UsuarioRegistro), fuenteNormal)

        documento.Add(tablaRecibo)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarInformacionPropietario(documento As Document, apartamento As Apartamento, fuenteSubtitulo As iTextSharp.text.Font, fuenteNormal As iTextSharp.text.Font)
        Dim tituloInfo As New Paragraph("INFORMACIÓN DEL PROPIETARIO", fuenteSubtitulo)

        tituloInfo.Alignment = Element.ALIGN_LEFT
        documento.Add(tituloInfo)

        Dim tablaInfo As New PdfPTable(2)
        tablaInfo.WidthPercentage = 100
        tablaInfo.SetWidths({30.0F, 70.0F})
        documento.Add(New Paragraph(" "))


        AgregarFilaInfo(tablaInfo, "Propietario:", If(String.IsNullOrEmpty(apartamento.NombreResidente), "No registrado", apartamento.NombreResidente), fuenteNormal)
        AgregarFilaInfo(tablaInfo, "Email:", If(String.IsNullOrEmpty(apartamento.Correo), "No registrado", apartamento.Correo), fuenteNormal)
        AgregarFilaInfo(tablaInfo, "Apartamento:", $"T{apartamento.Torre}-{apartamento.NumeroApartamento}", fuenteNormal)
        AgregarFilaInfo(tablaInfo, "Matrícula Inmobiliaria:", If(String.IsNullOrEmpty(apartamento.MatriculaInmobiliaria), "No registrada", apartamento.MatriculaInmobiliaria), fuenteNormal)
        AgregarFilaInfo(tablaInfo, "Teléfono:", If(String.IsNullOrEmpty(apartamento.Telefono), "No registrado", apartamento.Telefono), fuenteNormal)

        documento.Add(tablaInfo)
    End Sub

    Private Shared Sub AgregarObservaciones(documento As Document, pago As PagoModel, fuenteSubtitulo As iTextSharp.text.Font, fuenteNormal As iTextSharp.text.Font)
        documento.Add(New Paragraph(" "))

        Dim tituloObs As New Paragraph("OBSERVACIONES", fuenteSubtitulo)
        tituloObs.Alignment = Element.ALIGN_LEFT
        documento.Add(tituloObs)

        Dim observaciones As String = If(String.IsNullOrEmpty(pago.Observaciones), "adm. Feb. 2021", pago.Observaciones)
        Dim parrafoObs As New Paragraph(observaciones, fuenteNormal)
        documento.Add(parrafoObs)
    End Sub

    Private Shared Sub AgregarFooter(documento As Document, fuenteNormal As iTextSharp.text.Font)
        documento.Add(New Paragraph(" "))
        documento.Add(New Paragraph("_________________________________________________________________________________________________________________", fuenteNormal))
        documento.Add(New Paragraph(" "))

        Dim footer As New Paragraph("Este recibo fue generado automáticamente por el sistema ClickApT" & vbLf &
                                   "Fecha de generación: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") & vbLf &
                                   "Para consultas comuníquese al (601) 234-5678", fuenteNormal)
        footer.Alignment = Element.ALIGN_CENTER
        documento.Add(footer)
    End Sub

    ' ✅ CORREGIDO: Métodos auxiliares para crear celdas con tipos explícitos
    Private Shared Sub AgregarCeldaInfo(tabla As PdfPTable, etiqueta As String, valor As String, fuente As iTextSharp.text.Font)
        Dim celdaEtiqueta As New PdfPCell(New Phrase(etiqueta, fuente))
        celdaEtiqueta.Border = iTextSharp.text.Rectangle.NO_BORDER
        celdaEtiqueta.Padding = 5
        tabla.AddCell(celdaEtiqueta)

        Dim celdaValor As New PdfPCell(New Phrase(valor, fuente))
        celdaValor.Border = iTextSharp.text.Rectangle.NO_BORDER
        celdaValor.Padding = 5
        tabla.AddCell(celdaValor)
    End Sub

    Private Shared Sub AgregarFilaInfo(tabla As PdfPTable, etiqueta As String, valor As String, fuente As iTextSharp.text.Font)
        Dim celdaEtiqueta As New PdfPCell(New Phrase(etiqueta, fuente))
        celdaEtiqueta.BackgroundColor = New BaseColor(211, 211, 211) ' ✅ CORREGIDO: Equivalente a LIGHT_GRAY
        celdaEtiqueta.Padding = 5
        tabla.AddCell(celdaEtiqueta)

        Dim celdaValor As New PdfPCell(New Phrase(valor, fuente))
        celdaValor.Padding = 5
        tabla.AddCell(celdaValor)
    End Sub

    Private Shared Sub AgregarFilaTabla(tabla As PdfPTable, concepto As String, valor As String, fuente As iTextSharp.text.Font)
        Dim celdaConcepto As New PdfPCell(New Phrase(concepto, fuente))
        celdaConcepto.Padding = 5
        celdaConcepto.HorizontalAlignment = Element.ALIGN_LEFT
        tabla.AddCell(celdaConcepto)

        Dim celdaValor As New PdfPCell(New Phrase(valor, fuente))
        celdaValor.Padding = 5
        celdaValor.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla.AddCell(celdaValor)
    End Sub

End Class