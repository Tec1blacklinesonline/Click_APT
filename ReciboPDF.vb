Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class ReciboPDF

    ' Constantes del conjunto según la imagen del recibo
    Private Const NOMBRE_CONJUNTO As String = "COOPDIASAM"
    Private Const DESCRIPCION_CONJUNTO As String = "Conjunto Habitacional"
    Private Const NIT As String = "900 225 635 - 8"
    Private Const DIRECCION As String = "Barrio Villa Café - Ibagué Tolima"
    Private Const EMAIL_CONJUNTO As String = "conjuntocoopdiasam@yahoo.es"
    Private Const EMAIL_COPIA As String = "apolo20136@gmail.com"
    Private Const ADMINISTRADOR_NOMBRE As String = "Fernando Gamba"
    Private Const ADMINISTRADOR_TELEFONO As String = "+57 321-9597100"
    Private Const ADMINISTRADOR_CARGO As String = "ADMINISTRADOR CONJUNTO RESIDENCIAL COOPDIASAM"

    Public Shared Function GenerarReciboPDF(pago As PagoModel, apartamento As Apartamento, rutaArchivo As String) As Boolean
        Try
            ' Crear el documento PDF con márgenes ajustados
            Dim documento As New Document(PageSize.A4, 25, 25, 30, 30)
            Dim writer As PdfWriter = PdfWriter.GetInstance(documento, New FileStream(rutaArchivo, FileMode.Create))

            documento.Open()

            ' Agregar contenido siguiendo el formato exacto del recibo
            AgregarTituloRecibo(documento)
            AgregarEncabezadoConDatos(documento, apartamento, pago)
            AgregarNumeroRecibo(documento, pago)
            AgregarEstadoCuenta(documento, pago)
            AgregarDetallePagos(documento, pago)
            AgregarSaldos(documento, pago)
            AgregarObservaciones(documento, pago)
            AgregarFirmaAdministrador(documento)

            documento.Close()
            Return True

        Catch ex As Exception
            Throw New Exception($"Error al generar PDF: {ex.Message}")
        End Try
    End Function

    Private Shared Sub AgregarTituloRecibo(documento As Document)
        ' Título principal centrado con borde
        Dim fontTitulo As New Font(Font.FontFamily.HELVETICA, 14, Font.BOLD)

        Dim tablaTitulo As New PdfPTable(1)
        tablaTitulo.WidthPercentage = 100

        Dim celdaTitulo As New PdfPCell(New Phrase("RECIBO MENSUAL DE PAGO", fontTitulo))
        celdaTitulo.HorizontalAlignment = Element.ALIGN_CENTER
        celdaTitulo.VerticalAlignment = Element.ALIGN_MIDDLE
        celdaTitulo.Border = Rectangle.BOX
        celdaTitulo.BorderWidth = 2
        celdaTitulo.Padding = 8
        celdaTitulo.BackgroundColor = BaseColor.LIGHT_GRAY

        tablaTitulo.AddCell(celdaTitulo)
        documento.Add(tablaTitulo)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarEncabezadoConDatos(documento As Document, apartamento As Apartamento, pago As PagoModel)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD)
        Dim fontGrande As New Font(Font.FontFamily.HELVETICA, 12, Font.BOLD)
        Dim fontMuyGrande As New Font(Font.FontFamily.HELVETICA, 24, Font.BOLD)

        ' Crear tabla principal con 2 columnas
        Dim tablaEncabezado As New PdfPTable(2)
        tablaEncabezado.WidthPercentage = 100
        tablaEncabezado.SetWidths({50, 50})

        ' LADO IZQUIERDO - Logo y datos del conjunto
        Dim celdaIzquierda As New PdfPCell()
        celdaIzquierda.Border = Rectangle.BOX
        celdaIzquierda.Padding = 10

        ' Logo (simulado con texto)
        Dim parrafoLogo As New Paragraph()
        parrafoLogo.Add(New Chunk("🏢 ", New Font(Font.FontFamily.HELVETICA, 16, Font.NORMAL)))
        parrafoLogo.Add(New Chunk(NOMBRE_CONJUNTO, fontGrande))
        parrafoLogo.Add(Chunk.NEWLINE)
        parrafoLogo.Add(New Chunk(DESCRIPCION_CONJUNTO, fontNormal))
        parrafoLogo.Add(Chunk.NEWLINE)
        parrafoLogo.Add(New Chunk($"NIT. {NIT}", fontNormal))
        parrafoLogo.Add(Chunk.NEWLINE)
        parrafoLogo.Add(New Chunk(DIRECCION, fontNormal))

        celdaIzquierda.AddElement(parrafoLogo)

        ' Fechas de impresión y copia
        Dim parrafoFechas As New Paragraph()
        parrafoFechas.Add(Chunk.NEWLINE)
        parrafoFechas.Add(New Chunk($"FECHA IMPRESIÓN: {DateTime.Now:dd-MMM-yyyy HH:mm}", New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL)))
        parrafoFechas.Add(Chunk.NEWLINE)
        parrafoFechas.Add(New Chunk($"COPIA AL CORREO: {apartamento.Correo}", New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL)))

        celdaIzquierda.AddElement(parrafoFechas)

        ' LADO DERECHO - Datos del propietario
        Dim celdaDerecha As New PdfPCell()
        celdaDerecha.Border = Rectangle.BOX
        celdaDerecha.Padding = 10

        ' Nombre del propietario
        Dim parrafoNombre As New Paragraph()
        parrafoNombre.Add(New Chunk(apartamento.NombreResidente.ToUpper(), fontBold))
        parrafoNombre.Alignment = Element.ALIGN_RIGHT
        celdaDerecha.AddElement(parrafoNombre)

        ' Número de apartamento grande
        Dim parrafoApto As New Paragraph()
        parrafoApto.Add(New Chunk($"APT - {apartamento.Torre}{apartamento.NumeroApartamento}", fontMuyGrande))
        parrafoApto.Alignment = Element.ALIGN_RIGHT
        celdaDerecha.AddElement(parrafoApto)

        ' Datos adicionales
        Dim parrafoDatos As New Paragraph()
        parrafoDatos.Add(Chunk.NEWLINE)
        parrafoDatos.Add(New Chunk($"FICHA CATASTRAL: {pago.MatriculaInmobiliaria}", fontNormal))
        parrafoDatos.Add(Chunk.NEWLINE)
        parrafoDatos.Add(New Chunk($"CEL: {apartamento.Telefono}", fontNormal))
        parrafoDatos.Add(Chunk.NEWLINE)
        parrafoDatos.Add(New Chunk($"EMAIL: {apartamento.Correo}", fontNormal))
        parrafoDatos.Alignment = Element.ALIGN_RIGHT
        celdaDerecha.AddElement(parrafoDatos)

        tablaEncabezado.AddCell(celdaIzquierda)
        tablaEncabezado.AddCell(celdaDerecha)

        documento.Add(tablaEncabezado)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarNumeroRecibo(documento As Document, pago As PagoModel)
        Dim fontGrande As New Font(Font.FontFamily.HELVETICA, 18, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL)

        ' RECIBO DE CAJA
        Dim parrafoRecibo As New Paragraph()
        parrafoRecibo.Add(New Chunk("RECIBO DE CAJA", fontNormal))
        parrafoRecibo.Add(Chunk.NEWLINE)
        parrafoRecibo.Add(New Chunk($"No. {pago.NumeroRecibo}", fontGrande))
        parrafoRecibo.Alignment = Element.ALIGN_LEFT

        documento.Add(parrafoRecibo)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarEstadoCuenta(documento As Document, pago As PagoModel)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        Dim parrafoEstado As New Paragraph()
        parrafoEstado.Add(New Chunk($"ESTADO FINAL DE CUENTA  : $ {pago.SaldoActual:N0}", fontNormal))
        parrafoEstado.Add(Chunk.NEWLINE)
        parrafoEstado.Add(New Chunk($"FECHA ULTIMO PAGO       : {pago.FechaPago:dd-MMM-yyyy}", fontNormal))

        documento.Add(parrafoEstado)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarDetallePagos(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 12, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Crear tabla con 2 columnas para detalle y saldos
        Dim tablaCompleta As New PdfPTable(2)
        tablaCompleta.WidthPercentage = 100
        tablaCompleta.SetWidths({50, 50})

        ' LADO IZQUIERDO - DETALLE DE PAGOS
        Dim celdaDetalle As New PdfPCell()
        celdaDetalle.Border = Rectangle.NO_BORDER
        celdaDetalle.PaddingRight = 10

        Dim parrafoDetalle As New Paragraph()
        parrafoDetalle.Add(New Chunk("DETALLE DE PAGOS", fontBold))
        parrafoDetalle.Add(Chunk.NEWLINE)
        parrafoDetalle.Add(Chunk.NEWLINE)
        parrafoDetalle.Add(New Chunk($"SALDO ANTERIOR           : $ {pago.SaldoAnterior:N0}", fontNormal))
        parrafoDetalle.Add(Chunk.NEWLINE)
        parrafoDetalle.Add(New Chunk($"V/R PAGADO ADMINISTRACION : $ {pago.PagoAdministracion:N0}", fontNormal))
        parrafoDetalle.Add(Chunk.NEWLINE)
        parrafoDetalle.Add(New Chunk($"V/R PAGADO INTERESES     : $ {pago.PagoIntereses:N0}", fontNormal))
        parrafoDetalle.Add(Chunk.NEWLINE)
        parrafoDetalle.Add(New Chunk($"TOTAL                    : $ {pago.TotalPagado:N0}", fontNormal))

        celdaDetalle.AddElement(parrafoDetalle)

        ' LADO DERECHO - SALDOS
        Dim celdaSaldos As New PdfPCell()
        celdaSaldos.Border = Rectangle.NO_BORDER
        celdaSaldos.PaddingLeft = 10

        Dim parrafoSaldos As New Paragraph()
        parrafoSaldos.Add(New Chunk("SALDOS", fontBold))
        parrafoSaldos.Add(Chunk.NEWLINE)
        parrafoSaldos.Add(Chunk.NEWLINE)
        parrafoSaldos.Add(New Chunk($"INTERÉS     : $ {pago.PagoIntereses:N0}", fontNormal))
        parrafoSaldos.Add(Chunk.NEWLINE)
        parrafoSaldos.Add(New Chunk($"CUOTA       : $ {pago.CuotaActual:N0}", fontNormal))
        parrafoSaldos.Add(Chunk.NEWLINE)
        parrafoSaldos.Add(New Chunk($"TOTAL       : $ {pago.TotalPagado:N0}", fontNormal))

        celdaSaldos.AddElement(parrafoSaldos)

        tablaCompleta.AddCell(celdaDetalle)
        tablaCompleta.AddCell(celdaSaldos)

        documento.Add(tablaCompleta)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarObservaciones(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 11, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Título
        Dim titulo As New Paragraph("OBSERVACION:", fontBold)
        documento.Add(titulo)

        ' Campo de observaciones con borde
        Dim tablaObs As New PdfPTable(1)
        tablaObs.WidthPercentage = 100

        Dim celdaObs As New PdfPCell(New Phrase(If(String.IsNullOrEmpty(pago.Observaciones), " ", pago.Observaciones), fontNormal))
        celdaObs.Border = Rectangle.BOX
        celdaObs.Padding = 8
        celdaObs.MinimumHeight = 40
        tablaObs.AddCell(celdaObs)

        documento.Add(tablaObs)
        documento.Add(New Paragraph(" "))
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarFirmaAdministrador(documento As Document)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 9, Font.BOLD)

        ' Texto "Respetuosamente"
        Dim parrafoRespetuosamente As New Paragraph("Respetuosamente,", fontNormal)
        parrafoRespetuosamente.Alignment = Element.ALIGN_LEFT
        documento.Add(parrafoRespetuosamente)

        documento.Add(New Paragraph(" "))
        documento.Add(New Paragraph(" "))

        ' Imagen de firma (simulada con texto)
        Dim parrafoFirma As New Paragraph("Fernando Gamba", New Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLUE))
        parrafoFirma.Alignment = Element.ALIGN_LEFT
        documento.Add(parrafoFirma)

        ' Línea para firma
        Dim parrafoLinea As New Paragraph("____________________", fontNormal)
        parrafoLinea.Alignment = Element.ALIGN_LEFT
        documento.Add(parrafoLinea)

        ' Información del administrador
        Dim parrafoAdmin As New Paragraph()
        parrafoAdmin.Add(New Chunk(ADMINISTRADOR_NOMBRE, fontNormal))
        parrafoAdmin.Add(Chunk.NEWLINE)
        parrafoAdmin.Add(New Chunk(ADMINISTRADOR_TELEFONO, fontNormal))
        parrafoAdmin.Add(Chunk.NEWLINE)
        parrafoAdmin.Add(New Chunk(ADMINISTRADOR_CARGO, fontNormal))
        parrafoAdmin.Alignment = Element.ALIGN_LEFT

        documento.Add(parrafoAdmin)
    End Sub

    ' Método para generar PDF directamente desde FormPagos
    Public Shared Function GenerarReciboDesdeFormulario(pago As PagoModel, apartamento As Apartamento) As String
        Try
            ' Crear nombre de archivo único
            Dim nombreArchivo As String = $"Recibo_{pago.NumeroRecibo}_{apartamento.Torre}{apartamento.NumeroApartamento}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
            Dim rutaArchivo As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), nombreArchivo)

            ' Generar el PDF
            If GenerarReciboPDF(pago, apartamento, rutaArchivo) Then
                Return rutaArchivo
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Throw New Exception($"Error al generar recibo: {ex.Message}")
        End Try
    End Function

End Class