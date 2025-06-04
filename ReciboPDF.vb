Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class ReciboPDF

    ' Constantes del conjunto según especificaciones
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
            ' Crear el documento PDF
            Dim documento As New Document(PageSize.A4, 30, 30, 40, 40)
            Dim writer As PdfWriter = PdfWriter.GetInstance(documento, New FileStream(rutaArchivo, FileMode.Create))

            documento.Open()

            ' Agregar contenido al PDF siguiendo el formato exacto del recibo
            AgregarEncabezadoPrincipal(documento)
            AgregarDatosIdentificacion(documento, pago)
            AgregarDatosResidente(documento, apartamento, pago)
            AgregarDatosRecibo(documento, pago)
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

    Private Shared Sub AgregarEncabezadoPrincipal(documento As Document)
        ' Crear tabla para el encabezado principal
        Dim tablaEncabezado As New PdfPTable(2)
        tablaEncabezado.WidthPercentage = 100
        tablaEncabezado.SetWidths({70, 30})

        ' Lado izquierdo - Información del conjunto
        Dim fontTitulo As New Font(Font.FontFamily.HELVETICA, 14, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)
        Dim fontSmall As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL)

        Dim celdaIzquierda As New PdfPCell()
        celdaIzquierda.Border = Rectangle.BOX
        celdaIzquierda.Padding = 8

        ' Logo y nombre del conjunto (como texto ya que no tienes logo)
        Dim parrafoLogo As New Paragraph()
        parrafoLogo.Add(New Chunk("🏢 ", New Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL)))
        parrafoLogo.Add(New Chunk(NOMBRE_CONJUNTO, fontTitulo))
        parrafoLogo.Add(Chunk.NEWLINE)
        parrafoLogo.Add(New Chunk(DESCRIPCION_CONJUNTO, fontNormal))
        parrafoLogo.Add(Chunk.NEWLINE)
        parrafoLogo.Add(New Chunk($"NIT. {NIT}", fontSmall))

        celdaIzquierda.AddElement(parrafoLogo)

        ' Lado derecho - Información del propietario (se llenará después)
        Dim celdaDerecha As New PdfPCell(New Phrase(""))
        celdaDerecha.Border = Rectangle.BOX
        celdaDerecha.Padding = 8

        tablaEncabezado.AddCell(celdaIzquierda)
        tablaEncabezado.AddCell(celdaDerecha)

        documento.Add(tablaEncabezado)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarDatosIdentificacion(documento As Document, pago As PagoModel)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL)

        ' Fecha de impresión
        Dim fechaImpresion As New Paragraph($"FECHA IMPRESIÓN: {DateTime.Now:dd-MMM-yyyy HH:mm}", fontNormal)
        fechaImpresion.Alignment = Element.ALIGN_LEFT
        documento.Add(fechaImpresion)

        ' Email para copia
        Dim emailCopia As New Paragraph($"COPIA AL CORREO: {EMAIL_COPIA}", fontNormal)
        emailCopia.Alignment = Element.ALIGN_LEFT
        documento.Add(emailCopia)

        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarDatosResidente(documento As Document, apartamento As Apartamento, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 10, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Título de sección
        Dim tituloResidente As New Paragraph("DATOS DEL RESIDENTE", fontBold)
        tituloResidente.Alignment = Element.ALIGN_CENTER
        documento.Add(tituloResidente)

        documento.Add(New Paragraph(" "))

        ' Crear tabla para datos del residente
        Dim tabla As New PdfPTable(2)
        tabla.WidthPercentage = 100
        tabla.SetWidths({30, 70})

        ' Nombre
        AgregarCeldaSimple(tabla, "Nombre:", fontBold)
        AgregarCeldaSimple(tabla, apartamento.NombreResidente.ToUpper(), fontBold)

        ' Apartamento
        AgregarCeldaSimple(tabla, "Apartamento:", fontBold)
        AgregarCeldaSimple(tabla, $"APT - {apartamento.NumeroApartamento}", fontBold)

        ' Ficha catastral
        AgregarCeldaSimple(tabla, "Ficha catastral:", fontNormal)
        AgregarCeldaSimple(tabla, pago.MatriculaInmobiliaria, fontNormal)

        ' Celular
        AgregarCeldaSimple(tabla, "Celular:", fontNormal)
        AgregarCeldaSimple(tabla, apartamento.Telefono, fontNormal)

        ' Correo electrónico
        AgregarCeldaSimple(tabla, "Correo electrónico:", fontNormal)
        AgregarCeldaSimple(tabla, apartamento.Correo, fontNormal)

        documento.Add(tabla)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarDatosRecibo(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 12, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Título de sección
        Dim tituloRecibo As New Paragraph("DATOS DEL RECIBO", fontBold)
        tituloRecibo.Alignment = Element.ALIGN_CENTER
        documento.Add(tituloRecibo)

        documento.Add(New Paragraph(" "))

        ' RECIBO DE CAJA
        Dim reciboTitulo As New Paragraph("RECIBO DE CAJA", fontBold)
        reciboTitulo.Alignment = Element.ALIGN_LEFT
        documento.Add(reciboTitulo)

        Dim numeroRecibo As New Paragraph($"No. {pago.NumeroRecibo}", fontBold)
        numeroRecibo.Alignment = Element.ALIGN_LEFT
        documento.Add(numeroRecibo)

        documento.Add(New Paragraph(" "))

        ' Estado final de cuenta
        Dim estadoFinal As New Paragraph($"Estado final de cuenta: ${pago.SaldoActual:N0}", fontNormal)
        documento.Add(estadoFinal)

        ' Fecha último pago
        Dim fechaUltimoPago As New Paragraph($"Fecha del último pago: {pago.FechaPago:dd-MMM-yyyy}", fontNormal)
        documento.Add(fechaUltimoPago)

        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarDetallePagos(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 11, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Título
        Dim titulo As New Paragraph("DETALLE DE PAGOS", fontBold)
        titulo.Alignment = Element.ALIGN_CENTER
        documento.Add(titulo)

        documento.Add(New Paragraph(" "))

        ' Crear tabla con encabezados
        Dim tabla As New PdfPTable(2)
        tabla.WidthPercentage = 100
        tabla.SetWidths({60, 40})

        ' Encabezados
        AgregarCeldaConBorde(tabla, "Concepto", fontBold, Element.ALIGN_CENTER, BaseColor.LIGHT_GRAY)
        AgregarCeldaConBorde(tabla, "Valor", fontBold, Element.ALIGN_CENTER, BaseColor.LIGHT_GRAY)

        ' Saldo anterior
        AgregarCeldaConBorde(tabla, "Saldo Anterior", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaConBorde(tabla, $"${pago.SaldoAnterior:N0}", fontBold, Element.ALIGN_RIGHT)

        ' V/R Pagado Administración
        AgregarCeldaConBorde(tabla, "V/R Pagado Administración", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaConBorde(tabla, $"${pago.PagoAdministracion:N0}", fontBold, Element.ALIGN_RIGHT)

        ' V/R Pagado Intereses
        AgregarCeldaConBorde(tabla, "V/R Pagado Intereses", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaConBorde(tabla, $"${pago.PagoIntereses:N0}", fontBold, Element.ALIGN_RIGHT)

        ' Total Pagado
        AgregarCeldaConBorde(tabla, "Total Pagado", fontBold, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
        AgregarCeldaConBorde(tabla, $"${pago.TotalPagado:N0}", fontBold, Element.ALIGN_RIGHT, BaseColor.LIGHT_GRAY)

        documento.Add(tabla)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarSaldos(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 11, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Título
        Dim titulo As New Paragraph("SALDOS", fontBold)
        titulo.Alignment = Element.ALIGN_CENTER
        documento.Add(titulo)

        documento.Add(New Paragraph(" "))

        ' Crear tabla
        Dim tabla As New PdfPTable(2)
        tabla.WidthPercentage = 100
        tabla.SetWidths({60, 40})

        ' Encabezados
        AgregarCeldaConBorde(tabla, "Concepto", fontBold, Element.ALIGN_CENTER, BaseColor.LIGHT_GRAY)
        AgregarCeldaConBorde(tabla, "Valor", fontBold, Element.ALIGN_CENTER, BaseColor.LIGHT_GRAY)

        ' Interés
        AgregarCeldaConBorde(tabla, "Interés", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaConBorde(tabla, $"${pago.PagoIntereses:N0}", fontBold, Element.ALIGN_RIGHT)

        ' Cuota
        AgregarCeldaConBorde(tabla, "Cuota", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaConBorde(tabla, $"${pago.CuotaActual:N0}", fontBold, Element.ALIGN_RIGHT)

        ' Total
        AgregarCeldaConBorde(tabla, "Total", fontBold, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
        AgregarCeldaConBorde(tabla, $"${pago.TotalPagado:N0}", fontBold, Element.ALIGN_RIGHT, BaseColor.LIGHT_GRAY)

        documento.Add(tabla)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarObservaciones(documento As Document, pago As PagoModel)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 11, Font.BOLD)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)

        ' Título
        Dim titulo As New Paragraph("OBSERVACIÓN", fontBold)
        titulo.Alignment = Element.ALIGN_CENTER
        documento.Add(titulo)

        documento.Add(New Paragraph(" "))

        ' Campo de observaciones (con borde)
        Dim tablaObs As New PdfPTable(1)
        tablaObs.WidthPercentage = 100

        Dim celdaObs As New PdfPCell(New Phrase(If(String.IsNullOrEmpty(pago.Observaciones), " ", pago.Observaciones), fontNormal))
        celdaObs.Border = Rectangle.BOX
        celdaObs.Padding = 10
        celdaObs.MinimumHeight = 30
        tablaObs.AddCell(celdaObs)

        documento.Add(tablaObs)
        documento.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarFirmaAdministrador(documento As Document)
        Dim fontNormal As New Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL)
        Dim fontBold As New Font(Font.FontFamily.HELVETICA, 9, Font.BOLD)

        documento.Add(New Paragraph(" "))

        ' Texto "Respetuosamente"
        Dim parrafoFirma As New Paragraph("Respetuosamente,", fontNormal)
        parrafoFirma.Alignment = Element.ALIGN_LEFT
        documento.Add(parrafoFirma)

        documento.Add(New Paragraph(" "))
        documento.Add(New Paragraph(" "))

        ' Crear tabla para firma
        Dim tablaFirma As New PdfPTable(1)
        tablaFirma.WidthPercentage = 50
        tablaFirma.HorizontalAlignment = Element.ALIGN_LEFT

        ' Línea para firma
        Dim celdaLinea As New PdfPCell(New Phrase("_________________________", fontNormal))
        celdaLinea.Border = Rectangle.NO_BORDER
        celdaLinea.HorizontalAlignment = Element.ALIGN_LEFT
        tablaFirma.AddCell(celdaLinea)

        ' Nombre del administrador
        Dim celdaNombre As New PdfPCell(New Phrase(ADMINISTRADOR_NOMBRE, fontBold))
        celdaNombre.Border = Rectangle.NO_BORDER
        celdaNombre.HorizontalAlignment = Element.ALIGN_LEFT
        tablaFirma.AddCell(celdaNombre)

        ' Teléfono
        Dim celdaTelefono As New PdfPCell(New Phrase(ADMINISTRADOR_TELEFONO, fontNormal))
        celdaTelefono.Border = Rectangle.NO_BORDER
        celdaTelefono.HorizontalAlignment = Element.ALIGN_LEFT
        tablaFirma.AddCell(celdaTelefono)

        ' Cargo
        Dim celdaCargo As New PdfPCell(New Phrase(ADMINISTRADOR_CARGO, fontNormal))
        celdaCargo.Border = Rectangle.NO_BORDER
        celdaCargo.HorizontalAlignment = Element.ALIGN_LEFT
        tablaFirma.AddCell(celdaCargo)

        documento.Add(tablaFirma)
    End Sub

    Private Shared Sub AgregarCeldaSimple(tabla As PdfPTable, texto As String, fuente As Font)
        Dim celda As New PdfPCell(New Phrase(texto, fuente))
        celda.Border = Rectangle.NO_BORDER
        celda.Padding = 3
        tabla.AddCell(celda)
    End Sub

    Private Shared Sub AgregarCeldaConBorde(tabla As PdfPTable, texto As String, fuente As Font, alineacion As Integer, Optional colorFondo As BaseColor = Nothing)
        Dim celda As New PdfPCell(New Phrase(texto, fuente))
        celda.HorizontalAlignment = alineacion
        celda.Border = Rectangle.BOX
        celda.Padding = 5
        celda.BorderWidth = 1

        If colorFondo IsNot Nothing Then
            celda.BackgroundColor = colorFondo
        End If

        tabla.AddCell(celda)
    End Sub

End Class