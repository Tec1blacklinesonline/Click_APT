' ============================================================================
' GENERADOR DE RECIBOS PDF PARA PAGOS EXTRA
' ✅ Extiende ReciboPDF para manejar multas, adiciones, sanciones, etc.
' ✅ Diseño diferenciado para pagos extra
' ============================================================================

Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Configuration
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.draw

Public Class ReciboPDFExtra

    ''' <summary>
    ''' Genera un recibo PDF específico para pagos extra (multas, adiciones, etc.)
    ''' </summary>
    Public Shared Function GenerarReciboPagoExtra(pagoExtra As PagoModel, apartamento As Apartamento) As String
        Dim rutaArchivo As String = String.Empty

        Try
            ' Obtener ruta de recibos
            Dim rutaRecibos As String = ConfigurationManager.AppSettings("RutaRecibos")
            If String.IsNullOrEmpty(rutaRecibos) Then
                rutaRecibos = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM", "Recibos", "PagosExtra")
            Else
                rutaRecibos = Path.Combine(rutaRecibos, "PagosExtra")
            End If

            ' Crear directorio específico para pagos extra
            If Not Directory.Exists(rutaRecibos) Then
                Directory.CreateDirectory(rutaRecibos)
            End If

            ' Nombre de archivo con tipo de pago
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")
            Dim tipoPago As String = If(String.IsNullOrEmpty(pagoExtra.TipoPago), "EXTRA", pagoExtra.TipoPago.Replace(" ", "_"))
            Dim nombreArchivo As String = SanitizarNombreArchivo($"ReciboExtra_{tipoPago}_{pagoExtra.NumeroRecibo}_{apartamento.ObtenerCodigoApartamento()}_{timestamp}.pdf")
            rutaArchivo = Path.Combine(rutaRecibos, nombreArchivo)

            ' Generar PDF usando bloques Using para garantizar liberación
            Using docPdf As New Document(PageSize.Letter, 50, 50, 50, 50)
                Using fs As New FileStream(rutaArchivo, FileMode.Create, FileAccess.Write, FileShare.None)
                    Using pdfWriter As PdfWriter = PdfWriter.GetInstance(docPdf, fs)
                        docPdf.Open()

                        ' Generar contenido específico para pago extra
                        AgregarEncabezadoPagoExtra(docPdf, pagoExtra)
                        AgregarInformacionGeneral(docPdf, pagoExtra)
                        AgregarInformacionPropietario(docPdf, apartamento)
                        AgregarDetallesPagoExtra(docPdf, pagoExtra, apartamento)
                        AgregarAdvertenciaLegal(docPdf, pagoExtra)
                        AgregarFirma(docPdf)

                        docPdf.Close()
                    End Using
                End Using
            End Using

            ' Esperar liberación de recursos
            System.Threading.Thread.Sleep(200)
            Return rutaArchivo

        Catch ex As Exception
            MessageBox.Show($"Error al generar recibo de pago extra: {ex.Message}", "Error de PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Genera PDF temporal para envío por correo
    ''' </summary>
    Public Shared Function GenerarReciboPagoExtraTemporal(pagoExtra As PagoModel, apartamento As Apartamento) As String
        Try
            ' Crear carpeta temporal específica
            Dim carpetaTemporal As String = Path.Combine(Path.GetTempPath(), "COOPDIASAM_Temp_PDFs_Extra")
            If Not Directory.Exists(carpetaTemporal) Then
                Directory.CreateDirectory(carpetaTemporal)
            End If

            ' Nombre único para archivo temporal
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMddHHmmssfff")
            Dim tipoPago As String = If(String.IsNullOrEmpty(pagoExtra.TipoPago), "EXTRA", pagoExtra.TipoPago.Replace(" ", "_"))
            Dim nombreArchivo As String = $"TempReciboExtra_{tipoPago}_{pagoExtra.NumeroRecibo}_{timestamp}.pdf"
            Dim rutaArchivo As String = Path.Combine(carpetaTemporal, nombreArchivo)

            ' Usar el método principal pero con ruta específica
            Return GenerarReciboPagoExtraEspecifico(pagoExtra, apartamento, rutaArchivo)

        Catch ex As Exception
            MessageBox.Show($"Error al generar PDF temporal de pago extra: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Genera PDF en ubicación específica
    ''' </summary>
    Public Shared Function GenerarReciboPagoExtraEspecifico(pagoExtra As PagoModel, apartamento As Apartamento, rutaCompleta As String) As String
        Try
            ' Verificar directorio
            Dim directorio As String = Path.GetDirectoryName(rutaCompleta)
            If Not Directory.Exists(directorio) Then
                Directory.CreateDirectory(directorio)
            End If

            Using docPdf As New Document(PageSize.Letter, 50, 50, 50, 50)
                Using fs As New FileStream(rutaCompleta, FileMode.Create, FileAccess.Write, FileShare.None)
                    Using pdfWriter As PdfWriter = PdfWriter.GetInstance(docPdf, fs)
                        docPdf.Open()

                        AgregarEncabezadoPagoExtra(docPdf, pagoExtra)
                        AgregarInformacionGeneral(docPdf, pagoExtra)
                        AgregarInformacionPropietario(docPdf, apartamento)
                        AgregarDetallesPagoExtra(docPdf, pagoExtra, apartamento)
                        AgregarAdvertenciaLegal(docPdf, pagoExtra)
                        AgregarFirma(docPdf)

                        docPdf.Close()
                    End Using
                End Using
            End Using

            Return rutaCompleta

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error generando PDF específico de pago extra: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' ============================================================================
    ' MÉTODOS PARA GENERAR CONTENIDO DEL PDF
    ' ============================================================================

    Private Shared Sub AgregarEncabezadoPagoExtra(document As Document, pagoExtra As PagoModel)
        Dim fontTitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 22, BaseColor.Red)
        Dim fontSubtitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.Gray)
        Dim fontTipo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.Red)

        ' Logo (reutilizar método existente)
        AgregarLogo(document)

        ' Título principal con color diferente para pagos extra
        Dim tituloParrafo As New Paragraph("RECIBO DE PAGO EXTRA", fontTitulo) With {
            .Alignment = Element.ALIGN_CENTER,
            .SpacingAfter = 5.0F
        }
        document.Add(tituloParrafo)

        ' Tipo de pago extra destacado
        Dim tipoPago As String = If(String.IsNullOrEmpty(pagoExtra.TipoPago), "PAGO ADICIONAL", pagoExtra.TipoPago.ToUpper())
        Dim tipoParrafo As New Paragraph($"({tipoPago})", fontTipo) With {
            .Alignment = Element.ALIGN_CENTER,
            .SpacingAfter = 10.0F
        }
        document.Add(tipoParrafo)

        ' Número de recibo
        Dim numeroParrafo As New Paragraph($"No. {pagoExtra.NumeroRecibo}", fontSubtitulo) With {
            .Alignment = Element.ALIGN_CENTER,
            .SpacingAfter = 20.0F
        }
        document.Add(numeroParrafo)

        ' Línea separadora en color
        Dim linea As New LineSeparator(2.0F, 100.0F, BaseColor.Red, Element.ALIGN_CENTER, -2)
        document.Add(New Chunk(linea))
        document.Add(New Paragraph(" "))
    End Sub

    Private Shared Sub AgregarInformacionGeneral(document As Document, pagoExtra As PagoModel)
        Dim fontNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black)
        Dim fontDato As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.Black)

        Dim tableInfo As New PdfPTable(2) With {
            .WidthPercentage = 100,
            .SpacingBefore = 15.0F,
            .SpacingAfter = 15.0F
        }
        tableInfo.SetWidths({1.5F, 2.5F})

        AgregarCeldaTabla(tableInfo, "Fecha de Pago:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tableInfo, pagoExtra.FechaPago.ToString("dd/MM/yyyy"), fontDato, Element.ALIGN_LEFT)

        AgregarCeldaTabla(tableInfo, "Tipo de Pago:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tableInfo, If(String.IsNullOrEmpty(pagoExtra.TipoPago), "PAGO EXTRA", pagoExtra.TipoPago), fontDato, Element.ALIGN_LEFT)

        AgregarCeldaTabla(tableInfo, "Administrador:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tableInfo, "Fernando Gamba", fontDato, Element.ALIGN_LEFT)

        AgregarCeldaTabla(tableInfo, "Conjunto Residencial:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tableInfo, "COOPDIASAM", fontDato, Element.ALIGN_LEFT)

        document.Add(tableInfo)
    End Sub

    Private Shared Sub AgregarInformacionPropietario(document As Document, apartamento As Apartamento)
        Dim fontNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black)
        Dim fontDato As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.Black)

        Dim tablePropietario As New PdfPTable(2) With {
            .WidthPercentage = 100,
            .SpacingBefore = 10.0F,
            .SpacingAfter = 15.0F
        }
        tablePropietario.SetWidths({1.5F, 2.5F})

        AgregarCeldaTabla(tablePropietario, "Propietario:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tablePropietario, If(String.IsNullOrEmpty(apartamento.NombreResidente), "No registrado", apartamento.NombreResidente), fontDato, Element.ALIGN_LEFT)

        AgregarCeldaTabla(tablePropietario, "Apartamento:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tablePropietario, apartamento.ObtenerCodigoApartamento(), fontDato, Element.ALIGN_LEFT)

        AgregarCeldaTabla(tablePropietario, "Matrícula Inmobiliaria:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
        AgregarCeldaTabla(tablePropietario, If(String.IsNullOrEmpty(apartamento.MatriculaInmobiliaria), "No registrada", apartamento.MatriculaInmobiliaria), fontDato, Element.ALIGN_LEFT)

        If Not String.IsNullOrEmpty(apartamento.Telefono) Then
            AgregarCeldaTabla(tablePropietario, "Teléfono:", fontNormal, Element.ALIGN_LEFT, New BaseColor(255, 230, 230))
            AgregarCeldaTabla(tablePropietario, apartamento.Telefono, fontDato, Element.ALIGN_LEFT)
        End If

        document.Add(tablePropietario)
    End Sub

    Private Shared Sub AgregarDetallesPagoExtra(document As Document, pagoExtra As PagoModel, apartamento As Apartamento)
        Dim fontSubtitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Red)
        Dim fontEncabezado As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.White)
        Dim fontNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black)
        Dim fontTotal As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.Red)

        document.Add(New Paragraph("DETALLES DEL PAGO EXTRA", fontSubtitulo) With {
            .Alignment = Element.ALIGN_LEFT,
            .SpacingBefore = 10.0F,
            .SpacingAfter = 10.0F
        })

        Dim tableDetalles As New PdfPTable(2) With {
            .WidthPercentage = 100,
            .SpacingAfter = 15.0F
        }
        tableDetalles.SetWidths({2.5F, 1.0F})

        ' Encabezados con color rojo para pagos extra
        AgregarCeldaTabla(tableDetalles, "Concepto", fontEncabezado, Element.ALIGN_LEFT, New BaseColor(220, 20, 60))
        AgregarCeldaTabla(tableDetalles, "Valor", fontEncabezado, Element.ALIGN_RIGHT, New BaseColor(220, 20, 60))

        ' Detalles del pago extra
        Dim tipoPago As String = If(String.IsNullOrEmpty(pagoExtra.TipoPago), "PAGO EXTRA", pagoExtra.TipoPago)
        AgregarCeldaTabla(tableDetalles, $"Tipo de Pago: {tipoPago}", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaTabla(tableDetalles, "", fontNormal, Element.ALIGN_RIGHT)

        ' Extraer concepto del detalle si existe
        Dim concepto As String = ""
        If Not String.IsNullOrEmpty(pagoExtra.Detalle) AndAlso pagoExtra.Detalle.Contains(":") Then
            Dim partes As String() = pagoExtra.Detalle.Split(":"c)
            If partes.Length > 1 Then
                concepto = partes(1).Trim()
            End If
        End If

        If Not String.IsNullOrEmpty(concepto) Then
            AgregarCeldaTabla(tableDetalles, $"Concepto: {concepto}", fontNormal, Element.ALIGN_LEFT)
            AgregarCeldaTabla(tableDetalles, "", fontNormal, Element.ALIGN_RIGHT)
        End If

        ' Valor del pago extra
        AgregarCeldaTabla(tableDetalles, "Valor del Pago Extra", fontNormal, Element.ALIGN_LEFT)
        AgregarCeldaTabla(tableDetalles, pagoExtra.TotalPagado.ToString("C"), fontNormal, Element.ALIGN_RIGHT)

        ' Si hay observaciones
        If Not String.IsNullOrEmpty(pagoExtra.Observaciones) Then
            AgregarCeldaTabla(tableDetalles, "Observaciones", fontNormal, Element.ALIGN_LEFT)
            AgregarCeldaTabla(tableDetalles, "", fontNormal, Element.ALIGN_RIGHT)
        End If

        ' Línea separadora
        AgregarCeldaTabla(tableDetalles, "", fontNormal, Element.ALIGN_LEFT, New BaseColor(211, 211, 211))
        AgregarCeldaTabla(tableDetalles, "", fontNormal, Element.ALIGN_RIGHT, New BaseColor(211, 211, 211))

        ' Total destacado
        AgregarCeldaTabla(tableDetalles, "TOTAL PAGADO", fontTotal, Element.ALIGN_LEFT)
        AgregarCeldaTabla(tableDetalles, pagoExtra.TotalPagado.ToString("C"), fontTotal, Element.ALIGN_RIGHT)

        document.Add(tableDetalles)

        ' Agregar observaciones detalladas si existen
        If Not String.IsNullOrEmpty(pagoExtra.Observaciones) Then
            Dim fontObservaciones As Font = FontFactory.GetFont(FontFactory.HELVETICA, 9, BaseColor.Black)
            document.Add(New Paragraph("OBSERVACIONES ADICIONALES:", fontSubtitulo) With {
                .SpacingBefore = 10.0F,
                .SpacingAfter = 5.0F
            })
            document.Add(New Paragraph(pagoExtra.Observaciones, fontObservaciones) With {
                .SpacingAfter = 15.0F
            })
        End If
    End Sub

    Private Shared Sub AgregarAdvertenciaLegal(document As Document, pagoExtra As PagoModel)
        Dim fontAdvertencia As Font = FontFactory.GetFont(FontFactory.HELVETICA, 8, BaseColor.Gray)
        Dim fontTitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.Red)

        document.Add(New Paragraph("INFORMACIÓN LEGAL", fontTitulo) With {
            .Alignment = Element.ALIGN_LEFT,
            .SpacingBefore = 20.0F,
            .SpacingAfter = 5.0F
        })

        Dim textoLegal As String = ""

        ' Personalizar mensaje según tipo de pago
        Select Case pagoExtra.TipoPago?.ToUpper()
            Case "MULTA"
                textoLegal = "Este recibo corresponde al pago de una multa impuesta según el reglamento de propiedad horizontal. " &
                           "El pago de esta multa no exime al propietario de futuras sanciones por reincidencia en la falta."
            Case "SANCION"
                textoLegal = "Este recibo corresponde al pago de una sanción económica por incumplimiento del reglamento interno. " &
                           "Se recomienda revisar las normas de convivencia para evitar futuras infracciones."
            Case "ADICION"
                textoLegal = "Este recibo corresponde a un cobro adicional autorizado por la administración del conjunto. " &
                           "Este pago es independiente de las cuotas de administración regulares."
            Case "REPARACION"
                textoLegal = "Este recibo corresponde al cobro por reparaciones o daños causados a las áreas comunes. " &
                           "El propietario es responsable de los daños causados por él, sus familiares o visitantes."
            Case Else
                textoLegal = "Este recibo corresponde a un pago extra autorizado por la administración del conjunto residencial. " &
                           "Para cualquier aclaración, comuníquese con la administración."
        End Select

        document.Add(New Paragraph(textoLegal, fontAdvertencia) With {
            .Alignment = Element.ALIGN_JUSTIFIED,
            .SpacingAfter = 10.0F
        })

        document.Add(New Paragraph("IMPORTANTE: Conserve este recibo como comprobante de pago. " &
                                 "Este documento tiene validez legal y contable.", fontAdvertencia) With {
            .Alignment = Element.ALIGN_JUSTIFIED,
            .SpacingAfter = 15.0F
        })
    End Sub

    ' ============================================================================
    ' MÉTODOS AUXILIARES (REUTILIZAN FUNCIONES DE ReciboPDF)
    ' ============================================================================

    Private Shared Sub AgregarLogo(document As Document)
        Try
            ' Reutilizar la lógica de ReciboPDF para el logo
            Dim rutasLogo As String() = {
                ConfigurationManager.AppSettings("LogoConjunto"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "logo_coopdiasam.png"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "logo_coopdiasam.png"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo_coopdiasam.png"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM", "Images", "logo_coopdiasam.png")
            }

            For Each rutaLogo In rutasLogo
                If Not String.IsNullOrEmpty(rutaLogo) AndAlso File.Exists(rutaLogo) Then
                    Try
                        Dim logo As Image = Image.GetInstance(rutaLogo)
                        logo.ScaleAbsolute(120.0F, 60.0F)
                        logo.Alignment = Element.ALIGN_RIGHT
                        logo.SpacingAfter = 10.0F
                        document.Add(logo)
                        Return
                    Catch logoEx As Exception
                        Continue For
                    End Try
                End If
            Next

            ' Fallback: texto del logo
            Dim fontLogo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.Blue)
            Dim textoLogo As New Paragraph("CONJUNTO RESIDENCIAL COOPDIASAM", fontLogo) With {
                .Alignment = Element.ALIGN_CENTER,
                .SpacingAfter = 10.0F
            }
            document.Add(textoLogo)

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al cargar logo en pago extra: {ex.Message}")
        End Try
    End Sub

    Private Shared Sub AgregarFirma(document As Document)
        Dim fontNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.Black)
        Dim fontDato As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.Black)

        document.Add(New Paragraph(" ", fontNormal) With {.SpacingBefore = 20.0F})

        ' Intentar agregar imagen de firma (reutilizar lógica de ReciboPDF)
        Try
            Dim rutasFirma As String() = {
                ConfigurationManager.AppSettings("FirmaAdministrador"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "firma_administrador.png"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "firma_administrador.png"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "COOPDIASAM", "Images", "firma_administrador.png")
            }

            For Each rutaFirma In rutasFirma
                If Not String.IsNullOrEmpty(rutaFirma) AndAlso File.Exists(rutaFirma) Then
                    Try
                        Dim firma As Image = Image.GetInstance(rutaFirma)
                        firma.ScaleAbsolute(100.0F, 50.0F)
                        firma.Alignment = Element.ALIGN_LEFT
                        document.Add(firma)
                        Exit For
                    Catch
                        Continue For
                    End Try
                End If
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al cargar firma: {ex.Message}")
        End Try

        document.Add(New Paragraph("FIRMADO POR:", fontNormal) With {
            .Alignment = Element.ALIGN_LEFT,
            .SpacingBefore = 10.0F
        })

        document.Add(New Paragraph("Fernando Gamba", fontDato) With {
            .Alignment = Element.ALIGN_LEFT
        })

        document.Add(New Paragraph("Administrador Conjunto Residencial COOPDIASAM", fontNormal) With {
            .Alignment = Element.ALIGN_LEFT
        })

        document.Add(New Paragraph("Teléfono: +57 321-9597100", fontNormal) With {
            .Alignment = Element.ALIGN_LEFT
        })

        document.Add(New Paragraph($"Recibo generado el: {DateTime.Now:dd/MM/yyyy HH:mm:ss}", fontNormal) With {
            .Alignment = Element.ALIGN_RIGHT,
            .SpacingBefore = 20.0F
        })
    End Sub

    Private Shared Sub AgregarCeldaTabla(table As PdfPTable, text As String, font As Font, alignment As Integer, Optional backgroundColor As BaseColor = Nothing)
        Dim cell As New PdfPCell(New Phrase(text, font)) With {
            .HorizontalAlignment = alignment,
            .VerticalAlignment = Element.ALIGN_MIDDLE,
            .Padding = 8,
            .Border = Rectangle.BOX,
            .BorderWidth = 0.5F,
            .BorderColor = BaseColor.Gray
        }

        If backgroundColor IsNot Nothing Then
            cell.BackgroundColor = backgroundColor
        End If

        table.AddCell(cell)
    End Sub

    Private Shared Function SanitizarNombreArchivo(nombreArchivo As String) As String
        Try
            Dim caracteresInvalidos As Char() = Path.GetInvalidFileNameChars()
            For Each caracter In caracteresInvalidos
                nombreArchivo = nombreArchivo.Replace(caracter, "_"c)
            Next
            Return nombreArchivo
        Catch
            Return "ReciboExtra_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"
        End Try
    End Function

End Class