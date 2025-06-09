' ReciboPDF.vb
' Este archivo maneja la generación de recibos de pago en formato PDF.

Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Diagnostics ' Para abrir el PDF automáticamente
Imports System.Windows.Forms ' Para MessageBox (considerar logging para producción)

Public Class ReciboPDF

    ''' <summary>
    ''' Genera un recibo de pago en formato PDF.
    ''' </summary>
    ''' <param name="pago">El objeto PagoModel que contiene la información del pago.</param>
    ''' <param name="apartamento">El objeto Apartamento que contiene la información del apartamento.</param>
    ''' <returns>La ruta completa del archivo PDF generado, o String.Empty si hay un error.</returns>
    Public Shared Function GenerarReciboDePago(pago As PagoModel, apartamento As Apartamento) As String
        Dim rutaArchivo As String = String.Empty

        Try
            ' Ruta donde se guardará el PDF (por defecto, la carpeta de descargas del usuario)
            Dim folderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim downloadsPath As String = Path.Combine(folderPath, "Downloads")

            If Not Directory.Exists(downloadsPath) Then
                Directory.CreateDirectory(downloadsPath)
            End If

            ' Nombre del archivo PDF
            Dim nombreArchivo As String = $"Recibo_Pago_No_{pago.NumeroRecibo}_{apartamento.NumeroApartamento}.pdf"
            rutaArchivo = Path.Combine(downloadsPath, nombreArchivo)

            ' Configuración del documento PDF
            Dim document As New Document(PageSize.LETTER, 50, 50, 50, 50) ' Margen
            Dim writer As PdfWriter = PdfWriter.GetInstance(document, New FileStream(rutaArchivo, FileMode.Create))

            document.Open()

            ' --- Estilos y Fuentes ---
            Dim fontTitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18, BaseColor.DARK_GRAY)
            Dim fontSubtitulo As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.DARK_GRAY)
            Dim fontEncabezadoTabla As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.WHITE)
            Dim fontNormal As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.BLACK)
            Dim fontDatoClave As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)
            Dim fontTotal As Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 11, BaseColor.BLUE)

            ' Color de fondo para encabezados de tabla
            Dim colorEncabezadoTabla As BaseColor = New BaseColor(66, 133, 244) ' Azul Google

            ' --- Encabezado del Recibo (Logo y Título) ---
            ' Añadir Logo (Necesitarás un logo.png o .jpg en una ruta accesible)
            ' Por ahora, usaremos un placeholder. Cuando tengas el logo, actualiza la ruta.
            ' Si el logo está en la carpeta de recursos del proyecto, puedes usar:
            ' Dim logoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "logo_coopdiasam.png")
            ' O simplemente una ruta conocida.
            Dim logoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "logo_coopdiasam.png") ' Asume que el logo está en la misma carpeta del ejecutable o en una subcarpeta "logo"
            If File.Exists(logoPath) Then
                Try
                    Dim logo As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(logoPath)
                    logo.ScaleAbsolute(100.0F, 50.0F) ' Ajusta el tamaño del logo
                    logo.Alignment = Element.ALIGN_RIGHT
                    document.Add(logo)
                Catch ex As Exception
                    ' Manejar error si el logo no se puede cargar, pero no detener la generación del PDF
                    MessageBox.Show("No se pudo cargar el logo: " & ex.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try
            Else
                MessageBox.Show($"Advertencia: No se encontró el archivo del logo en la ruta: {logoPath}. El PDF se generará sin logo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ' Título del recibo
            document.Add(New Paragraph("RECIBO DE CAJA", fontTitulo) With {.Alignment = Element.ALIGN_CENTER})
            document.Add(New Paragraph($"No. {pago.NumeroRecibo}", fontSubtitulo) With {.Alignment = Element.ALIGN_CENTER})
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            ' --- Sección 1: Información General del Recibo ---
            Dim tableInfoGeneral As New PdfPTable(2) ' 2 columnas
            tableInfoGeneral.WidthPercentage = 100
            tableInfoGeneral.SetWidths({1.5F, 2.5F}) ' Ancho relativo de las columnas
            tableInfoGeneral.SpacingBefore = 10.0F

            AddCellToTable(tableInfoGeneral, "Fecha:", fontNormal, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
            AddCellToTable(tableInfoGeneral, pago.FechaPago.ToString("dd/MM/yyyy"), fontDatoClave, Element.ALIGN_LEFT)

            AddCellToTable(tableInfoGeneral, "Administrador:", fontNormal, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
            AddCellToTable(tableInfoGeneral, "Fernando Gamba", fontDatoClave, Element.ALIGN_LEFT) ' Esto podría venir de una configuración o tabla de usuarios

            AddCellToTable(tableInfoGeneral, "Conjunto Residencial:", fontNormal, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
            AddCellToTable(tableInfoGeneral, "COOPDIASAM", fontDatoClave, Element.ALIGN_LEFT) ' Esto podría venir de una configuración

            document.Add(tableInfoGeneral)
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            ' --- Sección 2: Información del Propietario / Apartamento ---
            Dim tablePropietario As New PdfPTable(2)
            tablePropietario.WidthPercentage = 100
            tablePropietario.SetWidths({1.5F, 2.5F})
            tablePropietario.SpacingBefore = 10.0F

            AddCellToTable(tablePropietario, "Propietario:", fontNormal, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
            AddCellToTable(tablePropietario, apartamento.NombreResidente, fontDatoClave, Element.ALIGN_LEFT)

            AddCellToTable(tablePropietario, "Apartamento:", fontNormal, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
            AddCellToTable(tablePropietario, apartamento.ObtenerCodigoApartamento(), fontDatoClave, Element.ALIGN_LEFT)

            AddCellToTable(tablePropietario, "Matrícula Inmobiliaria:", fontNormal, Element.ALIGN_LEFT, BaseColor.LIGHT_GRAY)
            AddCellToTable(tablePropietario, apartamento.MatriculaInmobiliaria, fontDatoClave, Element.ALIGN_LEFT)

            document.Add(tablePropietario)
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            ' --- Sección 3: Detalles del Pago ---
            document.Add(New Paragraph("DETALLES DEL PAGO", fontSubtitulo) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            Dim tableDetalles As New PdfPTable(2)
            tableDetalles.WidthPercentage = 100
            tableDetalles.SetWidths({2.0F, 1.0F})

            AddCellToTable(tableDetalles, "Concepto", fontEncabezadoTabla, Element.ALIGN_LEFT, colorEncabezadoTabla)
            AddCellToTable(tableDetalles, "Valor", fontEncabezadoTabla, Element.ALIGN_RIGHT, colorEncabezadoTabla)

            AddCellToTable(tableDetalles, "Saldo Anterior", fontNormal, Element.ALIGN_LEFT)
            AddCellToTable(tableDetalles, pago.SaldoAnterior.ToString("C2"), fontNormal, Element.ALIGN_RIGHT)

            AddCellToTable(tableDetalles, "Valor Pagado Administración", fontNormal, Element.ALIGN_LEFT)
            AddCellToTable(tableDetalles, pago.PagoAdministracion.ToString("C2"), fontNormal, Element.ALIGN_RIGHT)

            AddCellToTable(tableDetalles, "Valor Pagado Intereses", fontNormal, Element.ALIGN_LEFT)
            AddCellToTable(tableDetalles, pago.PagoIntereses.ToString("C2"), fontNormal, Element.ALIGN_RIGHT)

            AddCellToTable(tableDetalles, "Cuota Actual", fontNormal, Element.ALIGN_LEFT)
            AddCellToTable(tableDetalles, pago.CuotaActual.ToString("C2"), fontNormal, Element.ALIGN_RIGHT)

            AddCellToTable(tableDetalles, "Total Pagado", fontTotal, Element.ALIGN_LEFT)
            AddCellToTable(tableDetalles, pago.TotalPagado.ToString("C2"), fontTotal, Element.ALIGN_RIGHT)

            document.Add(tableDetalles)
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            ' --- Sección 4: Saldos ---
            document.Add(New Paragraph("SALDOS", fontSubtitulo) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            Dim tableSaldos As New PdfPTable(2)
            tableSaldos.WidthPercentage = 100
            tableSaldos.SetWidths({2.0F, 1.0F})

            AddCellToTable(tableSaldos, "Concepto", fontEncabezadoTabla, Element.ALIGN_LEFT, colorEncabezadoTabla)
            AddCellToTable(tableSaldos, "Valor", fontEncabezadoTabla, Element.ALIGN_RIGHT, colorEncabezadoTabla)

            ' Calcular el saldo de intereses (podría ser 0 si ya se pagó o no hay mora)
            Dim saldoIntereses As Decimal = ApartamentoDAL.ObtenerTotalInteresesCalculados(apartamento.IdApartamento) ' Asumo un método en ApartamentoDAL
            If pago.PagoIntereses > 0 AndAlso pago.PagoIntereses <= saldoIntereses Then
                saldoIntereses = saldoIntereses - pago.PagoIntereses ' Si se pagaron intereses, reducir el saldo
            ElseIf pago.PagoIntereses > 0 Then
                saldoIntereses = 0 ' Si el pago de intereses cubre o excede, se asume 0
            End If

            AddCellToTable(tableSaldos, "Intereses", fontNormal, Element.ALIGN_LEFT)
            AddCellToTable(tableSaldos, saldoIntereses.ToString("C2"), fontNormal, Element.ALIGN_RIGHT)

            AddCellToTable(tableSaldos, "Cuota", fontNormal, Element.ALIGN_LEFT)
            AddCellToTable(tableSaldos, pago.SaldoActual.ToString("C2"), fontNormal, Element.ALIGN_RIGHT) ' Saldo actual es el saldo de cuotas pendientes

            AddCellToTable(tableSaldos, "Total Saldo Pendiente", fontTotal, Element.ALIGN_LEFT)
            AddCellToTable(tableSaldos, (saldoIntereses + pago.SaldoActual).ToString("C2"), fontTotal, Element.ALIGN_RIGHT)

            document.Add(tableSaldos)
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio


            ' --- Sección 5: Observaciones ---
            document.Add(New Paragraph("OBSERVACIONES", fontSubtitulo) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio
            document.Add(New Paragraph(pago.Observaciones, fontNormal))
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            ' --- Sección 6: Firma ---
            document.Add(New Paragraph("FIRMADO POR:", fontNormal) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph("Fernando Gamba", fontDatoClave) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph("Teléfono: +57 321-9597100", fontNormal) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph("Cargo: Administrador Conjunto Residencial COOPDIASAM", fontNormal) With {.Alignment = Element.ALIGN_LEFT})
            document.Add(New Paragraph(" ", fontNormal)) ' Espacio

            ' Añadir espacio para la firma escaneada (si es necesario)
            ' Por ahora, solo un placeholder

            Dim firmaPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "firma_administrador.png")
            If File.Exists(firmaPath) Then
                Dim firma As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(firmaPath)
                firma.ScaleAbsolute(80.0F, 40.0F)
                firma.Alignment = Element.ALIGN_LEFT
                document.Add(firma)
            End If

            document.Close()

            ' Abrir el PDF automáticamente
            Process.Start(New ProcessStartInfo(rutaArchivo) With {.UseShellExecute = True})

            Return rutaArchivo

        Catch ex As Exception
            MessageBox.Show($"Error al generar el recibo PDF: {ex.Message}", "Error de PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Método auxiliar para añadir celdas a una tabla PdfPTable.
    ''' </summary>
    Private Shared Sub AddCellToTable(table As PdfPTable, text As String, font As Font, alignment As Integer, Optional backgroundColor As BaseColor = Nothing)
        Dim cell As New PdfPCell(New Phrase(text, font))
        cell.HorizontalAlignment = alignment
        cell.VerticalAlignment = Element.ALIGN_MIDDLE
        cell.Padding = 5
        cell.Border = Rectangle.NO_BORDER ' Sin bordes por defecto, puedes ajustarlo
        If backgroundColor IsNot Nothing Then
            cell.BackgroundColor = backgroundColor
        End If
        table.AddCell(cell)
    End Sub

    ''' <summary>
    ''' Método auxiliar para añadir celdas a una tabla PdfPTable con rowspan.
    ''' </summary>
    Private Shared Sub AddCellToTable(table As PdfPTable, text As String, font As Font, alignment As Integer, colspan As Integer, Optional backgroundColor As BaseColor = Nothing)
        Dim cell As New PdfPCell(New Phrase(text, font))
        cell.HorizontalAlignment = alignment
        cell.VerticalAlignment = Element.ALIGN_MIDDLE
        cell.Padding = 5
        cell.Colspan = colspan
        cell.Border = Rectangle.NO_BORDER ' Sin bordes por defecto, puedes ajustarlo
        If backgroundColor IsNot Nothing Then
            cell.BackgroundColor = backgroundColor
        End If
        table.AddCell(cell)
    End Sub

    ' Necesitarás este método en ApartamentoDAL para obtener el total de intereses calculados si aún no existe.
    ' Agrega esto en ApartamentoDAL.vb
    ' Public Shared Function ObtenerTotalInteresesCalculados(idApartamento As Integer) As Decimal
    '     Dim totalIntereses As Decimal = 0D
    '     Try
    '         Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
    '             conexion.Open()
    '             Dim consulta As String = "SELECT COALESCE(SUM(valor_interes), 0) FROM calculos_interes WHERE id_apartamentos = @idApartamento AND pagado = 0"
    '             Using comando As New SQLiteCommand(consulta, conexion)
    '                 comando.Parameters.AddWithValue("@idApartamento", idApartamento)
    '                 Dim resultado = comando.ExecuteScalar()
    '                 If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
    '                     totalIntereses = Convert.ToDecimal(resultado)
    '                 End If
    '             End Using
    '         End Using
    '     Catch ex As Exception
    '         MessageBox.Show("Error al obtener el total de intereses calculados: " & ex.Message, "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '     End Try
    '     Return totalIntereses
    ' End Function

End Class