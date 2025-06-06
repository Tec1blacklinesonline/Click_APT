Imports System.Data.SQLite
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Diagnostics

Public Class FormPagos
    Inherits Form

    Private numeroTorre As Integer
    Private dgvPagos As DataGridView
    Private lblTitulo As Label
    Private btnRegistrar As Button
    Private btnCancelar As Button
    Private apartamentos As List(Of Apartamento)
    Private panelBotones As Panel
    Private isUpdating As Boolean = False ' Para evitar eventos recursivos

    Public Sub New(torre As Integer)
        MyBase.New()
        numeroTorre = torre
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Configuración del formulario
        Me.Text = $"Registro de Pagos - Torre {numeroTorre}"
        Me.Size = New Size(1450, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ConfigurarFormulario()
    End Sub

    Private Sub ConfigurarFormulario()
        ' Panel principal
        Dim panelPrincipal As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.FromArgb(240, 240, 240),
            .Padding = New Padding(10)
        }
        Me.Controls.Add(panelPrincipal)

        ' Encabezado
        Dim panelEncabezado As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.FromArgb(52, 152, 219)
        }
        panelPrincipal.Controls.Add(panelEncabezado)

        ' Título
        lblTitulo = New Label With {
            .Text = $"REGISTRO DE PAGOS - TORRE {numeroTorre}",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        panelEncabezado.Controls.Add(lblTitulo)

        ' Botón Volver
        Dim btnVolver As New Button With {
            .Text = "← Volver",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(52, 73, 94),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(100, 35),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(panelEncabezado.Width - 120, 12)
        }
        btnVolver.FlatAppearance.BorderSize = 0
        AddHandler btnVolver.Click, AddressOf btnVolver_Click
        panelEncabezado.Controls.Add(btnVolver)

        ' CREAR ENCABEZADOS MANUALES CON LABELS
        CrearEncabezadosPersonalizados(panelPrincipal)

        ' Panel del DataGridView
        Dim panelGrid As New Panel With {
            .Location = New Point(10, 130),
            .Size = New Size(Me.ClientSize.Width - 30, Me.ClientSize.Height - 200),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelPrincipal.Controls.Add(panelGrid)

        ' DataGridView
        dgvPagos = New DataGridView()

        With dgvPagos
            .Location = New Point(0, 0)
            .Size = New Size(panelGrid.Width - 2, panelGrid.Height - 2)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            .BackgroundColor = Color.White
            .GridColor = Color.LightGray
            .BorderStyle = BorderStyle.None
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .MultiSelect = False
            .ReadOnly = False
            .ScrollBars = ScrollBars.Vertical
            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .DefaultCellStyle.Padding = New Padding(3)
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
            .DefaultCellStyle.SelectionForeColor = Color.White
            .RowTemplate.Height = 28
        End With

        panelGrid.Controls.Add(dgvPagos)

        ' Panel de botones
        panelBotones = New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 60,
            .BackColor = Color.FromArgb(240, 240, 240)
        }
        panelPrincipal.Controls.Add(panelBotones)

        ConfigurarColumnas()
        ConfigurarBotones()
    End Sub

    Private Sub CrearEncabezadosPersonalizados(panelPadre As Panel)
        ' Panel para encabezados manuales
        Dim panelEncabezados As New Panel With {
            .Location = New Point(10, 70),
            .Size = New Size(Me.ClientSize.Width - 30, 50),
            .BackColor = Color.FromArgb(34, 45, 50)
        }
        panelPadre.Controls.Add(panelEncabezados)

        ' Definir anchos de columnas
        Dim anchos() As Integer = {100, 100, 120, 120, 120, 150, 100, 100, 120, 130, 50, 50}
        Dim titulos() As String = {"APARTAMENTO", "FECHA PAGO", "SALDO ANT.", "PAGO ADMIN", "PAGO INTER", "OBSERVAC.", "TOTAL", "INTERESES", "TOTAL GRAL", "No. RECIBO", "✉", "PDF"}

        Dim xPos As Integer = 0
        For i As Integer = 0 To titulos.Length - 1
            Dim lblHeader As New Label With {
                .Text = titulos(i),
                .Font = New Font("Segoe UI", 9, FontStyle.Bold),
                .ForeColor = Color.White,
                .BackColor = Color.FromArgb(34, 45, 50),
                .Location = New Point(xPos, 0),
                .Size = New Size(anchos(i), 50),
                .TextAlign = ContentAlignment.MiddleCenter,
                .BorderStyle = BorderStyle.FixedSingle
            }
            panelEncabezados.Controls.Add(lblHeader)
            xPos += anchos(i)
        Next
    End Sub

    Private Sub ConfigurarColumnas()
        dgvPagos.Columns.Clear()

        ' Configurar columnas
        Dim anchos() As Integer = {100, 100, 120, 120, 120, 150, 100, 100, 120, 130}
        Dim nombres() As String = {"Apartamento", "FechaPago", "SaldoAnterior", "PagoAdministracion", "PagoInteres", "Observaciones", "Total", "Intereses", "TotalGeneral", "NumeroRecibo"}
        Dim editables() As Boolean = {False, True, False, True, True, True, False, True, False, False}
        Dim coloresBack() As Color = {Color.FromArgb(245, 245, 245), Color.LightYellow, Color.FromArgb(245, 245, 245), Color.Yellow, Color.LightYellow, Color.LightYellow, Color.FromArgb(230, 230, 230), Color.LightYellow, Color.FromArgb(255, 255, 200), Color.FromArgb(245, 245, 245)}

        For i As Integer = 0 To nombres.Length - 1
            Dim col As New DataGridViewTextBoxColumn With {
                .Name = nombres(i),
                .Width = anchos(i),
                .ReadOnly = Not editables(i),
                .DefaultCellStyle = New DataGridViewCellStyle With {
                    .BackColor = coloresBack(i),
                    .Padding = New Padding(3)
                }
            }

            ' Configuraciones específicas
            Select Case nombres(i)
                Case "Apartamento"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    col.DefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                Case "FechaPago"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                Case "SaldoAnterior", "PagoAdministracion", "PagoInteres", "Total", "Intereses", "TotalGeneral"
                    col.DefaultCellStyle.Format = "C"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Case "NumeroRecibo"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    col.DefaultCellStyle.Font = New Font("Segoe UI", 8)
            End Select

            dgvPagos.Columns.Add(col)
        Next

        ' Columna oculta para ID
        dgvPagos.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "IdApartamento", .Visible = False})

        ' Columnas de botones
        Dim colBotonCorreo As New DataGridViewButtonColumn With {
            .Name = "BtnCorreo",
            .HeaderText = "✉",
            .Text = "✉",
            .UseColumnTextForButtonValue = True,
            .Width = 50,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(52, 152, 219),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
            }
        }
        dgvPagos.Columns.Add(colBotonCorreo)

        Dim colBotonPDF As New DataGridViewButtonColumn With {
            .Name = "BtnPDF",
            .HeaderText = "PDF",
            .Text = "PDF",
            .UseColumnTextForButtonValue = True,
            .Width = 50,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(231, 76, 60),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold)
            }
        }
        dgvPagos.Columns.Add(colBotonPDF)

        ' Eventos
        AddHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
        AddHandler dgvPagos.CellEndEdit, AddressOf dgvPagos_CellEndEdit
        AddHandler dgvPagos.CellClick, AddressOf dgvPagos_CellClick
        AddHandler dgvPagos.CellPainting, AddressOf dgvPagos_CellPainting
        AddHandler dgvPagos.CellFormatting, AddressOf dgvPagos_CellFormatting
        AddHandler dgvPagos.EditingControlShowing, AddressOf dgvPagos_EditingControlShowing
    End Sub

    Private Sub dgvPagos_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)

            ' Solo procesar clicks en botones
            If e.ColumnIndex = dgvPagos.Columns("BtnCorreo").Index OrElse e.ColumnIndex = dgvPagos.Columns("BtnPDF").Index Then
                ' Verificar si tiene número de recibo
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

                If String.IsNullOrEmpty(numeroRecibo) Then
                    MessageBox.Show("Debe registrar el pago primero para generar el recibo.",
                                  "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Obtener datos del apartamento
                Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                Dim apartamento As Apartamento = Nothing

                For Each apt In apartamentos
                    If apt.IdApartamento = idApartamento Then
                        apartamento = apt
                        Exit For
                    End If
                Next

                If apartamento Is Nothing Then Return

                ' Crear objeto PagoModel
                Dim pago As New PagoModel With {
                    .idApartamento = idApartamento,
                    .numeroRecibo = numeroRecibo,
                    .FechaPago = DateTime.Parse(row.Cells("FechaPago").Value.ToString()),
                    .SaldoAnterior = ConvertirADecimal(row.Cells("SaldoAnterior").Value),
                    .PagoAdministracion = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                    .PagoIntereses = ConvertirADecimal(row.Cells("PagoInteres").Value),
                    .TotalPagado = ConvertirADecimal(row.Cells("Total").Value),
                    .SaldoActual = ConvertirADecimal(row.Cells("SaldoAnterior").Value) - ConvertirADecimal(row.Cells("Total").Value),
                    .Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value),
                    .MatriculaInmobiliaria = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento)
                }
                pago.CuotaActual = pago.PagoAdministracion

                ' Procesar según el botón
                If e.ColumnIndex = dgvPagos.Columns("BtnCorreo").Index Then
                    ProcesarEnvioCorreo(pago, apartamento)
                ElseIf e.ColumnIndex = dgvPagos.Columns("BtnPDF").Index Then
                    ProcesarGeneracionPDF(pago, apartamento)
                End If
            End If
        End If
    End Sub

    Private Sub ProcesarEnvioCorreo(pago As PagoModel, apartamento As Apartamento)
        Try
            If String.IsNullOrEmpty(apartamento.Correo) Then
                MessageBox.Show("No se encuentra el correo en la base de datos. " & vbCrLf &
                              "Por favor actualice la información en la sección de Propietarios.",
                              "Correo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim mensaje As String = $"¿Enviar recibo por correo a {apartamento.Correo}?"
            If MessageBox.Show(mensaje, "Confirmar envío", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.Cursor = Cursors.WaitCursor

                Dim rutaPDF As String = ReciboPDF.GenerarReciboDesdeFormulario(pago, apartamento)

                If Not String.IsNullOrEmpty(rutaPDF) Then
                    If EmailService.EnviarReciboPorCorreo(pago, apartamento, rutaPDF) Then
                        MessageBox.Show($"Recibo enviado correctamente a {apartamento.Correo}",
                                      "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Error al enviar el correo. Verifique la configuración SMTP.",
                                      "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                    Try
                        File.Delete(rutaPDF)
                    Catch
                    End Try
                End If

                Me.Cursor = Cursors.Default
            End If

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show($"Error al enviar correo: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ProcesarGeneracionPDF(pago As PagoModel, apartamento As Apartamento)
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim rutaPDF As String = ReciboPDF.GenerarReciboDesdeFormulario(pago, apartamento)

            If Not String.IsNullOrEmpty(rutaPDF) Then
                Dim mensaje As String = $"PDF generado correctamente en:{vbCrLf}{rutaPDF}{vbCrLf}{vbCrLf}¿Desea abrir el archivo?"

                If MessageBox.Show(mensaje, "PDF Generado", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo(rutaPDF) With {.UseShellExecute = True})
                End If
            End If

            Me.Cursor = Cursors.Default

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show($"Error al generar PDF: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvPagos_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs)
        If e.RowIndex >= 0 AndAlso (e.ColumnIndex = dgvPagos.Columns("BtnCorreo").Index Or
                                     e.ColumnIndex = dgvPagos.Columns("BtnPDF").Index) Then

            Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)
            Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

            If String.IsNullOrEmpty(numeroRecibo) Then
                e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(230, 230, 230)), e.CellBounds)
                e.Graphics.DrawRectangle(New Pen(Color.FromArgb(200, 200, 200)), e.CellBounds)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub ConfigurarBotones()
        ' Botón Registrar
        btnRegistrar = New Button With {
            .Text = "✅ REGISTRAR PAGOS",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(39, 174, 96),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(180, 40),
            .Location = New Point(20, 10)
        }
        btnRegistrar.FlatAppearance.BorderSize = 0
        AddHandler btnRegistrar.Click, AddressOf btnRegistrar_Click
        panelBotones.Controls.Add(btnRegistrar)

        ' Botón Limpiar
        btnCancelar = New Button With {
            .Text = "❌ LIMPIAR",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(231, 76, 60),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(120, 40),
            .Location = New Point(220, 10)
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
        panelBotones.Controls.Add(btnCancelar)

        ' Botón Pago Masivo
        Dim btnPagoMasivo As New Button With {
            .Text = "💰 PAGO MASIVO",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(52, 152, 219),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(140, 40),
            .Location = New Point(360, 10)
        }
        btnPagoMasivo.FlatAppearance.BorderSize = 0
        AddHandler btnPagoMasivo.Click, AddressOf btnPagoMasivo_Click
        panelBotones.Controls.Add(btnPagoMasivo)

        ' Botón Enviar Todos
        Dim btnEnviarTodos As New Button With {
            .Text = "📧 ENVIAR TODOS",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(41, 128, 185),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(140, 40),
            .Location = New Point(520, 10)
        }
        btnEnviarTodos.FlatAppearance.BorderSize = 0
        AddHandler btnEnviarTodos.Click, AddressOf btnEnviarTodos_Click
        panelBotones.Controls.Add(btnEnviarTodos)

        ' Botón PDF Todos
        Dim btnPDFTodos As New Button With {
            .Text = "📑 PDF TODOS",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(192, 57, 43),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(130, 40),
            .Location = New Point(680, 10)
        }
        btnPDFTodos.FlatAppearance.BorderSize = 0
        AddHandler btnPDFTodos.Click, AddressOf btnPDFTodos_Click
        panelBotones.Controls.Add(btnPDFTodos)

        ' Botón Resumen
        Dim btnResumen As New Button With {
            .Text = "📊 RESUMEN",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(155, 89, 182),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(120, 40),
            .Location = New Point(830, 10)
        }
        btnResumen.FlatAppearance.BorderSize = 0
        AddHandler btnResumen.Click, AddressOf btnResumen_Click
        panelBotones.Controls.Add(btnResumen)

        ' Información
        Dim lblInfo As New Label With {
            .Text = "💡 AMARILLO = Editable | GRIS = Solo lectura",
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(127, 140, 141),
            .AutoSize = True,
            .Location = New Point(970, 20)
        }
        panelBotones.Controls.Add(lblInfo)
    End Sub

    Private Sub FormPagos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            CargarApartamentos()
        Catch ex As Exception
            MessageBox.Show($"Error al cargar apartamentos: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarApartamentos()
        Try
            isUpdating = True
            apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(numeroTorre)
            dgvPagos.Rows.Clear()

            For Each apartamento In apartamentos
                Dim fila As Integer = dgvPagos.Rows.Add()
                Dim ultimoSaldo As Decimal = PagosDAL.ObtenerUltimoSaldo(apartamento.IdApartamento)

                dgvPagos.Rows(fila).Cells("IdApartamento").Value = apartamento.IdApartamento
                dgvPagos.Rows(fila).Cells("Apartamento").Value = $"T{numeroTorre}-{apartamento.NumeroApartamento}"
                dgvPagos.Rows(fila).Cells("FechaPago").Value = DateTime.Now.ToString("dd/MM/yyyy")
                dgvPagos.Rows(fila).Cells("SaldoAnterior").Value = ultimoSaldo
                dgvPagos.Rows(fila).Cells("PagoAdministracion").Value = 0
                dgvPagos.Rows(fila).Cells("PagoInteres").Value = 0
                dgvPagos.Rows(fila).Cells("Observaciones").Value = ""
                dgvPagos.Rows(fila).Cells("Intereses").Value = 0
                dgvPagos.Rows(fila).Cells("Total").Value = 0
                dgvPagos.Rows(fila).Cells("TotalGeneral").Value = 0
                dgvPagos.Rows(fila).Cells("NumeroRecibo").Value = ""
            Next

            isUpdating = False
        Catch ex As Exception
            isUpdating = False
            Throw New Exception($"Error al cargar apartamentos: {ex.Message}")
        End Try
    End Sub

    Private Sub dgvPagos_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Not isUpdating AndAlso e.RowIndex >= 0 Then
            CalcularTotales(e.RowIndex)
        End If
    End Sub

    Private Sub dgvPagos_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If Not isUpdating AndAlso e.RowIndex >= 0 Then
            CalcularTotales(e.RowIndex)
        End If
    End Sub

    Private Sub CalcularTotales(rowIndex As Integer)
        If isUpdating Then Return

        Try
            isUpdating = True
            Dim row As DataGridViewRow = dgvPagos.Rows(rowIndex)

            If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                Dim pagoInteres As Decimal = ConvertirADecimal(row.Cells("PagoInteres").Value)
                Dim intereses As Decimal = ConvertirADecimal(row.Cells("Intereses").Value)

                Dim total As Decimal = pagoAdmin + pagoInteres
                row.Cells("Total").Value = total

                Dim totalGeneral As Decimal = total + intereses
                row.Cells("TotalGeneral").Value = totalGeneral

                dgvPagos.InvalidateRow(rowIndex)
            End If

            isUpdating = False
        Catch ex As Exception
            isUpdating = False
        End Try
    End Sub

    Private Function ObtenerValorCelda(valor As Object) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return ""
        End If
        Return valor.ToString()
    End Function

    Private Function ConvertirADecimal(valor As Object) As Decimal
        If valor Is Nothing OrElse IsDBNull(valor) OrElse String.IsNullOrEmpty(valor.ToString()) Then
            Return 0
        End If

        Dim valorStr As String = valor.ToString().Replace("$", "").Replace(",", "").Trim()
        Dim resultado As Decimal
        If Decimal.TryParse(valorStr, resultado) Then
            Return resultado
        End If
        Return 0
    End Function

    Private Function GenerarNumeroRecibo(idApartamento As Integer, fechaPago As DateTime) As String
        Try
            Dim matricula As String = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento)
            Dim numeroRecibo As String
            Dim intentos As Integer = 0

            Do
                Dim fechaFormato As String = fechaPago.ToString("yyMMdd")
                Dim horaFormato As String = fechaPago.AddMinutes(intentos).ToString("HHmm")
                numeroRecibo = $"{matricula}{fechaFormato}{horaFormato}"

                If Not PagosDAL.ExisteNumeroRecibo(numeroRecibo) Then
                    Exit Do
                End If

                intentos += 1
            Loop While intentos < 100

            Return numeroRecibo

        Catch ex As Exception
            Return $"REC{DateTime.Now.Ticks.ToString().Substring(10, 8)}"
        End Try
    End Function

    Private Sub btnRegistrar_Click(sender As Object, e As EventArgs)
        Try
            Dim filasParaRegistrar As Integer = 0
            Dim filasConError As New List(Of Integer)()

            For i As Integer = 0 To dgvPagos.Rows.Count - 1
                Dim row As DataGridViewRow = dgvPagos.Rows(i)

                If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)

                    If pagoAdmin > 0 Then
                        Dim fechaStr As String = ObtenerValorCelda(row.Cells("FechaPago").Value)
                        Dim fecha As DateTime

                        If Not DateTime.TryParse(fechaStr, fecha) Then
                            filasConError.Add(i + 1)
                            Continue For
                        End If

                        If fecha > DateTime.Now Then
                            MessageBox.Show($"La fecha de pago del apartamento en la fila {i + 1} no puede ser futura.",
                                          "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            dgvPagos.CurrentCell = row.Cells("FechaPago")
                            Return
                        End If

                        filasParaRegistrar += 1
                    End If
                End If
            Next

            If filasConError.Count > 0 Then
                MessageBox.Show($"Hay errores en las siguientes filas: {String.Join(", ", filasConError)}" & vbCrLf &
                              "Por favor corrija las fechas antes de continuar.",
                              "Errores de validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If filasParaRegistrar = 0 Then
                MessageBox.Show("No hay pagos para registrar. Ingrese al menos un pago de administración mayor a $0.",
                              "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim resultado As DialogResult = MessageBox.Show(
                $"¿Confirma el registro de {filasParaRegistrar} pago(s)?",
                "Confirmar Registro",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                Me.Cursor = Cursors.WaitCursor

                Dim pagosRegistrados As Integer = RegistrarPagos()

                Me.Cursor = Cursors.Default

                If pagosRegistrados > 0 Then
                    MessageBox.Show($"Se registraron correctamente {pagosRegistrados} pago(s).",
                                  "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' NO recargar, solo actualizar el estado visual
                    ActualizarEstadoVisual()
                Else
                    MessageBox.Show("No se pudo registrar ningún pago.",
                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show($"Error al registrar pagos: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function RegistrarPagos() As Integer
        Dim pagosRegistrados As Integer = 0

        For Each row As DataGridViewRow In dgvPagos.Rows
            Try
                If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)

                    If pagoAdmin > 0 Then
                        Dim fechaPago As DateTime = DateTime.Parse(row.Cells("FechaPago").Value.ToString())
                        Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                        Dim numeroRecibo As String = GenerarNumeroRecibo(idApartamento, fechaPago)

                        Dim pago As New PagoModel With {
                            .idApartamento = idApartamento,
                            .MatriculaInmobiliaria = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento),
                            .fechaPago = fechaPago,
                            .numeroRecibo = numeroRecibo,
                            .SaldoAnterior = ConvertirADecimal(row.Cells("SaldoAnterior").Value),
                            .PagoAdministracion = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                            .PagoIntereses = ConvertirADecimal(row.Cells("PagoInteres").Value),
                            .Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value),
                            .Detalle = $"Pago registrado Torre {numeroTorre}"
                        }

                        pago.CalcularTotales()

                        If PagosDAL.RegistrarPago(pago) Then
                            pagosRegistrados += 1
                            row.Cells("NumeroRecibo").Value = numeroRecibo
                            ColorearFilaRegistrada(row)
                        End If
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine($"Error al registrar pago de fila {row.Index}: {ex.Message}")
                Continue For
            End Try
        Next

        Return pagosRegistrados
    End Function

    Private Sub ColorearFilaRegistrada(row As DataGridViewRow)
        For Each cell As DataGridViewCell In row.Cells
            If cell.ColumnIndex < dgvPagos.Columns("BtnCorreo").Index Then
                cell.Style.BackColor = Color.FromArgb(230, 230, 230)
                cell.ReadOnly = True
            End If
        Next

        row.Cells("BtnCorreo").Style.BackColor = Color.FromArgb(52, 152, 219)
        row.Cells("BtnPDF").Style.BackColor = Color.FromArgb(231, 76, 60)
    End Sub

    Private Sub ActualizarEstadoVisual()
        For Each row As DataGridViewRow In dgvPagos.Rows
            If Not String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                ColorearFilaRegistrada(row)
            End If
        Next
        dgvPagos.Refresh()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Dim resultado As DialogResult = MessageBox.Show(
            "¿Limpiar todos los campos editados?",
            "Confirmar",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If resultado = DialogResult.Yes Then
            CargarApartamentos()
        End If
    End Sub

    Private Sub btnVolver_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub dgvPagos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)

            If e.ColumnIndex < dgvPagos.Columns("BtnCorreo").Index Then
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

                If Not String.IsNullOrEmpty(numeroRecibo) Then
                    e.CellStyle.BackColor = Color.FromArgb(230, 230, 230)
                    e.CellStyle.ForeColor = Color.FromArgb(100, 100, 100)
                Else
                    Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)

                    If pagoAdmin > 0 Then
                        If e.ColumnIndex = dgvPagos.Columns("PagoAdministracion").Index OrElse
                           e.ColumnIndex = dgvPagos.Columns("Total").Index OrElse
                           e.ColumnIndex = dgvPagos.Columns("TotalGeneral").Index Then
                            e.CellStyle.BackColor = Color.FromArgb(200, 255, 200)
                            e.CellStyle.Font = New Font(dgvPagos.DefaultCellStyle.Font, FontStyle.Bold)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub dgvPagos_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs)
        RemoveHandler e.Control.KeyPress, AddressOf NumericCell_KeyPress

        If dgvPagos.CurrentCell.ColumnIndex = dgvPagos.Columns("PagoAdministracion").Index OrElse
           dgvPagos.CurrentCell.ColumnIndex = dgvPagos.Columns("PagoInteres").Index OrElse
           dgvPagos.CurrentCell.ColumnIndex = dgvPagos.Columns("Intereses").Index Then

            AddHandler e.Control.KeyPress, AddressOf NumericCell_KeyPress
        End If
    End Sub

    Private Sub NumericCell_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "."c Then
            e.Handled = True
        End If

        If e.KeyChar = "."c AndAlso CType(sender, TextBox).Text.Contains(".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnPagoMasivo_Click(sender As Object, e As EventArgs)
        Using inputForm As New Form()
            inputForm.Text = "Pago Masivo"
            inputForm.Size = New Size(350, 150)
            inputForm.StartPosition = FormStartPosition.CenterParent
            inputForm.FormBorderStyle = FormBorderStyle.FixedDialog
            inputForm.MaximizeBox = False
            inputForm.MinimizeBox = False

            Dim lblInstruccion As New Label With {
                .Text = "Ingrese el monto a aplicar a todos los apartamentos:",
                .Location = New Point(10, 20),
                .Size = New Size(320, 20)
            }

            Dim txtMonto As New TextBox With {
                .Location = New Point(10, 45),
                .Size = New Size(150, 25),
                .Font = New Font("Segoe UI", 10)
            }

            Dim btnAplicar As New Button With {
                .Text = "Aplicar",
                .Location = New Point(170, 43),
                .Size = New Size(75, 30),
                .DialogResult = DialogResult.OK
            }

            Dim btnCancelar As New Button With {
                .Text = "Cancelar",
                .Location = New Point(250, 43),
                .Size = New Size(75, 30),
                .DialogResult = DialogResult.Cancel
            }

            inputForm.Controls.AddRange({lblInstruccion, txtMonto, btnAplicar, btnCancelar})

            If inputForm.ShowDialog() = DialogResult.OK Then
                Dim monto As Decimal
                If Decimal.TryParse(txtMonto.Text, monto) AndAlso monto > 0 Then
                    AplicarPagoMasivo(monto)
                Else
                    MessageBox.Show("Por favor ingrese un monto válido mayor a 0.",
                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        End Using
    End Sub

    Private Sub AplicarPagoMasivo(monto As Decimal)
        Dim filasAplicadas As Integer = 0

        For Each row As DataGridViewRow In dgvPagos.Rows
            If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                row.Cells("PagoAdministracion").Value = monto
                filasAplicadas += 1
            End If
        Next

        If filasAplicadas > 0 Then
            MessageBox.Show($"Se aplicó el monto ${monto:N0} a {filasAplicadas} apartamento(s).",
                          "Pago Masivo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub btnEnviarTodos_Click(sender As Object, e As EventArgs)
        Try
            Dim recibosParaEnviar As New List(Of DataGridViewRow)()

            ' Recolectar filas con recibos generados
            For Each row As DataGridViewRow In dgvPagos.Rows
                If Not String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                    Dim apartamento As Apartamento = apartamentos.FirstOrDefault(Function(a) a.IdApartamento = idApartamento)

                    If apartamento IsNot Nothing AndAlso Not String.IsNullOrEmpty(apartamento.Correo) Then
                        recibosParaEnviar.Add(row)
                    End If
                End If
            Next

            If recibosParaEnviar.Count = 0 Then
                MessageBox.Show("No hay recibos para enviar o los apartamentos no tienen correo configurado.",
                              "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim mensaje As String = $"Se enviarán {recibosParaEnviar.Count} recibo(s) por correo.{vbCrLf}¿Desea continuar?"
            If MessageBox.Show(mensaje, "Confirmar envío masivo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                Me.Cursor = Cursors.WaitCursor
                Dim enviados As Integer = 0
                Dim errores As Integer = 0

                ' Crear formulario de progreso
                Using frmProgreso As New Form()
                    frmProgreso.Text = "Enviando correos..."
                    frmProgreso.Size = New Size(400, 150)
                    frmProgreso.StartPosition = FormStartPosition.CenterParent
                    frmProgreso.FormBorderStyle = FormBorderStyle.FixedDialog
                    frmProgreso.ControlBox = False

                    Dim lblProgreso As New Label With {
                        .Text = "Preparando envío...",
                        .Location = New Point(20, 20),
                        .Size = New Size(360, 20)
                    }

                    Dim progressBar As New ProgressBar With {
                        .Location = New Point(20, 50),
                        .Size = New Size(360, 30),
                        .Maximum = recibosParaEnviar.Count,
                        .Value = 0
                    }

                    frmProgreso.Controls.AddRange({lblProgreso, progressBar})
                    frmProgreso.Show()
                    Application.DoEvents()

                    For Each row In recibosParaEnviar
                        Try
                            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                            Dim apartamento As Apartamento = apartamentos.FirstOrDefault(Function(a) a.IdApartamento = idApartamento)

                            lblProgreso.Text = $"Enviando a {apartamento.Torre}{apartamento.NumeroApartamento} ({apartamento.Correo})..."
                            Application.DoEvents()

                            ' Crear objeto pago
                            Dim pago As New PagoModel With {
                                .idApartamento = idApartamento,
                                .NumeroRecibo = ObtenerValorCelda(row.Cells("NumeroRecibo").Value),
                                .FechaPago = DateTime.Parse(row.Cells("FechaPago").Value.ToString()),
                                .SaldoAnterior = ConvertirADecimal(row.Cells("SaldoAnterior").Value),
                                .PagoAdministracion = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                                .PagoIntereses = ConvertirADecimal(row.Cells("PagoInteres").Value),
                                .TotalPagado = ConvertirADecimal(row.Cells("Total").Value),
                                .SaldoActual = ConvertirADecimal(row.Cells("SaldoAnterior").Value) - ConvertirADecimal(row.Cells("Total").Value),
                                .Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value),
                                .MatriculaInmobiliaria = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento)
                            }
                            pago.CuotaActual = pago.PagoAdministracion

                            ' Generar PDF temporal
                            Dim rutaPDF As String = ReciboPDF.GenerarReciboDesdeFormulario(pago, apartamento)

                            If Not String.IsNullOrEmpty(rutaPDF) Then
                                If EmailService.EnviarReciboPorCorreo(pago, apartamento, rutaPDF) Then
                                    enviados += 1
                                Else
                                    errores += 1
                                End If

                                Try
                                    File.Delete(rutaPDF)
                                Catch
                                End Try
                            End If

                        Catch ex As Exception
                            errores += 1
                        End Try

                        progressBar.Value += 1
                        Application.DoEvents()
                    Next

                    frmProgreso.Close()
                End Using

                Me.Cursor = Cursors.Default

                Dim resumen As String = $"Proceso completado:{vbCrLf}{vbCrLf}"
                resumen &= $"Correos enviados: {enviados}{vbCrLf}"
                resumen &= $"Errores: {errores}"

                MessageBox.Show(resumen, "Envío masivo completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show($"Error durante el envío masivo: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPDFTodos_Click(sender As Object, e As EventArgs)
        Try
            Dim recibosParaGenerar As New List(Of DataGridViewRow)()

            ' Recolectar filas con recibos generados
            For Each row As DataGridViewRow In dgvPagos.Rows
                If Not String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    recibosParaGenerar.Add(row)
                End If
            Next

            If recibosParaGenerar.Count = 0 Then
                MessageBox.Show("No hay recibos para generar. Primero debe registrar los pagos.",
                              "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Seleccionar carpeta de destino
            Using folderDialog As New FolderBrowserDialog()
                folderDialog.Description = "Seleccione la carpeta donde guardar los PDFs"
                folderDialog.ShowNewFolderButton = True

                If folderDialog.ShowDialog() = DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim generados As Integer = 0
                    Dim errores As Integer = 0

                    ' Crear subcarpeta con fecha
                    Dim carpetaDestino As String = Path.Combine(folderDialog.SelectedPath,
                                                               $"Recibos_Torre{numeroTorre}_{DateTime.Now:yyyyMMdd_HHmmss}")
                    Directory.CreateDirectory(carpetaDestino)

                    For Each row In recibosParaGenerar
                        Try
                            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                            Dim apartamento As Apartamento = apartamentos.FirstOrDefault(Function(a) a.IdApartamento = idApartamento)

                            If apartamento IsNot Nothing Then
                                ' Crear objeto pago
                                Dim pago As New PagoModel With {
                                    .idApartamento = idApartamento,
                                    .NumeroRecibo = ObtenerValorCelda(row.Cells("NumeroRecibo").Value),
                                    .FechaPago = DateTime.Parse(row.Cells("FechaPago").Value.ToString()),
                                    .SaldoAnterior = ConvertirADecimal(row.Cells("SaldoAnterior").Value),
                                    .PagoAdministracion = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                                    .PagoIntereses = ConvertirADecimal(row.Cells("PagoInteres").Value),
                                    .TotalPagado = ConvertirADecimal(row.Cells("Total").Value),
                                    .SaldoActual = ConvertirADecimal(row.Cells("SaldoAnterior").Value) - ConvertirADecimal(row.Cells("Total").Value),
                                    .Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value),
                                    .MatriculaInmobiliaria = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento)
                                }
                                pago.CuotaActual = pago.PagoAdministracion

                                ' Generar PDF
                                Dim nombreArchivo As String = $"Recibo_{pago.NumeroRecibo}_T{numeroTorre}-{apartamento.NumeroApartamento}.pdf"
                                Dim rutaPDF As String = Path.Combine(carpetaDestino, nombreArchivo)

                                If ReciboPDF.GenerarReciboPDF(pago, apartamento, rutaPDF) Then
                                    generados += 1
                                Else
                                    errores += 1
                                End If
                            End If

                        Catch ex As Exception
                            errores += 1
                        End Try
                    Next

                    Me.Cursor = Cursors.Default

                    Dim mensaje As String = $"Proceso completado:{vbCrLf}{vbCrLf}"
                    mensaje &= $"PDFs generados: {generados}{vbCrLf}"
                    mensaje &= $"Errores: {errores}{vbCrLf}{vbCrLf}"
                    mensaje &= $"Los archivos se guardaron en:{vbCrLf}{carpetaDestino}{vbCrLf}{vbCrLf}"
                    mensaje &= "¿Desea abrir la carpeta?"

                    If MessageBox.Show(mensaje, "PDFs generados", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                        Process.Start("explorer.exe", carpetaDestino)
                    End If
                End If
            End Using

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show($"Error al generar PDFs: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnResumen_Click(sender As Object, e As EventArgs)
        Dim resumen As String = GenerarResumenPagos()

        Using formResumen As New Form()
            formResumen.Text = "Resumen de Pagos"
            formResumen.Size = New Size(500, 400)
            formResumen.StartPosition = FormStartPosition.CenterParent

            Dim txtResumen As New TextBox With {
                .Multiline = True,
                .ReadOnly = True,
                .Text = resumen,
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10),
                .BackColor = Color.White,
                .ScrollBars = ScrollBars.Vertical
            }

            formResumen.Controls.Add(txtResumen)
            formResumen.ShowDialog()
        End Using
    End Sub

    Private Function GenerarResumenPagos() As String
        Dim totalAdmin As Decimal = 0
        Dim totalIntereses As Decimal = 0
        Dim totalGeneral As Decimal = 0
        Dim apartamentosPagaron As Integer = 0
        Dim apartamentosPendientes As Integer = 0

        For Each row As DataGridViewRow In dgvPagos.Rows
            If Not String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                totalAdmin += ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                totalIntereses += ConvertirADecimal(row.Cells("PagoInteres").Value)
                totalGeneral += ConvertirADecimal(row.Cells("TotalGeneral").Value)
                apartamentosPagaron += 1
            Else
                Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                If pagoAdmin > 0 Then
                    apartamentosPendientes += 1
                End If
            End If
        Next

        Dim resumen As String = $"RESUMEN DE PAGOS - TORRE {numeroTorre}" & vbCrLf
        resumen &= "=" & New String("=", 40) & vbCrLf & vbCrLf

        resumen &= "ESTADÍSTICAS:" & vbCrLf
        resumen &= $"Total apartamentos: {apartamentos.Count}" & vbCrLf
        resumen &= $"Apartamentos con pago registrado: {apartamentosPagaron}" & vbCrLf
        resumen &= $"Apartamentos con pago pendiente: {apartamentosPendientes}" & vbCrLf
        resumen &= $"Apartamentos sin pago: {apartamentos.Count - apartamentosPagaron - apartamentosPendientes}" & vbCrLf & vbCrLf

        resumen &= "TOTALES RECAUDADOS:" & vbCrLf
        resumen &= $"Total Administración: ${totalAdmin:N0}" & vbCrLf
        resumen &= $"Total Intereses: ${totalIntereses:N0}" & vbCrLf
        resumen &= $"Total General: ${totalGeneral:N0}" & vbCrLf & vbCrLf

        If apartamentosPagaron > 0 Then
            resumen &= "PROMEDIOS:" & vbCrLf
            resumen &= $"Promedio por apartamento: ${totalGeneral / apartamentosPagaron:N0}" & vbCrLf
        End If

        Return resumen
    End Function

    Private Function IsValidEmail(email As String) As Boolean
        Try
            Dim addr As New System.Net.Mail.MailAddress(email)
            Return addr.Address = email
        Catch
            Return False
        End Try
    End Function




End Class