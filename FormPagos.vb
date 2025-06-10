' ============================================================================
' FORMULARIO DE PAGOS CORREGIDO
' Versión corregida que implementa correctamente el cálculo de intereses
' y generación de número de recibo según especificaciones
' ============================================================================

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Linq
Imports System.Diagnostics

Public Class FormPagos
    Inherits Form
    Private numeroTorre As Integer
    Private apartamentos As List(Of Apartamento)
    Private dgvPagos As DataGridView
    Private btnRegistrar As Button
    Private btnCancelar As Button
    Private panelBotones As Panel

    Public Sub New(numeroTorre As Integer)
        Me.numeroTorre = numeroTorre
        InitializeComponent()
        ConfigurarFormulario()
        CargarApartamentos()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.Text = "Registro de Pagos - Torre " & numeroTorre.ToString()
        Me.Size = New Size(1400, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(250, 250, 250)
        Me.WindowState = FormWindowState.Maximized
        Me.ResumeLayout(False)
    End Sub

    Private Sub ConfigurarFormulario()
        ' Header
        Dim headerPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(52, 152, 219)
        }

        Dim lblTitulo As New Label With {
            .Text = "📋 REGISTRO DE PAGOS - TORRE " & numeroTorre.ToString(),
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .ForeColor = Color.White,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Fill
        }
        headerPanel.Controls.Add(lblTitulo)

        ' DataGridView
        dgvPagos = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(44, 62, 80),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold),
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            },
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .Font = New Font("Segoe UI", 9),
                .SelectionBackColor = Color.FromArgb(52, 152, 219),
                .SelectionForeColor = Color.White
            },
            .RowHeadersVisible = False,
            .AllowUserToResizeRows = False,
            .ColumnHeadersHeight = 40,
            .RowTemplate = New DataGridViewRow With {.Height = 30}
        }

        ConfigurarColumnas()
        AddHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
        AddHandler dgvPagos.CellClick, AddressOf dgvPagos_CellClick

        ' Panel de botones
        panelBotones = New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 70,
            .BackColor = Color.FromArgb(236, 240, 241)
        }

        ConfigurarBotones()

        Me.Controls.Add(dgvPagos)
        Me.Controls.Add(panelBotones)
        Me.Controls.Add(headerPanel)
    End Sub

    Private Sub ConfigurarColumnas()
        With dgvPagos.Columns
            .Add("IdApartamento", "ID")
            .Add("Apartamento", "APARTAMENTO")
            .Add("FechaPago", "FECHA PAGO")
            .Add("SaldoAnterior", "SALDO ANT.")
            .Add("PagoAdministracion", "PAGO ADMIN")
            .Add("PagoInteres", "PAGO INTER")
            .Add("Observaciones", "OBSERVAC.")
            .Add("Total", "TOTAL")
            .Add("Intereses", "INTERESES")
            .Add("TotalGeneral", "TOTAL GRAL")
            .Add("NumeroRecibo", "No. RECIBO")
        End With

        ' Configurar propiedades de columnas
        dgvPagos.Columns("IdApartamento").Visible = False
        dgvPagos.Columns("Apartamento").Width = 100
        dgvPagos.Columns("FechaPago").Width = 100
        dgvPagos.Columns("SaldoAnterior").Width = 120
        dgvPagos.Columns("PagoAdministracion").Width = 100
        dgvPagos.Columns("PagoInteres").Width = 100
        dgvPagos.Columns("Observaciones").Width = 150
        dgvPagos.Columns("Total").Width = 100
        dgvPagos.Columns("Intereses").Width = 100
        dgvPagos.Columns("TotalGeneral").Width = 100
        dgvPagos.Columns("NumeroRecibo").Width = 120

        ' Configurar campos editables
        dgvPagos.Columns("SaldoAnterior").ReadOnly = True
        dgvPagos.Columns("Total").ReadOnly = True
        dgvPagos.Columns("TotalGeneral").ReadOnly = True
        dgvPagos.Columns("NumeroRecibo").ReadOnly = True

        ' Configurar formato de moneda
        For Each columna As String In {"SaldoAnterior", "PagoAdministracion", "PagoInteres", "Total", "Intereses", "TotalGeneral"}
            dgvPagos.Columns(columna).DefaultCellStyle.Format = "C"
            dgvPagos.Columns(columna).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next

        ' Agregar botones de correo y PDF
        Dim btnCorreoColumn As New DataGridViewButtonColumn With {
            .Name = "BtnCorreo",
            .HeaderText = "✉",
            .Text = "✉",
            .UseColumnTextForButtonValue = True,
            .Width = 40,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = Color.FromArgb(39, 174, 96),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
            }
        }

        Dim btnPDFColumn As New DataGridViewButtonColumn With {
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

        dgvPagos.Columns.Add(btnCorreoColumn)
        dgvPagos.Columns.Add(btnPDFColumn)
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

        ' Botón Generar PDF
        Dim btnGenerarPDF As New Button With {
            .Text = "📄 GENERAR PDF",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(255, 165, 0),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(150, 40),
            .Location = New Point(360, 10)
        }
        btnGenerarPDF.FlatAppearance.BorderSize = 0
        AddHandler btnGenerarPDF.Click, AddressOf btnGenerarPDF_Click
        panelBotones.Controls.Add(btnGenerarPDF)

        ' Botón Enviar Correo
        Dim btnEnviarCorreo As New Button With {
            .Text = "✉️ ENVIAR CORREO",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(0, 128, 0),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(160, 40),
            .Location = New Point(530, 10)
        }
        btnEnviarCorreo.FlatAppearance.BorderSize = 0
        AddHandler btnEnviarCorreo.Click, AddressOf btnEnviarCorreo_Click
        panelBotones.Controls.Add(btnEnviarCorreo)

        ' Botón Volver
        Dim btnVolver As New Button With {
            .Text = "← VOLVER",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(127, 140, 141),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(100, 40),
            .Location = New Point(710, 10)
        }
        btnVolver.FlatAppearance.BorderSize = 0
        AddHandler btnVolver.Click, AddressOf btnVolver_Click
        panelBotones.Controls.Add(btnVolver)

        ' Información
        Dim lblInfo As New Label With {
            .Text = "💡 Campos AMARILLOS = Editables | Campos GRISES = Solo lectura | TOTAL se calcula automáticamente",
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(127, 140, 141),
            .AutoSize = True,
            .Location = New Point(830, 20)
        }
        panelBotones.Controls.Add(lblInfo)
    End Sub

    Private Sub CargarApartamentos()
        Try
            apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(numeroTorre)
            dgvPagos.Rows.Clear()

            For Each apartamento In apartamentos
                Dim fila As Integer = dgvPagos.Rows.Add()

                ' CORREGIDO: Obtener el último saldo usando PagosDAL
                Dim ultimoSaldo As Decimal = PagosDAL.ObtenerUltimoSaldo(apartamento.IdApartamento)

                ' Llenar datos básicos
                dgvPagos.Rows(fila).Cells("IdApartamento").Value = apartamento.IdApartamento
                dgvPagos.Rows(fila).Cells("Apartamento").Value = "T" & numeroTorre.ToString() & "-" & apartamento.NumeroApartamento
                dgvPagos.Rows(fila).Cells("FechaPago").Value = DateTime.Now.ToString("dd/MM/yyyy")
                dgvPagos.Rows(fila).Cells("SaldoAnterior").Value = ultimoSaldo
                dgvPagos.Rows(fila).Cells("PagoAdministracion").Value = 0
                dgvPagos.Rows(fila).Cells("PagoInteres").Value = 0
                dgvPagos.Rows(fila).Cells("Observaciones").Value = ""
                dgvPagos.Rows(fila).Cells("Intereses").Value = 0
                dgvPagos.Rows(fila).Cells("Total").Value = 0
                dgvPagos.Rows(fila).Cells("TotalGeneral").Value = 0
                dgvPagos.Rows(fila).Cells("NumeroRecibo").Value = ""

                ' Configurar colores de campos editables
                dgvPagos.Rows(fila).Cells("PagoAdministracion").Style.BackColor = Color.Yellow
                dgvPagos.Rows(fila).Cells("PagoInteres").Style.BackColor = Color.LightYellow
                dgvPagos.Rows(fila).Cells("Observaciones").Style.BackColor = Color.LightYellow
                dgvPagos.Rows(fila).Cells("Intereses").Style.BackColor = Color.LightYellow
            Next

        Catch ex As Exception
            MessageBox.Show("Error al cargar apartamentos: " & ex.Message, "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' EVENTO PRINCIPAL CORREGIDO - Cálculo automático de intereses
    Private Sub dgvPagos_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Try
                Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)
                Dim columnaEditada As String = dgvPagos.Columns(e.ColumnIndex).Name

                ' Solo procesar si la fila no tiene número de recibo (no está registrada)
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)
                If Not String.IsNullOrEmpty(numeroRecibo) Then
                    Return ' No procesar filas ya registradas
                End If

                ' Obtener el ID del apartamento desde la fila
                Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)

                ' Buscar el apartamento en la lista cargada
                Dim apartamento As Apartamento = Nothing
                For Each apt In apartamentos
                    If apt.IdApartamento = idApartamento Then
                        apartamento = apt
                        Exit For
                    End If
                Next

                If apartamento Is Nothing Then Return

                ' Solo recalcular si se edita PagoAdministracion o SaldoAnterior
                If columnaEditada = "PagoAdministracion" OrElse columnaEditada = "SaldoAnterior" Then

                    ' Obtener valores actuales del DataGridView
                    Dim pagoAdministracion As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                    Dim saldoAnterior As Decimal = ConvertirADecimal(row.Cells("SaldoAnterior").Value)

                    ' --- CÁLCULO AUTOMÁTICO DE INTERESES CORREGIDO VB.NET ---
                    Dim montoIntereses As Decimal = 0D

                    ' Actualizar el saldo actual del apartamento para el cálculo
                    apartamento.SaldoActual = saldoAnterior

                    ' Solo calcular intereses si hay saldo pendiente
                    If apartamento.SaldoActual > 0 Then
                        Try
                            ' CORREGIDO: Usar la estructura CuotaPendienteInfo implementada
                            Dim cuotaInfo As CuotasDAL.CuotaPendienteInfo = CuotasDAL.ObtenerCuotaPendienteMasAntigua(apartamento.IdApartamento)

                            If cuotaInfo.ExisteCuotaPendiente Then
                                ' Obtener la tasa de interés actual
                                Dim tasaInteresMoraAnual As Decimal = ParametrosDAL.ObtenerTasaInteresMoraActual()

                                ' Calcular días en mora del apartamento
                                Dim diasEnMora As Integer = cuotaInfo.DiasVencida

                                If diasEnMora > 0 AndAlso tasaInteresMoraAnual > 0 Then
                                    ' Fórmula de interés de mora: Capital * TasaAnual(%) / 100 * (DíasMora / 365)
                                    montoIntereses = (cuotaInfo.ValorCuota * tasaInteresMoraAnual / 100D) * (diasEnMora / 365D)
                                    montoIntereses = Math.Round(montoIntereses, 2) ' Redondear a 2 decimales
                                End If
                            End If
                        Catch ex As Exception
                            ' Si hay error al calcular intereses, mantener en 0
                            montoIntereses = 0D
                            Debug.WriteLine("Error calculando intereses: " & ex.Message)
                        End Try
                    End If

                    ' Asignar el monto de intereses calculado sin disparar eventos
                    RemoveHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
                    row.Cells("PagoInteres").Value = montoIntereses
                    AddHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
                End If

                ' --- RECALCULAR TOTALES SIEMPRE QUE CAMBIEN LOS VALORES ---
                If columnaEditada = "PagoAdministracion" OrElse
                   columnaEditada = "SaldoAnterior" OrElse
                   columnaEditada = "PagoInteres" OrElse
                   columnaEditada = "Intereses" Then

                    Dim pagoAdministracion As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                    Dim pagoIntereses As Decimal = ConvertirADecimal(row.Cells("PagoInteres").Value)
                    Dim interesesAdicionales As Decimal = ConvertirADecimal(row.Cells("Intereses").Value)

                    ' Calcular totales
                    Dim totalPago As Decimal = pagoAdministracion + pagoIntereses
                    Dim totalGeneral As Decimal = totalPago + interesesAdicionales

                    ' Actualizar las celdas calculadas sin disparar eventos
                    RemoveHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
                    row.Cells("Total").Value = totalPago
                    row.Cells("TotalGeneral").Value = totalGeneral
                    AddHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged

                    ' Forzar actualización visual solo de las celdas necesarias
                    dgvPagos.InvalidateCell(row.Cells("PagoInteres"))
                    dgvPagos.InvalidateCell(row.Cells("Total"))
                    dgvPagos.InvalidateCell(row.Cells("TotalGeneral"))
                End If

            Catch ex As Exception
                ' En caso de error, no interrumpir la operación
                Debug.WriteLine("Error en dgvPagos_CellValueChanged: " & ex.Message)
                MessageBox.Show("Error en cálculo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub

    Private Sub dgvPagos_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        ' Solo procesar clics en los botones de correo y PDF
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim nombreColumna As String = dgvPagos.Columns(e.ColumnIndex).Name

            ' Solo procesar si se hizo clic en los botones de correo o PDF
            If nombreColumna = "BtnCorreo" OrElse nombreColumna = "BtnPDF" Then
                Dim row As DataGridViewRow = dgvPagos.Rows(e.RowIndex)

                ' Verificar si la fila tiene número de recibo generado
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

                If String.IsNullOrEmpty(numeroRecibo) Then
                    MessageBox.Show("Debe registrar el pago primero para generar el recibo.",
                                  "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Procesar según el botón clickeado
                If nombreColumna = "BtnPDF" Then
                    GenerarPDFDesdeDataGrid(row)
                ElseIf nombreColumna = "BtnCorreo" Then
                    EnviarCorreoDesdeDataGrid(row)
                End If
            End If
        End If
    End Sub

    Private Sub btnRegistrar_Click(sender As Object, e As EventArgs)
        Try
            Dim filasParaRegistrar As Integer = 0
            For Each row As DataGridViewRow In dgvPagos.Rows
                If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                    If pagoAdmin > 0 Then
                        filasParaRegistrar += 1
                    End If
                End If
            Next

            If filasParaRegistrar = 0 Then
                MessageBox.Show("No hay pagos para registrar. Ingrese al menos un pago de administración mayor a $0.", "Aviso",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim resultado As DialogResult = MessageBox.Show(
                "¿Confirma el registro de " & filasParaRegistrar.ToString() & " pago(s)?",
                "Confirmar Registro",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                RegistrarPagos()
                MessageBox.Show("Pagos registrados exitosamente.", "Éxito",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                ActualizarVistaPostRegistro()
            End If

        Catch ex As Exception
            MessageBox.Show("Error al registrar pagos: " & ex.Message, "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RegistrarPagos()
        For Each row As DataGridViewRow In dgvPagos.Rows
            Try
                ' Solo registrar si no tiene número de recibo y hay pago de administración
                If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)

                    If pagoAdmin > 0 Then
                        ' CORREGIDO: Generar número de recibo según especificaciones
                        Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                        Dim numeroRecibo As String = GenerarNumeroRecibo(idApartamento)

                        ' Crear objeto PagoModel compatible con tu estructura
                        Dim pago As New PagoModel With {
                            .IdApartamento = idApartamento,
                            .NumeroRecibo = numeroRecibo,
                            .FechaPago = DateTime.Parse(row.Cells("FechaPago").Value.ToString()),
                            .SaldoAnterior = ConvertirADecimal(row.Cells("SaldoAnterior").Value),
                            .PagoAdministracion = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                            .PagoIntereses = ConvertirADecimal(row.Cells("PagoInteres").Value),
                            .CuotaActual = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                            .TotalPagado = ConvertirADecimal(row.Cells("Total").Value),
                            .SaldoActual = ConvertirADecimal(row.Cells("SaldoAnterior").Value) - ConvertirADecimal(row.Cells("Total").Value),
                            .Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value),
                            .EstadoPago = "REGISTRADO",
                            .UsuarioRegistro = "Sistema"
                        }

                        ' CORREGIDO: Usar PagosDAL.RegistrarPago implementado
                        If PagosDAL.RegistrarPago(pago) Then
                            row.Cells("NumeroRecibo").Value = numeroRecibo

                            ' Colorear la fila de gris para indicar que está registrada
                            For Each cell As DataGridViewCell In row.Cells
                                If cell.ColumnIndex < dgvPagos.Columns("BtnCorreo").Index Then
                                    cell.Style.BackColor = Color.FromArgb(230, 230, 230)
                                    cell.Style.ForeColor = Color.FromArgb(80, 80, 80)
                                End If
                            Next

                            ' Marcar las celdas editables como de solo lectura visualmente
                            row.Cells("FechaPago").ReadOnly = True
                            row.Cells("PagoAdministracion").ReadOnly = True
                            row.Cells("PagoInteres").ReadOnly = True
                            row.Cells("Observaciones").ReadOnly = True
                            row.Cells("Intereses").ReadOnly = True
                        End If
                    End If
                End If
            Catch ex As Exception
                ' Continuar con la siguiente fila si hay error
                MessageBox.Show("Error al registrar pago: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Continue For
            End Try
        Next
    End Sub

    ' CORREGIDO: Generar número de recibo según especificaciones de la documentación
    ' Formato: matricula_inmobiliaria + fecha + hora (ej: 1851442505031425)
    Private Function GenerarNumeroRecibo(idApartamento As Integer) As String
        Try
            ' Obtener matrícula inmobiliaria usando PagosDAL
            Dim matriculaInmobiliaria As String = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento)

            ' Si no hay matrícula, usar ID del apartamento como respaldo
            If String.IsNullOrEmpty(matriculaInmobiliaria) Then
                matriculaInmobiliaria = idApartamento.ToString().PadLeft(6, "0"c)
            End If

            ' Generar fecha y hora en formato compacto (YYMMDDHHmm)
            Dim fechaHora As String = DateTime.Now.ToString("yyMMddHHmm")

            ' Combinar según especificación: matricula + fecha + hora
            Return matriculaInmobiliaria & fechaHora

        Catch ex As Exception
            ' Fallback en caso de error: timestamp + random
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMddHHmmss")
            Dim random As New Random()
            Dim numeroAleatorio As Integer = random.Next(100, 999)
            Return timestamp & numeroAleatorio.ToString()
        End Try
    End Function

    ' Resto de métodos sin cambios significativos...
    Private Sub ActualizarVistaPostRegistro()
        Try
            ' Recorrer todas las filas y actualizar solo las que no tienen número de recibo
            For Each row As DataGridViewRow In dgvPagos.Rows
                Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

                ' Si la fila no tiene número de recibo, resetear solo los campos editables
                If String.IsNullOrEmpty(numeroRecibo) Then
                    row.Cells("PagoAdministracion").Value = 0
                    row.Cells("PagoInteres").Value = 0
                    row.Cells("Observaciones").Value = ""
                    row.Cells("Intereses").Value = 0
                    row.Cells("Total").Value = 0
                    row.Cells("TotalGeneral").Value = 0

                    ' Restaurar colores originales para campos editables
                    row.Cells("PagoAdministracion").Style.BackColor = Color.Yellow
                    row.Cells("PagoInteres").Style.BackColor = Color.LightYellow
                    row.Cells("Observaciones").Style.BackColor = Color.LightYellow
                    row.Cells("Intereses").Style.BackColor = Color.LightYellow
                End If
            Next

            ' Forzar actualización visual
            dgvPagos.Invalidate()

        Catch ex As Exception
            ' En caso de error, recargar apartamentos como fallback
            CargarApartamentos()
        End Try
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Try
            Dim resultado As DialogResult = MessageBox.Show(
                "¿Está seguro de que desea limpiar todos los campos?",
                "Confirmar Limpieza",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                CargarApartamentos()
            End If

        Catch ex As Exception
            MessageBox.Show("Error al limpiar campos: " & ex.Message, "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' GENERAR PDF DESDE BOTÓN DEL PANEL
    Private Sub btnGenerarPDF_Click(sender As Object, e As EventArgs)
        Try
            If dgvPagos.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = dgvPagos.SelectedRows(0)
                Dim numeroRecibo As String = ObtenerValorCelda(selectedRow.Cells("NumeroRecibo").Value)

                If Not String.IsNullOrEmpty(numeroRecibo) Then
                    GenerarPDFDesdeDataGrid(selectedRow)
                Else
                    MessageBox.Show("Debe registrar el pago primero para poder generar el recibo PDF.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("Seleccione una fila para generar el recibo PDF.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al generar PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ENVIAR CORREO DESDE BOTÓN DEL PANEL
    Private Sub btnEnviarCorreo_Click(sender As Object, e As EventArgs)
        Try
            If dgvPagos.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = dgvPagos.SelectedRows(0)
                Dim numeroRecibo As String = ObtenerValorCelda(selectedRow.Cells("NumeroRecibo").Value)

                If Not String.IsNullOrEmpty(numeroRecibo) Then
                    EnviarCorreoDesdeDataGrid(selectedRow)
                Else
                    MessageBox.Show("Debe registrar el pago primero para poder enviar el recibo por correo.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("Seleccione una fila para enviar el recibo por correo.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error al intentar enviar el correo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' GENERAR PDF DESDE DATAGRID
    Private Sub GenerarPDFDesdeDataGrid(row As DataGridViewRow)
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

            ' Obtener pago completo desde la base de datos usando PagosDAL
            Dim pago As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If pago IsNot Nothing AndAlso apartamento IsNot Nothing Then
                Dim rutaPdfGenerado As String = ReciboPDF.GenerarReciboDePago(pago, apartamento)

                If Not String.IsNullOrEmpty(rutaPdfGenerado) Then
                    Dim mensaje As String = "Recibo PDF generado y guardado en:" & vbCrLf & rutaPdfGenerado & vbCrLf & vbCrLf & "¿Desea abrir el archivo?"

                    If MessageBox.Show(mensaje, "PDF Generado", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                        Process.Start(New ProcessStartInfo(rutaPdfGenerado) With {.UseShellExecute = True})
                    End If
                Else
                    MessageBox.Show("No se pudo generar el recibo PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("No se encontraron los datos completos del pago o apartamento para generar el recibo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Me.Cursor = Cursors.Default

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("Error al generar PDF: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ENVIAR CORREO DESDE DATAGRID
    Private Sub EnviarCorreoDesdeDataGrid(row As DataGridViewRow)
        Try
            Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
            Dim numeroRecibo As String = ObtenerValorCelda(row.Cells("NumeroRecibo").Value)

            ' Obtener pago y apartamento completos desde la base de datos
            Dim pago As PagoModel = PagosDAL.ObtenerPagoPorNumeroRecibo(numeroRecibo)
            Dim apartamento As Apartamento = ApartamentoDAL.ObtenerApartamentoPorId(idApartamento)

            If pago IsNot Nothing AndAlso apartamento IsNot Nothing Then
                ' Verificar si el apartamento tiene correo
                If String.IsNullOrEmpty(apartamento.Correo) Then
                    MessageBox.Show("No se encuentra el correo en la base de datos." & vbCrLf &
                                  "Por favor actualice la información en la sección de Propietarios.",
                                  "Correo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' Mostrar mensaje de confirmación
                Dim mensaje As String = "¿Enviar recibo por correo a " & apartamento.Correo & "?"
                If MessageBox.Show(mensaje, "Confirmar envío", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                    Me.Cursor = Cursors.WaitCursor

                    ' 1. Generar el PDF primero
                    Dim rutaPdfGenerado As String = ReciboPDF.GenerarReciboDePago(pago, apartamento)

                    If Not String.IsNullOrEmpty(rutaPdfGenerado) Then
                        ' 2. Enviar por correo usando EmailService
                        If EmailService.EnviarRecibo(apartamento.Correo, apartamento.NombreResidente, numeroRecibo, rutaPdfGenerado) Then
                            MessageBox.Show("Correo enviado exitosamente.", "Envío de Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            ' 3. Eliminar PDF temporal
                            Try
                                System.IO.File.Delete(rutaPdfGenerado)
                            Catch
                                ' No es crítico si no se puede eliminar
                            End Try
                        Else
                            MessageBox.Show("No se pudo enviar el correo. Verifique la configuración y el correo del destinatario.", "Error de Correo", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    Else
                        MessageBox.Show("No se pudo generar el PDF para el correo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                    Me.Cursor = Cursors.Default
                End If
            Else
                MessageBox.Show("No se encontraron los datos completos del pago o apartamento para enviar el recibo por correo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("Error al enviar correo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnVolver_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' MÉTODOS AUXILIARES
    Private Function ObtenerValorCelda(valor As Object) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return ""
        End If
        Return valor.ToString().Trim()
    End Function

    Private Function ConvertirADecimal(valor As Object) As Decimal
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return 0D
        End If

        Dim valorString As String = valor.ToString().Trim()
        If String.IsNullOrEmpty(valorString) Then
            Return 0D
        End If

        Dim resultado As Decimal
        If Decimal.TryParse(valorString, resultado) Then
            Return resultado
        End If

        Return 0D
    End Function

End Class