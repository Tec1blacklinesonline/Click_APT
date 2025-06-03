Imports System.Data.SQLite
Imports System.Windows.Forms
Imports System.Drawing

Public Class FormPagos
    Inherits Form

    Private numeroTorre As Integer
    Private dgvPagos As DataGridView
    Private lblTitulo As Label
    Private btnRegistrar As Button
    Private btnCancelar As Button
    Private apartamentos As List(Of Apartamento)
    Private panelBotones As Panel

    Public Sub New(torre As Integer)
        MyBase.New()
        numeroTorre = torre
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Configuración del formulario
        Me.Text = $"Registro de Pagos - Torre {numeroTorre}"
        Me.Size = New Size(1400, 700)
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

        ' Panel del DataGridView sin encabezados nativos
        Dim panelGrid As New Panel With {
            .Location = New Point(10, 130),
            .Size = New Size(Me.ClientSize.Width - 30, Me.ClientSize.Height - 200),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelPrincipal.Controls.Add(panelGrid)

        ' DataGridView SIN encabezados (los haremos con labels)
        dgvPagos = New DataGridView()

        With dgvPagos
            .Location = New Point(0, 0)
            .Size = New Size(panelGrid.Width - 2, panelGrid.Height - 2)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            .BackgroundColor = Color.White
            .GridColor = Color.LightGray
            .BorderStyle = BorderStyle.None

            ' DESACTIVAR encabezados nativos
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False

            ' Configuración básica
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .MultiSelect = False
            .ReadOnly = False
            .ScrollBars = ScrollBars.Vertical

            ' Estilo de celdas
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

        ' Definir anchos de columnas (deben coincidir con ConfigurarColumnas)
        Dim anchos() As Integer = {100, 100, 120, 120, 120, 150, 100, 100, 120, 130}
        Dim titulos() As String = {"APARTAMENTO", "FECHA PAGO", "SALDO ANT.", "PAGO ADMIN", "PAGO INTER", "OBSERVAC.", "TOTAL", "INTERESES", "TOTAL GRAL", "No. RECIBO"}

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

        ' Configurar columnas con anchos exactos que coincidan con encabezados
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

            ' Configuraciones específicas por tipo de columna
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

        ' Eventos
        AddHandler dgvPagos.CellValueChanged, AddressOf dgvPagos_CellValueChanged
        AddHandler dgvPagos.CellEndEdit, AddressOf dgvPagos_CellEndEdit
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

        ' Información
        Dim lblInfo As New Label With {
            .Text = "💡 Campos AMARILLOS = Editables | Campos GRISES = Solo lectura | TOTAL se calcula automáticamente",
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(127, 140, 141),
            .AutoSize = True,
            .Location = New Point(360, 20)
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
            apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(numeroTorre)
            dgvPagos.Rows.Clear()

            For Each apartamento In apartamentos
                Dim fila As Integer = dgvPagos.Rows.Add()

                ' Obtener el último saldo
                Dim ultimoSaldo As Decimal = PagosDAL.ObtenerUltimoSaldo(apartamento.IdApartamento)

                ' Llenar datos
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

        Catch ex As Exception
            Throw New Exception($"Error al cargar apartamentos: {ex.Message}")
        End Try
    End Sub

    ' Eventos del DataGridView
    Private Sub dgvPagos_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            CalcularTotales(e.RowIndex)
        End If
    End Sub

    Private Sub dgvPagos_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            CalcularTotales(e.RowIndex)
        End If
    End Sub

    Private Sub CalcularTotales(rowIndex As Integer)
        Try
            Dim row As DataGridViewRow = dgvPagos.Rows(rowIndex)

            If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)
                Dim pagoInteres As Decimal = ConvertirADecimal(row.Cells("PagoInteres").Value)
                Dim intereses As Decimal = ConvertirADecimal(row.Cells("Intereses").Value)

                ' Calcular totales
                Dim total As Decimal = pagoAdmin + pagoInteres
                row.Cells("Total").Value = total

                Dim totalGeneral As Decimal = total + intereses
                row.Cells("TotalGeneral").Value = totalGeneral

                dgvPagos.InvalidateRow(rowIndex)
            End If

        Catch ex As Exception
            ' Ignorar errores
        End Try
    End Sub

    ' Funciones auxiliares
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

        Dim resultado As Decimal
        If Decimal.TryParse(valor.ToString(), resultado) Then
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

    ' Botones principales
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
                $"¿Confirma el registro de {filasParaRegistrar} pago(s)?",
                "Confirmar Registro",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.Yes Then
                RegistrarPagos()
                MessageBox.Show("Pagos registrados exitosamente.", "Éxito",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                CargarApartamentos()
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al registrar pagos: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RegistrarPagos()
        For Each row As DataGridViewRow In dgvPagos.Rows
            Try
                If String.IsNullOrEmpty(ObtenerValorCelda(row.Cells("NumeroRecibo").Value)) Then
                    Dim pagoAdmin As Decimal = ConvertirADecimal(row.Cells("PagoAdministracion").Value)

                    If pagoAdmin > 0 Then
                        Dim fechaPago As DateTime = DateTime.Parse(row.Cells("FechaPago").Value.ToString())
                        Dim idApartamento As Integer = Convert.ToInt32(row.Cells("IdApartamento").Value)
                        Dim numeroRecibo As String = GenerarNumeroRecibo(idApartamento, fechaPago)

                        Dim pago As New PagoModel With {
                            .IdApartamento = idApartamento,
                            .MatriculaInmobiliaria = PagosDAL.ObtenerMatriculaInmobiliaria(idApartamento),
                            .FechaPago = fechaPago,
                            .NumeroRecibo = numeroRecibo,
                            .SaldoAnterior = ConvertirADecimal(row.Cells("SaldoAnterior").Value),
                            .PagoAdministracion = ConvertirADecimal(row.Cells("PagoAdministracion").Value),
                            .PagoIntereses = ConvertirADecimal(row.Cells("PagoInteres").Value),
                            .Observaciones = ObtenerValorCelda(row.Cells("Observaciones").Value),
                            .Detalle = $"Pago registrado Torre {numeroTorre}"
                        }

                        pago.CalcularTotales()

                        If PagosDAL.RegistrarPago(pago) Then
                            row.Cells("NumeroRecibo").Value = numeroRecibo
                        End If
                    End If
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
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

End Class