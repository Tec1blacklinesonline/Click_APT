' ============================================================================
' FORMULARIO DE ESTADOS DE CUENTA
' Muestra el estado actual de todos los apartamentos
' ============================================================================

Imports System.Data.SQLite
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class FormEstados
    Private dtEstados As DataTable
    Private apartamentos As List(Of Apartamento)

    ' Controles del formulario
    Private WithEvents dgvEstados As DataGridView
    Private WithEvents cmbTorre As ComboBox
    Private WithEvents cmbEstado As ComboBox
    Private WithEvents txtBuscar As TextBox
    Private WithEvents btnBuscar As Button
    Private WithEvents btnLimpiar As Button
    Private WithEvents btnExportar As Button
    Private lblResumen As Label

    Private Sub FormEstados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarDatos()
    End Sub

    Private Sub ConfigurarFormulario()
        ' Form
        Me.Text = "Estados de Cuenta - COOPDIASAM"
        Me.Size = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Panel superior
        Dim panelSuperior As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(52, 73, 94)
        }

        Dim lblTitulo As New Label With {
            .Text = "ESTADOS DE CUENTA",
            .Font = New Font("Arial", 20, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 25)
        }
        panelSuperior.Controls.Add(lblTitulo)

        ' Panel de filtros
        Dim panelFiltros As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 100,
            .BackColor = Color.FromArgb(236, 240, 241),
            .Padding = New Padding(10)
        }

        ' ComboBox Torre
        Dim lblTorre As New Label With {
            .Text = "Torre:",
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblTorre)

        cmbTorre = New ComboBox With {
            .Location = New Point(70, 17),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        panelFiltros.Controls.Add(cmbTorre)

        ' ComboBox Estado
        Dim lblEstado As New Label With {
            .Text = "Estado:",
            .Location = New Point(250, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblEstado)

        cmbEstado = New ComboBox With {
            .Location = New Point(300, 17),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        panelFiltros.Controls.Add(cmbEstado)

        ' TextBox Búsqueda
        Dim lblBuscar As New Label With {
            .Text = "Buscar:",
            .Location = New Point(480, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblBuscar)

        txtBuscar = New TextBox With {
            .Location = New Point(530, 17),
            .Size = New Size(200, 25),
            .Font = New Font("Arial", 10)
        }
        panelFiltros.Controls.Add(txtBuscar)

        ' Botón Buscar
        btnBuscar = New Button With {
            .Text = "🔍 Buscar",
            .Location = New Point(740, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnBuscar.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnBuscar)

        ' Botón Limpiar
        btnLimpiar = New Button With {
            .Text = "🔄 Limpiar",
            .Location = New Point(850, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(149, 165, 166),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnLimpiar.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnLimpiar)

        ' Botón Exportar
        btnExportar = New Button With {
            .Text = "📊 Exportar",
            .Location = New Point(960, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnExportar.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnExportar)

        ' Panel de resumen
        Dim panelResumen As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 50,
            .BackColor = Color.FromArgb(44, 62, 80)
        }

        lblResumen = New Label With {
            .Dock = DockStyle.Fill,
            .ForeColor = Color.White,
            .Font = New Font("Arial", 11, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter
        }
        panelResumen.Controls.Add(lblResumen)

        ' DataGridView
        dgvEstados = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .RowHeadersVisible = False
        }

        ' Agregar controles al formulario
        Me.Controls.Add(dgvEstados)
        Me.Controls.Add(panelResumen)
        Me.Controls.Add(panelFiltros)
        Me.Controls.Add(panelSuperior)

        ' Configurar eventos
        ConfigurarEventos()

        ' Configurar DataGridView
        ConfigurarDataGridView()
    End Sub

    Private Sub ConfigurarEventos()
        ' Eventos
        AddHandler btnBuscar.Click, AddressOf BtnBuscar_Click
        AddHandler btnLimpiar.Click, AddressOf BtnLimpiar_Click
        AddHandler btnExportar.Click, AddressOf BtnExportar_Click
        AddHandler txtBuscar.KeyPress, AddressOf TxtBuscar_KeyPress
        AddHandler dgvEstados.CellFormatting, AddressOf DgvEstados_CellFormatting
        AddHandler cmbTorre.SelectedIndexChanged, AddressOf FiltrosChanged
        AddHandler cmbEstado.SelectedIndexChanged, AddressOf FiltrosChanged
    End Sub

    Private Sub ConfigurarDataGridView()
        dgvEstados.EnableHeadersVisualStyles = False
        dgvEstados.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 73, 94)
        dgvEstados.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvEstados.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
        dgvEstados.ColumnHeadersHeight = 35
        dgvEstados.DefaultCellStyle.Font = New Font("Arial", 9)
        dgvEstados.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
        dgvEstados.RowTemplate.Height = 30
    End Sub

    Private Sub CargarDatos()
        Try
            ' Cargar torres
            CargarTorres()

            ' Cargar estados
            CargarEstados()

            ' Cargar datos principales
            apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()

            If apartamentos IsNot Nothing AndAlso apartamentos.Count > 0 Then
                CrearDataTable()
                ActualizarDataGridView(apartamentos)
                ActualizarResumen(apartamentos)
            Else
                MessageBox.Show("No se encontraron apartamentos.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al cargar datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarTorres()
        Try
            cmbTorre.Items.Clear()
            cmbTorre.Items.Add("Todas")

            Dim listaTorres = TorresDAL.ObtenerTodasLasTorres()
            If listaTorres IsNot Nothing Then
                For Each torre In listaTorres
                    cmbTorre.Items.Add($"Torre {torre.IdTorre}")
                Next
            End If

            cmbTorre.SelectedIndex = 0
        Catch ex As Exception
            MessageBox.Show($"Error al cargar torres: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarEstados()
        cmbEstado.Items.Clear()
        cmbEstado.Items.AddRange({"Todos", "Al día", "Pendiente", "Saldo a favor"})
        cmbEstado.SelectedIndex = 0
    End Sub

    Private Sub CrearDataTable()
        dtEstados = New DataTable()
        dtEstados.Columns.Add("Torre", GetType(String))
        dtEstados.Columns.Add("Apartamento", GetType(String))
        dtEstados.Columns.Add("NombreResidente", GetType(String))
        dtEstados.Columns.Add("SaldoActual", GetType(Decimal))
        dtEstados.Columns.Add("Estado", GetType(String))
        dtEstados.Columns.Add("UltimoPago", GetType(DateTime))
        dtEstados.Columns.Add("DiasEnMora", GetType(Integer))
        dtEstados.Columns.Add("TotalIntereses", GetType(Decimal))
        dtEstados.Columns.Add("Telefono", GetType(String))
        dtEstados.Columns.Add("Correo", GetType(String))
    End Sub

    Private Sub ActualizarDataGridView(listaApartamentos As List(Of Apartamento))
        dtEstados.Clear()

        For Each apto In listaApartamentos
            Dim row As DataRow = dtEstados.NewRow()
            row("Torre") = $"Torre {apto.Torre}"
            row("Apartamento") = apto.NumeroApartamento
            row("NombreResidente") = apto.NombreResidente
            row("SaldoActual") = apto.SaldoActual
            row("Estado") = apto.ObtenerEstadoCuenta()

            ' Obtener último pago
            Dim ultimoPago = PagosDAL.ObtenerUltimoPago(apto.IdApartamento)
            If ultimoPago IsNot Nothing Then
                row("UltimoPago") = ultimoPago.FechaPago
            Else
                row("UltimoPago") = DBNull.Value
            End If

            ' Calcular días en mora e intereses
            Dim diasMora As Integer = 0
            Dim totalIntereses As Decimal = 0

            Dim cuotasPendientes = CuotasDAL.ObtenerCuotasPendientesPorApartamento(apto.IdApartamento)
            If cuotasPendientes IsNot Nothing Then
                For Each cuota In cuotasPendientes
                    If cuota.FechaVencimiento < DateTime.Now Then
                        Dim diasDiferencia As Integer = (DateTime.Now.Date - cuota.FechaVencimiento.Date).Days
                        diasMora = Math.Max(diasMora, diasDiferencia)
                    End If
                    totalIntereses += cuota.InteresesMora
                Next
            End If

            row("DiasEnMora") = diasMora
            row("TotalIntereses") = totalIntereses
            row("Telefono") = If(String.IsNullOrEmpty(apto.Telefono), "No registrado", apto.Telefono)
            row("Correo") = If(String.IsNullOrEmpty(apto.Correo), "No registrado", apto.Correo)

            dtEstados.Rows.Add(row)
        Next

        dgvEstados.DataSource = dtEstados
        FormatearColumnas()
    End Sub

    Private Sub FormatearColumnas()
        If dgvEstados.Columns.Count > 0 Then
            dgvEstados.Columns("Torre").Width = 80
            dgvEstados.Columns("Apartamento").Width = 100
            dgvEstados.Columns("NombreResidente").Width = 200
            dgvEstados.Columns("SaldoActual").DefaultCellStyle.Format = "C0"
            dgvEstados.Columns("SaldoActual").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvEstados.Columns("Estado").Width = 120
            dgvEstados.Columns("UltimoPago").DefaultCellStyle.Format = "dd/MM/yyyy"
            dgvEstados.Columns("DiasEnMora").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvEstados.Columns("TotalIntereses").DefaultCellStyle.Format = "C0"
            dgvEstados.Columns("TotalIntereses").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            ' Encabezados
            dgvEstados.Columns("NombreResidente").HeaderText = "Residente"
            dgvEstados.Columns("SaldoActual").HeaderText = "Saldo Actual"
            dgvEstados.Columns("UltimoPago").HeaderText = "Último Pago"
            dgvEstados.Columns("DiasEnMora").HeaderText = "Días Mora"
            dgvEstados.Columns("TotalIntereses").HeaderText = "Intereses"
            dgvEstados.Columns("Telefono").HeaderText = "Teléfono"
            dgvEstados.Columns("Correo").HeaderText = "Correo"
        End If
    End Sub

    Private Sub ActualizarResumen(apartamentos As List(Of Apartamento))
        Try
            Dim totalApartamentos As Integer = apartamentos.Count
            Dim alDia As Integer = apartamentos.Where(Function(a) a.ObtenerEstadoCuenta() = "Al día").Count
            Dim pendientes As Integer = apartamentos.Where(Function(a) a.ObtenerEstadoCuenta() = "Pendiente").Count
            Dim aFavor As Integer = apartamentos.Where(Function(a) a.ObtenerEstadoCuenta() = "Saldo a favor").Count

            Dim totalPendiente As Decimal = apartamentos.Where(Function(a) a.SaldoActual > 0).Sum(Function(a) a.SaldoActual)
            Dim totalAFavor As Decimal = Math.Abs(apartamentos.Where(Function(a) a.SaldoActual < 0).Sum(Function(a) a.SaldoActual))

            lblResumen.Text = $"Total: {totalApartamentos} | Al día: {alDia} | Pendientes: {pendientes} | A favor: {aFavor} | " &
                             $"Total pendiente: {totalPendiente:C} | Total a favor: {totalAFavor:C}"

        Catch ex As Exception
            lblResumen.Text = "Error al calcular resumen"
        End Try
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs)
        AplicarFiltros()
    End Sub

    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs)
        cmbTorre.SelectedIndex = 0
        cmbEstado.SelectedIndex = 0
        txtBuscar.Clear()
        CargarDatos()
    End Sub

    Private Sub BtnExportar_Click(sender As Object, e As EventArgs)
        Try
            Dim saveDialog As New SaveFileDialog With {
                .Filter = "Archivo CSV|*.csv|Archivo Excel|*.xlsx",
                .Title = "Exportar Estados de Cuenta",
                .FileName = $"EstadosCuenta_{DateTime.Now:yyyyMMdd}.csv"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                If saveDialog.FilterIndex = 1 Then
                    ExportarDatos(saveDialog.FileName)
                Else
                    MessageBox.Show("La exportación a Excel está en desarrollo.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al exportar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportarDatos(rutaArchivo As String)
        Try
            Using writer As New System.IO.StreamWriter(rutaArchivo, False, System.Text.Encoding.UTF8)
                ' Escribir encabezados
                Dim encabezados As String = "Torre,Apartamento,Residente,Saldo Actual,Estado,Último Pago,Días Mora,Intereses,Teléfono,Correo"
                writer.WriteLine(encabezados)

                ' Escribir datos
                For Each row As DataGridViewRow In dgvEstados.Rows
                    Dim linea As String = String.Join(",",
                        ObtenerValorCeldaSeguro(row.Cells("Torre").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("Apartamento").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("NombreResidente").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("SaldoActual").Value, "0"),
                        ObtenerValorCeldaSeguro(row.Cells("Estado").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("UltimoPago").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("DiasEnMora").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("TotalIntereses").Value, "0"),
                        ObtenerValorCeldaSeguro(row.Cells("Telefono").Value, ""),
                        ObtenerValorCeldaSeguro(row.Cells("Correo").Value, "")
                    )
                    writer.WriteLine(linea)
                Next
            End Using

            MessageBox.Show($"Datos exportados exitosamente a: {rutaArchivo}", "Exportación Exitosa",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Throw New Exception($"Error al escribir archivo: {ex.Message}")
        End Try
    End Sub

    Private Function ObtenerValorCeldaSeguro(valor As Object, valorPorDefecto As String) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return valorPorDefecto
        End If
        Return valor.ToString()
    End Function

    Private Sub TxtBuscar_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar = ChrW(Keys.Enter) Then
            AplicarFiltros()
        End If
    End Sub

    Private Sub FiltrosChanged(sender As Object, e As EventArgs)
        AplicarFiltros()
    End Sub

    Private Sub AplicarFiltros()
        Try
            Dim filtro As String = ""
            Dim filtros As New List(Of String)

            ' Filtro por torre
            If cmbTorre.SelectedIndex > 0 Then
                filtros.Add($"Torre = '{cmbTorre.SelectedItem.ToString()}'")
            End If

            ' Filtro por estado
            If cmbEstado.SelectedIndex > 0 Then
                filtros.Add($"Estado = '{cmbEstado.SelectedItem.ToString()}'")
            End If

            ' Filtro por búsqueda
            If Not String.IsNullOrWhiteSpace(txtBuscar.Text) Then
                Dim busqueda As String = txtBuscar.Text.Trim()
                filtros.Add($"(Apartamento LIKE '%{busqueda}%' OR NombreResidente LIKE '%{busqueda}%' OR Telefono LIKE '%{busqueda}%' OR Correo LIKE '%{busqueda}%')")
            End If

            ' Combinar filtros
            If filtros.Count > 0 Then
                filtro = String.Join(" AND ", filtros)
            End If

            ' Aplicar filtro
            Dim dv As New DataView(dtEstados)
            dv.RowFilter = filtro
            dgvEstados.DataSource = dv

            ' Actualizar resumen con datos filtrados
            Dim apartamentosFiltrados As New List(Of Apartamento)
            For Each drv As DataRowView In dv
                Dim apto = apartamentos.FirstOrDefault(Function(a) a.NumeroApartamento = drv("Apartamento").ToString())
                If apto IsNot Nothing Then
                    apartamentosFiltrados.Add(apto)
                End If
            Next
            ActualizarResumen(apartamentosFiltrados)

        Catch ex As Exception
            MessageBox.Show($"Error al aplicar filtros: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DgvEstados_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgvEstados.Rows(e.RowIndex)

            ' Colorear según estado
            If row.Cells("Estado").Value IsNot Nothing Then
                Select Case row.Cells("Estado").Value.ToString()
                    Case "Al día"
                        row.Cells("Estado").Style.BackColor = Color.FromArgb(39, 174, 96)
                        row.Cells("Estado").Style.ForeColor = Color.White
                    Case "Pendiente"
                        row.Cells("Estado").Style.BackColor = Color.FromArgb(231, 76, 60)
                        row.Cells("Estado").Style.ForeColor = Color.White
                    Case "Saldo a favor"
                        row.Cells("Estado").Style.BackColor = Color.FromArgb(52, 152, 219)
                        row.Cells("Estado").Style.ForeColor = Color.White
                End Select
            End If

            ' Resaltar días en mora
            If e.ColumnIndex = dgvEstados.Columns("DiasEnMora").Index Then
                If e.Value IsNot Nothing AndAlso IsNumeric(e.Value) Then
                    Dim diasMora As Integer = Convert.ToInt32(e.Value)
                    If diasMora > 90 Then
                        e.CellStyle.BackColor = Color.FromArgb(231, 76, 60)
                        e.CellStyle.ForeColor = Color.White
                    ElseIf diasMora > 60 Then
                        e.CellStyle.BackColor = Color.FromArgb(230, 126, 34)
                        e.CellStyle.ForeColor = Color.White
                    ElseIf diasMora > 30 Then
                        e.CellStyle.BackColor = Color.FromArgb(241, 196, 15)
                    End If
                End If
            End If
        End If
    End Sub

End Class