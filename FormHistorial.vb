' ============================================================================
' FORMULARIO DE HISTORIAL
' Muestra historial de pagos y cambios del sistema
' ============================================================================

Imports System.Data.SQLite
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class FormHistorial
    Private dtPagos As DataTable
    Private dtCambios As DataTable
    Private apartamentoSeleccionado As Apartamento

    ' Controles del formulario
    Private tabControl As TabControl
    Private tabPagos As TabPage
    Private tabCambios As TabPage
    Private WithEvents dgvPagos As DataGridView
    Private WithEvents dgvCambios As DataGridView
    Private WithEvents cmbApartamento As ComboBox
    Private WithEvents dtpDesde As DateTimePicker
    Private WithEvents dtpHasta As DateTimePicker
    Private WithEvents btnBuscarPagos As Button
    Private WithEvents btnExportarPagos As Button
    Private lblTotalesPagos As Label
    Private WithEvents cmbTabla As ComboBox
    Private WithEvents cmbTipoCambio As ComboBox
    Private WithEvents txtUsuarioCambio As TextBox
    Private WithEvents btnBuscarCambios As Button
    Private WithEvents btnExportarCambios As Button

    Private Sub FormHistorial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarDatos()
    End Sub

    Private Sub ConfigurarFormulario()
        ' Form
        Me.Text = "Historial - COOPDIASAM"
        Me.Size = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Panel superior
        Dim panelSuperior As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(52, 73, 94)
        }

        Dim lblTitulo As New Label With {
            .Text = "HISTORIAL DEL SISTEMA",
            .Font = New Font("Arial", 20, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 25)
        }
        panelSuperior.Controls.Add(lblTitulo)

        ' TabControl
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill,
            .Font = New Font("Arial", 10)
        }

        ' Tab de Pagos
        tabPagos = New TabPage("Historial de Pagos")
        ConfigurarTabPagos()
        tabControl.TabPages.Add(tabPagos)

        ' Tab de Cambios
        tabCambios = New TabPage("Historial de Cambios")
        ConfigurarTabCambios()
        tabControl.TabPages.Add(tabCambios)

        ' Agregar controles al formulario
        Me.Controls.Add(tabControl)
        Me.Controls.Add(panelSuperior)

        ' Configurar eventos
        ConfigurarEventos()
    End Sub

    Private Sub ConfigurarTabPagos()
        ' Panel de filtros
        Dim panelFiltros As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 100,
            .BackColor = Color.FromArgb(236, 240, 241),
            .Padding = New Padding(10)
        }

        ' ComboBox Apartamento
        Dim lblApartamento As New Label With {
            .Text = "Apartamento:",
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblApartamento)

        cmbApartamento = New ComboBox With {
            .Location = New Point(110, 17),
            .Size = New Size(200, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        panelFiltros.Controls.Add(cmbApartamento)

        ' DateTimePickers
        Dim lblDesde As New Label With {
            .Text = "Desde:",
            .Location = New Point(340, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblDesde)

        dtpDesde = New DateTimePicker With {
            .Location = New Point(390, 17),
            .Size = New Size(150, 25),
            .Format = DateTimePickerFormat.Short
        }
        dtpDesde.Value = DateTime.Now.AddMonths(-6)
        panelFiltros.Controls.Add(dtpDesde)

        Dim lblHasta As New Label With {
            .Text = "Hasta:",
            .Location = New Point(560, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblHasta)

        dtpHasta = New DateTimePicker With {
            .Location = New Point(605, 17),
            .Size = New Size(150, 25),
            .Format = DateTimePickerFormat.Short
        }
        panelFiltros.Controls.Add(dtpHasta)

        ' Botones
        btnBuscarPagos = New Button With {
            .Text = "🔍 Buscar",
            .Location = New Point(780, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnBuscarPagos.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnBuscarPagos)

        btnExportarPagos = New Button With {
            .Text = "📊 Exportar",
            .Location = New Point(890, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnExportarPagos.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnExportarPagos)

        ' Totales
        lblTotalesPagos = New Label With {
            .Location = New Point(20, 60),
            .AutoSize = True,
            .Font = New Font("Arial", 10, FontStyle.Bold)
        }
        panelFiltros.Controls.Add(lblTotalesPagos)

        ' DataGridView
        dgvPagos = New DataGridView With {
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
        ConfigurarDataGridView(dgvPagos)

        tabPagos.Controls.Add(dgvPagos)
        tabPagos.Controls.Add(panelFiltros)
    End Sub

    Private Sub ConfigurarTabCambios()
        ' Panel de filtros
        Dim panelFiltros As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 100,
            .BackColor = Color.FromArgb(236, 240, 241),
            .Padding = New Padding(10)
        }

        ' ComboBox Tabla
        Dim lblTabla As New Label With {
            .Text = "Tabla:",
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblTabla)

        cmbTabla = New ComboBox With {
            .Location = New Point(70, 17),
            .Size = New Size(200, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cmbTabla.Items.AddRange({"Todas", "Apartamentos", "Pagos", "Usuarios", "Cuotas", "Parámetros"})
        cmbTabla.SelectedIndex = 0
        panelFiltros.Controls.Add(cmbTabla)

        ' ComboBox Tipo Cambio
        Dim lblTipoCambio As New Label With {
            .Text = "Tipo:",
            .Location = New Point(300, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblTipoCambio)

        cmbTipoCambio = New ComboBox With {
            .Location = New Point(340, 17),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cmbTipoCambio.Items.AddRange({"Todos", "INSERT", "UPDATE", "DELETE"})
        cmbTipoCambio.SelectedIndex = 0
        panelFiltros.Controls.Add(cmbTipoCambio)

        ' TextBox Usuario
        Dim lblUsuario As New Label With {
            .Text = "Usuario:",
            .Location = New Point(520, 20),
            .AutoSize = True
        }
        panelFiltros.Controls.Add(lblUsuario)

        txtUsuarioCambio = New TextBox With {
            .Location = New Point(580, 17),
            .Size = New Size(150, 25)
        }
        panelFiltros.Controls.Add(txtUsuarioCambio)

        ' Botones
        btnBuscarCambios = New Button With {
            .Text = "🔍 Buscar",
            .Location = New Point(750, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnBuscarCambios.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnBuscarCambios)

        btnExportarCambios = New Button With {
            .Text = "📊 Exportar",
            .Location = New Point(860, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Cursor = Cursors.Hand
        }
        btnExportarCambios.FlatAppearance.BorderSize = 0
        panelFiltros.Controls.Add(btnExportarCambios)

        ' DataGridView
        dgvCambios = New DataGridView With {
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
        ConfigurarDataGridView(dgvCambios)

        tabCambios.Controls.Add(dgvCambios)
        tabCambios.Controls.Add(panelFiltros)
    End Sub

    Private Sub ConfigurarDataGridView(dgv As DataGridView)
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 73, 94)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
        dgv.ColumnHeadersHeight = 35
        dgv.DefaultCellStyle.Font = New Font("Arial", 9)
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
        dgv.RowTemplate.Height = 30
    End Sub

    Private Sub ConfigurarEventos()
        ' Eventos
        AddHandler btnBuscarPagos.Click, AddressOf BtnBuscarPagos_Click
        AddHandler btnExportarPagos.Click, AddressOf BtnExportarPagos_Click
        AddHandler btnBuscarCambios.Click, AddressOf BtnBuscarCambios_Click
        AddHandler btnExportarCambios.Click, AddressOf BtnExportarCambios_Click
        AddHandler cmbApartamento.SelectedIndexChanged, AddressOf CmbApartamento_SelectedIndexChanged
        AddHandler dgvPagos.CellFormatting, AddressOf DgvPagos_CellFormatting
        AddHandler dgvCambios.CellFormatting, AddressOf DgvCambios_CellFormatting
    End Sub

    Private Sub CargarDatos()
        Try
            ' Cargar apartamentos
            CargarApartamentos()

            ' Cargar historial inicial
            If cmbApartamento.Items.Count > 0 Then
                cmbApartamento.SelectedIndex = 0
            End If

            ' Cargar cambios recientes
            CargarHistorialCambios()

        Catch ex As Exception
            MessageBox.Show($"Error al cargar datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarApartamentos()
        Try
            cmbApartamento.Items.Clear()
            cmbApartamento.Items.Add("TODOS")

            Dim apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()
            If apartamentos IsNot Nothing Then
                For Each apto In apartamentos.OrderBy(Function(a) a.Torre).ThenBy(Function(a) a.NumeroApartamento)
                    cmbApartamento.Items.Add($"Torre {apto.Torre} - {apto.NumeroApartamento}")
                Next
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al cargar apartamentos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarHistorialPagos(Optional idApartamento As Integer = 0)
        Try
            dtPagos = New DataTable()
            dtPagos.Columns.Add("Fecha", GetType(DateTime))
            dtPagos.Columns.Add("NumeroRecibo", GetType(String))
            dtPagos.Columns.Add("Apartamento", GetType(String))
            dtPagos.Columns.Add("Residente", GetType(String))
            dtPagos.Columns.Add("SaldoAnterior", GetType(Decimal))
            dtPagos.Columns.Add("PagoAdministracion", GetType(Decimal))
            dtPagos.Columns.Add("PagoIntereses", GetType(Decimal))
            dtPagos.Columns.Add("TotalPagado", GetType(Decimal))
            dtPagos.Columns.Add("SaldoActual", GetType(Decimal))
            dtPagos.Columns.Add("EstadoPago", GetType(String))
            dtPagos.Columns.Add("Observaciones", GetType(String))

            Dim pagos As List(Of PagoModel)
            If idApartamento > 0 Then
                pagos = PagosDAL.ObtenerHistorialPagos(idApartamento)
            Else
                pagos = PagosDAL.ObtenerHistorialPagos(0) ' Todos los pagos
            End If

            If pagos IsNot Nothing Then
                ' Filtrar por fechas
                pagos = pagos.Where(Function(p) p.FechaPago >= dtpDesde.Value.Date AndAlso p.FechaPago <= dtpHasta.Value.Date.AddDays(1)).ToList()

                Dim totalPagado As Decimal = 0
                Dim totalAdministracion As Decimal = 0
                Dim totalIntereses As Decimal = 0

                For Each pago In pagos
                    Dim apto = ApartamentoDAL.ObtenerApartamentoPorId(pago.IdApartamento)
                    If apto IsNot Nothing Then
                        Dim row As DataRow = dtPagos.NewRow()
                        row("Fecha") = pago.FechaPago
                        row("NumeroRecibo") = pago.NumeroRecibo
                        row("Apartamento") = $"Torre {apto.Torre} - {apto.NumeroApartamento}"
                        row("Residente") = apto.NombreResidente
                        row("SaldoAnterior") = pago.SaldoAnterior
                        row("PagoAdministracion") = pago.PagoAdministracion
                        row("PagoIntereses") = pago.PagoIntereses
                        row("TotalPagado") = pago.TotalPagado
                        row("SaldoActual") = pago.SaldoActual
                        row("EstadoPago") = pago.EstadoPago
                        row("Observaciones") = pago.Observaciones

                        dtPagos.Rows.Add(row)

                        totalPagado += pago.TotalPagado
                        totalAdministracion += pago.PagoAdministracion
                        totalIntereses += pago.PagoIntereses
                    End If
                Next

                ' Actualizar totales
                lblTotalesPagos.Text = $"Total registros: {pagos.Count} | Total pagado: {totalPagado:C} | " &
                                      $"Administración: {totalAdministracion:C} | Intereses: {totalIntereses:C}"
            End If

            dgvPagos.DataSource = dtPagos
            FormatearColumnasPagos()

        Catch ex As Exception
            MessageBox.Show($"Error al cargar historial de pagos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarHistorialCambios()
        Try
            dtCambios = New DataTable()
            dtCambios.Columns.Add("Fecha", GetType(DateTime))
            dtCambios.Columns.Add("TablaAfectada", GetType(String))
            dtCambios.Columns.Add("TipoCambio", GetType(String))
            dtCambios.Columns.Add("IdRegistro", GetType(String))
            dtCambios.Columns.Add("UsuarioResponsable", GetType(String))
            dtCambios.Columns.Add("DetalleCambio", GetType(String))

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT * FROM historico_cambios WHERE 1=1"

                ' Aplicar filtros
                If cmbTabla.SelectedIndex > 0 Then
                    consulta &= $" AND tabla_afectada = '{cmbTabla.SelectedItem.ToString()}'"
                End If

                If cmbTipoCambio.SelectedIndex > 0 Then
                    consulta &= $" AND tipo_cambio = '{cmbTipoCambio.SelectedItem.ToString()}'"
                End If

                If Not String.IsNullOrWhiteSpace(txtUsuarioCambio.Text) Then
                    consulta &= $" AND usuario_responsable LIKE '%{txtUsuarioCambio.Text}%'"
                End If

                consulta &= " ORDER BY fecha_cambio DESC LIMIT 500"

                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim row As DataRow = dtCambios.NewRow()
                            row("Fecha") = Convert.ToDateTime(reader("fecha_cambio"))
                            row("TablaAfectada") = reader("tabla_afectada").ToString()
                            row("TipoCambio") = reader("tipo_cambio").ToString()
                            row("IdRegistro") = reader("id_registro_afectado").ToString()
                            row("UsuarioResponsable") = reader("usuario_responsable").ToString()
                            row("DetalleCambio") = reader("detalle_cambio").ToString()

                            dtCambios.Rows.Add(row)
                        End While
                    End Using
                End Using
            End Using

            dgvCambios.DataSource = dtCambios
            FormatearColumnasCambios()

        Catch ex As Exception
            ' Si la tabla no existe, no mostrar error
            If Not ex.Message.Contains("no such table") Then
                MessageBox.Show($"Error al cargar historial de cambios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    Private Sub FormatearColumnasPagos()
        If dgvPagos.Columns.Count > 0 Then
            dgvPagos.Columns("Fecha").DefaultCellStyle.Format = "dd/MM/yyyy"
            dgvPagos.Columns("Fecha").Width = 100
            dgvPagos.Columns("NumeroRecibo").HeaderText = "N° Recibo"
            dgvPagos.Columns("NumeroRecibo").Width = 100
            dgvPagos.Columns("Apartamento").Width = 120
            dgvPagos.Columns("Residente").Width = 200
            dgvPagos.Columns("SaldoAnterior").DefaultCellStyle.Format = "C0"
            dgvPagos.Columns("SaldoAnterior").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvPagos.Columns("SaldoAnterior").HeaderText = "Saldo Anterior"
            dgvPagos.Columns("PagoAdministracion").DefaultCellStyle.Format = "C0"
            dgvPagos.Columns("PagoAdministracion").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvPagos.Columns("PagoAdministracion").HeaderText = "Administración"
            dgvPagos.Columns("PagoIntereses").DefaultCellStyle.Format = "C0"
            dgvPagos.Columns("PagoIntereses").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvPagos.Columns("PagoIntereses").HeaderText = "Intereses"
            dgvPagos.Columns("TotalPagado").DefaultCellStyle.Format = "C0"
            dgvPagos.Columns("TotalPagado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvPagos.Columns("TotalPagado").HeaderText = "Total Pagado"
            dgvPagos.Columns("SaldoActual").DefaultCellStyle.Format = "C0"
            dgvPagos.Columns("SaldoActual").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvPagos.Columns("SaldoActual").HeaderText = "Saldo Actual"
            dgvPagos.Columns("EstadoPago").HeaderText = "Estado"
            dgvPagos.Columns("EstadoPago").Width = 100
            dgvPagos.Columns("Observaciones").Width = 200
        End If
    End Sub

    Private Sub FormatearColumnasCambios()
        If dgvCambios.Columns.Count > 0 Then
            dgvCambios.Columns("Fecha").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm"
            dgvCambios.Columns("Fecha").Width = 150
            dgvCambios.Columns("TablaAfectada").HeaderText = "Tabla"
            dgvCambios.Columns("TablaAfectada").Width = 120
            dgvCambios.Columns("TipoCambio").HeaderText = "Tipo"
            dgvCambios.Columns("TipoCambio").Width = 80
            dgvCambios.Columns("IdRegistro").HeaderText = "ID Registro"
            dgvCambios.Columns("IdRegistro").Width = 100
            dgvCambios.Columns("UsuarioResponsable").HeaderText = "Usuario"
            dgvCambios.Columns("UsuarioResponsable").Width = 150
            dgvCambios.Columns("DetalleCambio").HeaderText = "Detalle"
            dgvCambios.Columns("DetalleCambio").Width = 400
        End If
    End Sub

    Private Sub CmbApartamento_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbApartamento.SelectedIndex >= 0 Then
            If cmbApartamento.SelectedIndex = 0 Then
                ' Todos los apartamentos
                CargarHistorialPagos(0)
            Else
                ' Apartamento específico
                Dim textoSeleccionado As String = cmbApartamento.SelectedItem.ToString()
                Dim partes() As String = textoSeleccionado.Split("-"c)
                If partes.Length = 2 Then
                    Dim torreTexto As String = partes(0).Trim()
                    Dim numero As String = partes(1).Trim()

                    ' Extraer el número de torre
                    Dim torre As Integer = 0
                    If torreTexto.StartsWith("Torre ") Then
                        Integer.TryParse(torreTexto.Substring(6), torre)
                    End If

                    Dim apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()
                    If apartamentos IsNot Nothing Then
                        Dim apto = apartamentos.FirstOrDefault(
                            Function(a) a.Torre = torre AndAlso a.NumeroApartamento = numero)

                        If apto IsNot Nothing Then
                            CargarHistorialPagos(apto.IdApartamento)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub BtnBuscarPagos_Click(sender As Object, e As EventArgs)
        CmbApartamento_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub BtnBuscarCambios_Click(sender As Object, e As EventArgs)
        CargarHistorialCambios()
    End Sub

    Private Sub BtnExportarPagos_Click(sender As Object, e As EventArgs)
        Try
            Dim saveDialog As New SaveFileDialog With {
                .Filter = "Archivo CSV|*.csv",
                .Title = "Exportar Historial de Pagos",
                .FileName = $"HistorialPagos_{DateTime.Now:yyyyMMdd}.csv"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                ExportarHistorialPagos(saveDialog.FileName)
                MessageBox.Show("Historial exportado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al exportar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnExportarCambios_Click(sender As Object, e As EventArgs)
        Try
            Dim saveDialog As New SaveFileDialog With {
                .Filter = "Archivo CSV|*.csv",
                .Title = "Exportar Historial de Cambios",
                .FileName = $"HistorialCambios_{DateTime.Now:yyyyMMdd}.csv"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                ExportarHistorialCambios(saveDialog.FileName)
                MessageBox.Show("Historial exportado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al exportar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportarHistorialPagos(rutaArchivo As String)
        Using writer As New System.IO.StreamWriter(rutaArchivo, False, System.Text.Encoding.UTF8)
            writer.WriteLine("Fecha,Número Recibo,Apartamento,Residente,Saldo Anterior,Pago Administración,Pago Intereses,Total Pagado,Saldo Actual,Estado,Observaciones")

            For Each row As DataGridViewRow In dgvPagos.Rows
                Dim fecha As String = ""
                If row.Cells("Fecha").Value IsNot Nothing Then
                    fecha = Convert.ToDateTime(row.Cells("Fecha").Value).ToString("dd/MM/yyyy")
                End If

                Dim linea As String = String.Join(",",
                    fecha,
                    ObtenerValorCeldaSeguro(row.Cells("NumeroRecibo").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("Apartamento").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("Residente").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("SaldoAnterior").Value, "0"),
                    ObtenerValorCeldaSeguro(row.Cells("PagoAdministracion").Value, "0"),
                    ObtenerValorCeldaSeguro(row.Cells("PagoIntereses").Value, "0"),
                    ObtenerValorCeldaSeguro(row.Cells("TotalPagado").Value, "0"),
                    ObtenerValorCeldaSeguro(row.Cells("SaldoActual").Value, "0"),
                    ObtenerValorCeldaSeguro(row.Cells("EstadoPago").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("Observaciones").Value, "")
                )
                writer.WriteLine(linea)
            Next
        End Using
    End Sub

    Private Sub ExportarHistorialCambios(rutaArchivo As String)
        Using writer As New System.IO.StreamWriter(rutaArchivo, False, System.Text.Encoding.UTF8)
            writer.WriteLine("Fecha,Tabla Afectada,Tipo Cambio,ID Registro,Usuario,Detalle")

            For Each row As DataGridViewRow In dgvCambios.Rows
                Dim fecha As String = ""
                If row.Cells("Fecha").Value IsNot Nothing Then
                    fecha = Convert.ToDateTime(row.Cells("Fecha").Value).ToString("dd/MM/yyyy HH:mm")
                End If

                Dim linea As String = String.Join(",",
                    fecha,
                    ObtenerValorCeldaSeguro(row.Cells("TablaAfectada").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("TipoCambio").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("IdRegistro").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("UsuarioResponsable").Value, ""),
                    ObtenerValorCeldaSeguro(row.Cells("DetalleCambio").Value, "")
                )
                writer.WriteLine(linea)
            Next
        End Using
    End Sub

    Private Function ObtenerValorCeldaSeguro(valor As Object, valorPorDefecto As String) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return valorPorDefecto
        End If
        Return valor.ToString()
    End Function

    Private Sub DgvPagos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex >= 0 Then
            ' Colorear estado de pago
            If e.ColumnIndex = dgvPagos.Columns("EstadoPago").Index Then
                If e.Value IsNot Nothing Then
                    Select Case e.Value.ToString().ToUpper()
                        Case "PAGADO", "REGISTRADO"
                            e.CellStyle.BackColor = Color.FromArgb(39, 174, 96)
                            e.CellStyle.ForeColor = Color.White
                        Case "ANULADO"
                            e.CellStyle.BackColor = Color.FromArgb(231, 76, 60)
                            e.CellStyle.ForeColor = Color.White
                    End Select
                End If
            End If
        End If
    End Sub

    Private Sub DgvCambios_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex >= 0 Then
            ' Colorear tipo de cambio
            If e.ColumnIndex = dgvCambios.Columns("TipoCambio").Index Then
                If e.Value IsNot Nothing Then
                    Select Case e.Value.ToString().ToUpper()
                        Case "INSERT"
                            e.CellStyle.BackColor = Color.FromArgb(39, 174, 96)
                            e.CellStyle.ForeColor = Color.White
                        Case "UPDATE"
                            e.CellStyle.BackColor = Color.FromArgb(241, 196, 15)
                            e.CellStyle.ForeColor = Color.Black
                        Case "DELETE"
                            e.CellStyle.BackColor = Color.FromArgb(231, 76, 60)
                            e.CellStyle.ForeColor = Color.White
                    End Select
                End If
            End If
        End If
    End Sub

End Class