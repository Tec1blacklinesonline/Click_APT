Imports System.Windows.Forms

Public Class FormApartamentosTorre
    Private numeroTorre As Integer
    Private apartamentos As List(Of Apartamento)
    Private dgvApartamentos As DataGridView
    Private lblTitulo As Label
    Private btnVolver As Button
    Private lblResumen As Label

    ' Constructor que recibe el número de torre
    Public Sub New(torre As Integer)
        InitializeComponent()
        Me.numeroTorre = torre
    End Sub

    Private Sub ConfigurarFormulario()
        Me.SuspendLayout()

        ' Configuración del formulario
        Me.Text = $"Apartamentos Torre {numeroTorre}"
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' Panel superior
        Dim panelSuperior As New Panel With {
            .Size = New Size(Me.ClientSize.Width, 80),
            .Location = New Point(0, 0),
            .BackColor = Color.FromArgb(41, 128, 185),
            .Dock = DockStyle.Top
        }

        ' Título
        lblTitulo = New Label With {
            .Text = $"APARTAMENTOS - TORRE {numeroTorre}",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelSuperior.Controls.Add(lblTitulo)

        ' Botón volver
        btnVolver = New Button With {
            .Text = "← Volver",
            .Size = New Size(100, 35),
            .Location = New Point(panelSuperior.Width - 120, 22),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        btnVolver.FlatAppearance.BorderSize = 0
        AddHandler btnVolver.Click, AddressOf btnVolver_Click
        panelSuperior.Controls.Add(btnVolver)

        ' Label de resumen
        lblResumen = New Label With {
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 50)
        }
        panelSuperior.Controls.Add(lblResumen)

        Me.Controls.Add(panelSuperior)

        ' Panel para el DataGridView
        Dim panelGrid As New Panel With {
            .Location = New Point(20, 100),
            .Size = New Size(Me.ClientSize.Width - 40, Me.ClientSize.Height - 140),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .BackColor = Color.White
        }

        ' DataGridView para mostrar apartamentos
        dgvApartamentos = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .ReadOnly = False,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .RowHeadersVisible = False,
            .BackgroundColor = Color.White,
            .GridColor = Color.LightGray,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .Font = New Font("Segoe UI", 9),
                .SelectionBackColor = Color.FromArgb(52, 152, 219),
                .SelectionForeColor = Color.White
            },
            .ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
                .Font = New Font("Segoe UI", 10, FontStyle.Bold),
                .BackColor = Color.FromArgb(52, 73, 94),
                .ForeColor = Color.White,
                .Alignment = DataGridViewContentAlignment.MiddleCenter
            },
            .EnableHeadersVisualStyles = False
        }

        ' Configurar columnas
        ConfigurarColumnas()

        panelGrid.Controls.Add(dgvApartamentos)
        Me.Controls.Add(panelGrid)

        Me.ResumeLayout(False)
    End Sub

    Private Sub ConfigurarColumnas()
        dgvApartamentos.Columns.Clear()

        ' Columna Apartamento
        dgvApartamentos.Columns.Add("Apartamento", "Apartamento")
        dgvApartamentos.Columns("Apartamento").Width = 80
        dgvApartamentos.Columns("Apartamento").ReadOnly = True

        ' Columna Nombre Residente
        dgvApartamentos.Columns.Add("NombreResidente", "Nombre del Residente")
        dgvApartamentos.Columns("NombreResidente").Width = 200
        dgvApartamentos.Columns("NombreResidente").ReadOnly = False

        ' Columna Teléfono
        dgvApartamentos.Columns.Add("Telefono", "Teléfono")
        dgvApartamentos.Columns("Telefono").Width = 120
        dgvApartamentos.Columns("Telefono").ReadOnly = False

        ' Columna Correo
        dgvApartamentos.Columns.Add("Correo", "Correo Electrónico")
        dgvApartamentos.Columns("Correo").Width = 200
        dgvApartamentos.Columns("Correo").ReadOnly = False

        ' Columna Estado
        dgvApartamentos.Columns.Add("Estado", "Estado")
        dgvApartamentos.Columns("Estado").Width = 100
        dgvApartamentos.Columns("Estado").ReadOnly = True

        ' Columna Saldo
        dgvApartamentos.Columns.Add("Saldo", "Saldo")
        dgvApartamentos.Columns("Saldo").Width = 120
        dgvApartamentos.Columns("Saldo").ReadOnly = True
        dgvApartamentos.Columns("Saldo").DefaultCellStyle.Format = "C"
        dgvApartamentos.Columns("Saldo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        ' Columna Último Pago
        dgvApartamentos.Columns.Add("UltimoPago", "Último Pago")
        dgvApartamentos.Columns("UltimoPago").Width = 100
        dgvApartamentos.Columns("UltimoPago").ReadOnly = True
        dgvApartamentos.Columns("UltimoPago").DefaultCellStyle.Format = "dd/MM/yyyy"

        ' Columna oculta para ID
        dgvApartamentos.Columns.Add("IdApartamento", "ID")
        dgvApartamentos.Columns("IdApartamento").Visible = False
    End Sub

    Private Sub FormApartamentosTorre_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        ConfigurarControles()
        CargarApartamentos()
        MostrarResumen()

        ' Agregar eventos
        AddHandler dgvApartamentos.CellValueChanged, AddressOf dgvApartamentos_CellValueChanged
        AddHandler dgvApartamentos.CellFormatting, AddressOf dgvApartamentos_CellFormatting
    End Sub

    Private Sub ConfigurarControles()

    Private Sub CargarApartamentos()
        Try
            ' Obtener apartamentos de la base de datos
            apartamentos = ApartamentoDAL.ObtenerApartamentosPorTorre(numeroTorre)

            ' Limpiar el DataGridView
            dgvApartamentos.Rows.Clear()

            ' Llenar el DataGridView
            For Each apartamento In apartamentos
                Dim fila As Integer = dgvApartamentos.Rows.Add()
                Dim row As DataGridViewRow = dgvApartamentos.Rows(fila)

                row.Cells("Apartamento").Value = apartamento.ObtenerCodigoApartamento()
                row.Cells("NombreResidente").Value = apartamento.NombreResidente
                row.Cells("Telefono").Value = apartamento.Telefono
                row.Cells("Correo").Value = apartamento.Correo
                row.Cells("Estado").Value = apartamento.ObtenerEstadoCuenta()
                row.Cells("Saldo").Value = apartamento.SaldoActual
                row.Cells("UltimoPago").Value = If(apartamento.TieneUltimoPago,
                                                  CObj(apartamento.UltimoPago), Nothing)
                row.Cells("IdApartamento").Value = apartamento.IdApartamento

                ' Aplicar color según el estado
                Dim colorEstado As Color = apartamento.ObtenerColorEstado()
                row.Cells("Estado").Style.ForeColor = colorEstado
                row.Cells("Saldo").Style.ForeColor = colorEstado
            Next

        Catch ex As Exception
            MessageBox.Show($"Error al cargar apartamentos: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MostrarResumen()
        Try
            Dim resumen = ApartamentoDAL.ObtenerResumenTorre(numeroTorre)

            Dim alDia = CInt(resumen("apartamentos_al_dia"))
            Dim pendientes = CInt(resumen("apartamentos_pendientes"))
            Dim totalPendiente = CDec(resumen("total_pendiente"))
            Dim totalAFavor = CDec(resumen("total_a_favor"))

            lblResumen.Text = $"Al día: {alDia} | Pendientes: {pendientes} | " &
                             $"Total pendiente: {totalPendiente:C} | Total a favor: {totalAFavor:C}"

        Catch ex As Exception
            lblResumen.Text = "Error al cargar resumen"
        End Try
    End Sub

    Private Sub dgvApartamentos_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Try
                ' Obtener el apartamento modificado
                Dim row As DataGridViewRow = dgvApartamentos.Rows(e.RowIndex)
                Dim idApartamento As Integer = CInt(row.Cells("IdApartamento").Value)

                ' Buscar el apartamento en la lista
                Dim apartamento As Apartamento = Nothing
                For Each apt In apartamentos
                    If apt.IdApartamento = idApartamento Then
                        apartamento = apt
                        Exit For
                    End If
                Next

                If apartamento IsNot Nothing Then
                    ' Actualizar los valores
                    apartamento.NombreResidente = If(row.Cells("NombreResidente").Value IsNot Nothing,
                                                   row.Cells("NombreResidente").Value.ToString(), "")
                    apartamento.Telefono = If(row.Cells("Telefono").Value IsNot Nothing,
                                            row.Cells("Telefono").Value.ToString(), "")
                    apartamento.Correo = If(row.Cells("Correo").Value IsNot Nothing,
                                          row.Cells("Correo").Value.ToString(), "")

                    ' Guardar en la base de datos
                    If ApartamentoDAL.ActualizarPropietario(apartamento) Then
                        MessageBox.Show("Información actualizada correctamente", "Éxito",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Error al actualizar la información", "Error",
                                      MessageBoxButtons.OK, MessageBoxIcon.Error)
                        CargarApartamentos() ' Recargar datos originales
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show($"Error al guardar cambios: {ex.Message}", "Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                CargarApartamentos() ' Recargar datos originales
            End Try
        End If
    End Sub

    Private Sub dgvApartamentos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgvApartamentos.Rows(e.RowIndex)
            Dim estado As String = If(row.Cells("Estado").Value IsNot Nothing,
                                    row.Cells("Estado").Value.ToString(), "")

            ' Aplicar formato según el estado
            Select Case estado
                Case "Al día"
                    If e.ColumnIndex = dgvApartamentos.Columns("Estado").Index Or
                       e.ColumnIndex = dgvApartamentos.Columns("Saldo").Index Then
                        e.CellStyle.ForeColor = Color.Black
                        e.CellStyle.Font = New Font(e.CellStyle.Font, FontStyle.Bold)
                    End If
                Case "Saldo a favor"
                    If e.ColumnIndex = dgvApartamentos.Columns("Estado").Index Or
                       e.ColumnIndex = dgvApartamentos.Columns("Saldo").Index Then
                        e.CellStyle.ForeColor = Color.Green
                        e.CellStyle.Font = New Font(e.CellStyle.Font, FontStyle.Bold)
                    End If
                Case "Pendiente"
                    If e.ColumnIndex = dgvApartamentos.Columns("Estado").Index Or
                       e.ColumnIndex = dgvApartamentos.Columns("Saldo").Index Then
                        e.CellStyle.ForeColor = Color.Red
                        e.CellStyle.Font = New Font(e.CellStyle.Font, FontStyle.Bold)
                    End If
            End Select
        End If
    End Sub

    Private Sub btnVolver_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

End Class