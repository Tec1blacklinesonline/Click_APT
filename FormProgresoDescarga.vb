' ============================================================================
' FORMPROGESODESCARGA.VB - FORMULARIO DE PROGRESO PARA DESCARGA MASIVA
' ✅ Formulario dedicado para mostrar progreso de descarga de PDFs
' ✅ Compatible con .NET 8 y el sistema existente
' ============================================================================

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Threading

Public Class FormProgresoDescarga
    Inherits Form

    Private WithEvents progressBar As ProgressBar
    Private lblMensaje As Label
    Private lblEstadisticas As Label
    Private btnCancelar As Button
    Private lblTitulo As Label
    Private panelContenido As Panel
    Private cancellationTokenSource As CancellationTokenSource
    Private procesoCompletado As Boolean = False

    Public ReadOnly Property CancellationToken As CancellationToken
        Get
            Return cancellationTokenSource?.Token
        End Get
    End Property

    Public Sub New()
        InitializeComponent()
        ConfigurarFormulario()
        cancellationTokenSource = New CancellationTokenSource()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Configuración básica del formulario
        Me.Text = "Descarga de PDFs en Progreso"
        Me.Size = New Size(500, 280)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(248, 249, 250)
        Me.ShowInTaskbar = False

        Me.ResumeLayout(False)
    End Sub

    Private Sub ConfigurarFormulario()
        ' Panel principal con bordes
        panelContenido = New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(20),
            .BackColor = Color.White
        }

        ' Título del progreso
        lblTitulo = New Label With {
            .Text = "📁 DESCARGA MASIVA DE PDFs",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(155, 89, 182),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Top,
            .Height = 40
        }

        ' Mensaje de estado
        lblMensaje = New Label With {
            .Text = "Preparando descarga...",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .ForeColor = Color.FromArgb(44, 62, 80),
            .TextAlign = ContentAlignment.MiddleLeft,
            .AutoSize = False,
            .Height = 40,
            .Dock = DockStyle.Top
        }

        ' Barra de progreso
        progressBar = New ProgressBar With {
            .Minimum = 0,
            .Maximum = 100,
            .Value = 0,
            .Height = 25,
            .Dock = DockStyle.Top,
            .Style = ProgressBarStyle.Continuous,
            .ForeColor = Color.FromArgb(155, 89, 182)
        }

        ' Estadísticas del progreso
        lblEstadisticas = New Label With {
            .Text = "PDFs procesados: 0/0",
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(127, 140, 141),
            .TextAlign = ContentAlignment.MiddleLeft,
            .AutoSize = False,
            .Height = 30,
            .Dock = DockStyle.Top
        }

        ' Botón cancelar
        btnCancelar = New Button With {
            .Text = "❌ CANCELAR",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(231, 76, 60),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(120, 35),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        btnCancelar.Location = New Point(panelContenido.Width - btnCancelar.Width - 20, panelContenido.Height - btnCancelar.Height - 20)
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click

        ' Agregar controles al panel
        panelContenido.Controls.Add(btnCancelar)
        panelContenido.Controls.Add(lblEstadisticas)
        panelContenido.Controls.Add(progressBar)
        panelContenido.Controls.Add(lblMensaje)
        panelContenido.Controls.Add(lblTitulo)

        ' Agregar panel al formulario
        Me.Controls.Add(panelContenido)

        ' Configurar eventos de cierre
        AddHandler Me.FormClosing, AddressOf FormProgresoDescarga_FormClosing
    End Sub

    ''' <summary>
    ''' Actualiza el progreso de la descarga
    ''' </summary>
    Public Sub ActualizarProgreso(mensaje As String, porcentaje As Integer)
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub() ActualizarProgreso(mensaje, porcentaje))
                Return
            End If

            If Me.IsDisposed OrElse procesoCompletado Then
                Return
            End If

            ' Actualizar mensaje
            lblMensaje.Text = mensaje

            ' Actualizar barra de progreso
            If porcentaje >= 0 AndAlso porcentaje <= 100 Then
                progressBar.Value = porcentaje
            End If

            ' Actualizar estadísticas si el mensaje contiene información numérica
            If mensaje.Contains("/") Then
                Try
                    Dim partes As String() = mensaje.Split(New Char() {"/"c, ":"c, " "c}, StringSplitOptions.RemoveEmptyEntries)
                    For Each parte In partes
                        If parte.Contains("/") Then
                            lblEstadisticas.Text = $"Progreso: {parte} ({porcentaje}%)"
                            Exit For
                        End If
                    Next
                Catch
                    lblEstadisticas.Text = $"Progreso: {porcentaje}%"
                End Try
            Else
                lblEstadisticas.Text = $"Progreso: {porcentaje}%"
            End If

            ' Refresh para asegurar actualización visual
            Me.Refresh()
            Application.DoEvents()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error actualizando progreso: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Marca el proceso como completado
    ''' </summary>
    Public Sub MarcarCompletado(exitoso As Boolean, mensajeFinal As String)
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub() MarcarCompletado(exitoso, mensajeFinal))
                Return
            End If

            If Me.IsDisposed Then
                Return
            End If

            procesoCompletado = True

            ' Actualizar interfaz según el resultado
            If exitoso Then
                progressBar.Value = 100
                lblMensaje.Text = mensajeFinal
                lblMensaje.ForeColor = Color.FromArgb(39, 174, 96) ' Verde
                lblTitulo.Text = "✅ DESCARGA COMPLETADA"
                lblTitulo.ForeColor = Color.FromArgb(39, 174, 96)
                btnCancelar.Text = "✅ CERRAR"
                btnCancelar.BackColor = Color.FromArgb(39, 174, 96)
            Else
                lblMensaje.Text = mensajeFinal
                lblMensaje.ForeColor = Color.FromArgb(231, 76, 60) ' Rojo
                lblTitulo.Text = "❌ ERROR EN DESCARGA"
                lblTitulo.ForeColor = Color.FromArgb(231, 76, 60)
                btnCancelar.Text = "❌ CERRAR"
                btnCancelar.BackColor = Color.FromArgb(231, 76, 60)
            End If

            Me.Refresh()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error marcando completado: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Maneja el clic del botón cancelar
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Try
            If procesoCompletado Then
                ' Si ya terminó, simplemente cerrar
                Me.Close()
            Else
                ' Confirmar cancelación
                Dim resultado As DialogResult = MessageBox.Show(
                    "¿Está seguro de que desea cancelar la descarga?" & vbCrLf & vbCrLf &
                    "⚠️ Los PDFs que ya se hayan generado se conservarán",
                    "Confirmar Cancelación",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2)

                If resultado = DialogResult.Yes Then
                    ' Solicitar cancelación
                    cancellationTokenSource?.Cancel()
                    btnCancelar.Enabled = False
                    btnCancelar.Text = "CANCELANDO..."
                    btnCancelar.BackColor = Color.Gray
                    lblMensaje.Text = "🔄 Cancelando descarga..."
                    lblMensaje.ForeColor = Color.FromArgb(230, 126, 34)
                End If
            End If

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en cancelación: {ex.Message}")
            Me.Close()
        End Try
    End Sub

    ''' <summary>
    ''' Maneja el evento de cierre del formulario
    ''' </summary>
    Private Sub FormProgresoDescarga_FormClosing(sender As Object, e As FormClosingEventArgs)
        Try
            If Not procesoCompletado AndAlso e.CloseReason = CloseReason.UserClosing Then
                ' Prevenir cierre accidental durante el proceso
                e.Cancel = True
                btnCancelar_Click(Nothing, Nothing)
            Else
                ' Limpiar recursos
                cancellationTokenSource?.Dispose()
            End If

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en FormClosing: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Libera recursos del formulario
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                cancellationTokenSource?.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

End Class

