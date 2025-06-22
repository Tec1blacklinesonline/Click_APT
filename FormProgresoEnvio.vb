' ============================================================================
' FORMPROGRESOENVIO SIMPLE - COMPATIBLE CON TU ESTRUCTURA EXISTENTE
' ✅ Formulario ligero para mostrar progreso del envío masivo
' ✅ Reutiliza componentes estándar de Windows Forms
' ============================================================================

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Threading

Public Class FormProgresoEnvio
    Inherits Form

    Private WithEvents progressBar As ProgressBar
    Private lblMensaje As Label
    Private lblPorcentaje As Label
    Private WithEvents btnCancelar As Button
    Private tokenSource As CancellationTokenSource
    Private lblTitulo As Label


    Public ReadOnly Property CancellationToken As CancellationToken
        Get
            Return If(tokenSource?.Token, Threading.CancellationToken.None)
        End Get
    End Property


    Public Sub New()
        tokenSource = New CancellationTokenSource()
        InitializeComponent()
        ConfigurarFormulario()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Configuración básica del formulario
        Me.Text = "Enviando Correos - COOPDIASAM"
        Me.Size = New Size(480, 180)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.White
        Me.ShowInTaskbar = False
        Me.ControlBox = False ' Quitar botones de cerrar para evitar cierre accidental

        Me.ResumeLayout(False)
    End Sub

    Private Sub ConfigurarFormulario()
        ' Título del formulario
        lblTitulo = New Label With {
            .Text = "📧 ENVIANDO RECIBOS DE PAGOS EXTRA",
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.FromArgb(52, 152, 219),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Location = New Point(10, 15),
            .Size = New Size(450, 25),
            .BackColor = Color.Transparent
        }

        ' Mensaje de estado
        lblMensaje = New Label With {
            .Text = "Preparando envío...",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .ForeColor = Color.FromArgb(44, 62, 80),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Location = New Point(20, 50),
            .Size = New Size(430, 20),
            .BackColor = Color.Transparent
        }

        ' Barra de progreso
        progressBar = New ProgressBar With {
            .Minimum = 0,
            .Maximum = 100,
            .Value = 0,
            .Location = New Point(20, 75),
            .Size = New Size(430, 22),
            .Style = ProgressBarStyle.Continuous
        }

        ' Etiqueta de porcentaje
        lblPorcentaje = New Label With {
            .Text = "0%",
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(52, 152, 219),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Location = New Point(20, 105),
            .Size = New Size(430, 18),
            .BackColor = Color.Transparent
        }

        ' Botón cancelar
        btnCancelar = New Button With {
            .Text = "❌ CANCELAR",
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(231, 76, 60),
            .FlatStyle = FlatStyle.Flat,
            .Size = New Size(100, 28),
            .Location = New Point(190, 130),
            .TabIndex = 0
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click

        ' Agregar todos los controles al formulario
        Me.Controls.AddRange({lblTitulo, lblMensaje, progressBar, lblPorcentaje, btnCancelar})
    End Sub

    ''' <summary>
    ''' Actualiza el progreso del envío de forma segura
    ''' </summary>
    Public Sub ActualizarProgreso(mensaje As String, porcentaje As Integer)
        Try
            If Me.InvokeRequired Then
                Me.BeginInvoke(New Action(Sub() ActualizarProgresoInterno(mensaje, porcentaje)))
            Else
                ActualizarProgresoInterno(mensaje, porcentaje)
            End If
        Catch ex As Exception
            ' Error silencioso si el formulario está cerrado
            System.Diagnostics.Debug.WriteLine($"Error actualizando progreso: {ex.Message}")
        End Try
    End Sub

    Private Sub ActualizarProgresoInterno(mensaje As String, porcentaje As Integer)
        Try
            ' Actualizar mensaje (truncar si es muy largo)
            If lblMensaje IsNot Nothing AndAlso Not lblMensaje.IsDisposed Then
                Dim mensajeCorto As String = mensaje
                If mensaje.Length > 60 Then
                    mensajeCorto = mensaje.Substring(0, 57) + "..."
                End If
                lblMensaje.Text = mensajeCorto
            End If

            ' Actualizar barra de progreso
            If progressBar IsNot Nothing AndAlso Not progressBar.IsDisposed Then
                Dim valorSeguro As Integer = Math.Max(0, Math.Min(100, porcentaje))
                progressBar.Value = valorSeguro
            End If

            ' Actualizar etiqueta de porcentaje
            If lblPorcentaje IsNot Nothing AndAlso Not lblPorcentaje.IsDisposed Then
                lblPorcentaje.Text = $"{Math.Max(0, Math.Min(100, porcentaje))}%"

                ' Cambiar color según progreso
                If porcentaje >= 100 Then
                    lblPorcentaje.ForeColor = Color.FromArgb(39, 174, 96) ' Verde
                    btnCancelar.Text = "✅ CERRAR"
                    btnCancelar.BackColor = Color.FromArgb(39, 174, 96)
                    Me.ControlBox = True ' Permitir cerrar cuando termine
                ElseIf porcentaje >= 75 Then
                    lblPorcentaje.ForeColor = Color.FromArgb(243, 156, 18) ' Naranja
                End If
            End If

            ' Actualizar título de ventana
            If porcentaje < 100 Then
                Me.Text = $"Enviando Correos ({porcentaje}%) - COOPDIASAM"
            Else
                Me.Text = "Envío Completado - COOPDIASAM"
            End If

            ' Forzar repintado
            Me.Refresh()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en ActualizarProgresoInterno: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Maneja el evento de cancelación/cierre
    ''' </summary>
    ''' 

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Try
            If btnCancelar.Text.Contains("CERRAR") Then
                ' El proceso ya terminó, cerrar el formulario
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                ' Proceso en curso, confirmar cancelación
                Dim resultado As DialogResult = MessageBox.Show(
                    "¿Está seguro que desea cancelar el envío de correos?" & vbCrLf & vbCrLf &
                    "⚠️ Los correos ya enviados no se pueden detener." & vbCrLf &
                    "📧 Los correos pendientes no se enviarán.",
                    "Confirmar Cancelación",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2)

                If resultado = DialogResult.Yes Then
                    tokenSource?.Cancel()
                    btnCancelar.Enabled = False
                    btnCancelar.Text = "CANCELANDO..."
                    btnCancelar.BackColor = Color.FromArgb(127, 140, 141)
                    lblMensaje.Text = "❌ Cancelando envío... por favor espere"
                    lblMensaje.ForeColor = Color.FromArgb(231, 76, 60)
                    Me.Refresh()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error en cancelación: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Marca el proceso como completado
    ''' </summary>
    Public Sub MarcarCompletado(exitoso As Boolean, mensaje As String)
        Try
            If Me.InvokeRequired Then
                Me.BeginInvoke(New Action(Sub() MarcarCompletadoInterno(exitoso, mensaje)))
            Else
                MarcarCompletadoInterno(exitoso, mensaje)
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error marcando completado: {ex.Message}")
        End Try
    End Sub

    Private Sub MarcarCompletadoInterno(exitoso As Boolean, mensaje As String)
        Try
            If exitoso Then
                ActualizarProgreso("✅ " & mensaje, 100)
                btnCancelar.Text = "✅ CERRAR"
                btnCancelar.BackColor = Color.FromArgb(39, 174, 96)
                lblTitulo.Text = "🎉 ENVÍO COMPLETADO EXITOSAMENTE"
                lblTitulo.ForeColor = Color.FromArgb(39, 174, 96)
            Else
                ActualizarProgreso("❌ " & mensaje, 100)
                btnCancelar.Text = "❌ CERRAR"
                btnCancelar.BackColor = Color.FromArgb(231, 76, 60)
                lblTitulo.Text = "⚠️ ENVÍO COMPLETADO CON ERRORES"
                lblTitulo.ForeColor = Color.FromArgb(231, 76, 60)
            End If

            btnCancelar.Enabled = True
            Me.ControlBox = True

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error en MarcarCompletadoInterno: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Limpieza al cerrar el formulario
    ''' </summary>
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            tokenSource?.Cancel()
            tokenSource?.Dispose()
        Catch
            ' Error silencioso
        End Try
        MyBase.OnFormClosed(e)
    End Sub

    ''' <summary>
    ''' Prevenir cierre accidental durante el envío
    ''' </summary>
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        ' Solo permitir cierre si el proceso terminó o el usuario canceló explícitamente
        If btnCancelar.Text.Contains("CANCELAR") AndAlso btnCancelar.Enabled Then
            ' Proceso en curso, confirmar cierre
            Dim resultado As DialogResult = MessageBox.Show(
                "¿Desea cancelar el envío de correos y cerrar esta ventana?",
                "Confirmar Cierre",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)

            If resultado = DialogResult.No Then
                e.Cancel = True
                Return
            Else
                tokenSource?.Cancel()
            End If
        End If

        MyBase.OnFormClosing(e)
    End Sub

    ''' <summary>
    ''' Configurar formulario para un número específico de elementos
    ''' </summary>
    Public Sub ConfigurarParaTotal(totalItems As Integer, descripcion As String)
        Try
            If lblTitulo IsNot Nothing Then
                lblTitulo.Text = $"📧 {descripcion.ToUpper()}"
            End If

            If lblMensaje IsNot Nothing Then
                lblMensaje.Text = $"Preparando envío de {totalItems} elementos..."
            End If

            Me.Text = $"Enviando {totalItems} Correos - COOPDIASAM"

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error configurando formulario: {ex.Message}")
        End Try
    End Sub

End Class