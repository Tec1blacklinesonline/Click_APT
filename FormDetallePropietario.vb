' FormDetallePropietario.vb
Imports System.Windows.Forms
Imports System.Drawing

Public Class FormDetallePropietario
    Inherits Form

    Private apartamentoActual As Apartamento
    Private esModoEdicion As Boolean = False

    ' Controles del formulario
    Private lblTitulo As Label
    Private txtTorre As TextBox
    Private txtPiso As TextBox
    Private txtNumeroApartamento As TextBox
    Private txtNombreResidente As TextBox
    Private txtTelefono As TextBox
    Private txtCorreo As TextBox
    Private txtMatriculaInmobiliaria As TextBox
    Private btnGuardar As Button
    Private btnCancelar As Button

    ' Constructor para agregar nuevo
    Public Sub New()
        MyBase.New()
        esModoEdicion = False
        apartamentoActual = New Apartamento()
        InitializeComponent()
        ConfigurarModoAgregar()
    End Sub

    ' Constructor para editar existente
    Public Sub New(apartamento As Apartamento)
        MyBase.New()

        If apartamento Is Nothing Then
            Throw New ArgumentNullException("apartamento", "El apartamento no puede ser nulo")
        End If

        esModoEdicion = True
        apartamentoActual = apartamento
        InitializeComponent()
        ConfigurarModoEdicion()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = If(esModoEdicion, "Editar Propietario", "Agregar Nuevo Propietario")
        Me.Size = New Size(480, 450)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)

        CrearControles()
    End Sub

    Private Sub CrearControles()
        Me.SuspendLayout()

        ' Título
        lblTitulo = New Label() With {
            .Text = If(esModoEdicion, "Editar Información del Propietario", "Agregar Nuevo Propietario"),
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(41, 128, 185),
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        Me.Controls.Add(lblTitulo)

        Dim yPos As Integer = 60

        ' Torre
        CrearCampo("Torre:", txtTorre, yPos)
        yPos += 40

        ' Piso
        CrearCampo("Piso:", txtPiso, yPos)
        yPos += 40

        ' Número de Apartamento
        CrearCampo("Número Apartamento:", txtNumeroApartamento, yPos)
        yPos += 40

        ' Nombre del Residente
        CrearCampo("Nombre Residente:", txtNombreResidente, yPos)
        yPos += 40

        ' Teléfono
        CrearCampo("Teléfono:", txtTelefono, yPos)
        yPos += 40

        ' Correo
        CrearCampo("Correo:", txtCorreo, yPos)
        yPos += 40

        ' Matrícula Inmobiliaria
        CrearCampo("Matrícula Inmobiliaria:", txtMatriculaInmobiliaria, yPos)
        yPos += 60

        ' Botones
        CrearBotones(yPos)

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Sub CrearCampo(labelText As String, ByRef textBox As TextBox, yPos As Integer)
        Dim lbl As New Label() With {
            .Text = labelText,
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(20, yPos),
            .Size = New Size(140, 20),
            .ForeColor = Color.FromArgb(52, 73, 94)
        }
        Me.Controls.Add(lbl)

        textBox = New TextBox() With {
            .Location = New Point(170, yPos - 3),
            .Size = New Size(240, 25),
            .Font = New Font("Segoe UI", 9),
            .BorderStyle = BorderStyle.FixedSingle
        }
        Me.Controls.Add(textBox)
    End Sub

    Private Sub CrearBotones(yPos As Integer)
        ' Botón Guardar
        btnGuardar = New Button() With {
            .Text = If(esModoEdicion, "Actualizar", "Agregar"),
            .Size = New Size(120, 40),
            .Location = New Point(170, yPos),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .UseVisualStyleBackColor = False
        }
        btnGuardar.FlatAppearance.BorderSize = 0
        AddHandler btnGuardar.Click, AddressOf btnGuardar_Click
        Me.Controls.Add(btnGuardar)

        ' Botón Cancelar
        btnCancelar = New Button() With {
            .Text = "Cancelar",
            .Size = New Size(120, 40),
            .Location = New Point(300, yPos),
            .BackColor = Color.FromArgb(192, 57, 43),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .UseVisualStyleBackColor = False
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
        Me.Controls.Add(btnCancelar)

        ' Hacer que Enter active el botón Guardar
        Me.AcceptButton = btnGuardar
        Me.CancelButton = btnCancelar
    End Sub

    Private Sub ConfigurarModoAgregar()
        ' Valores por defecto para nuevo apartamento
        txtTorre.Text = "1"
        txtPiso.Text = "1"
        txtNumeroApartamento.Text = ""
        txtNombreResidente.Text = ""
        txtTelefono.Text = ""
        txtCorreo.Text = ""
        txtMatriculaInmobiliaria.Text = ""

        lblTitulo.Text = "Agregar Nuevo Propietario"
    End Sub

    Private Sub ConfigurarModoEdicion()
        ' Cargar datos del apartamento existente
        If apartamentoActual IsNot Nothing Then
            txtTorre.Text = apartamentoActual.Torre.ToString()
            txtPiso.Text = apartamentoActual.Piso.ToString()
            txtNumeroApartamento.Text = apartamentoActual.NumeroApartamento
            txtNombreResidente.Text = If(apartamentoActual.NombreResidente, "")
            txtTelefono.Text = If(apartamentoActual.Telefono, "")
            txtCorreo.Text = If(apartamentoActual.Correo, "")
            txtMatriculaInmobiliaria.Text = If(apartamentoActual.MatriculaInmobiliaria, "")

            ' Actualizar título con información específica
            lblTitulo.Text = $"Editar Propietario - Torre {apartamentoActual.Torre}, Apt {apartamentoActual.NumeroApartamento}"
        End If

        ' En modo edición, deshabilitamos los campos que no se deben cambiar
        txtTorre.ReadOnly = True
        txtPiso.ReadOnly = True
        txtNumeroApartamento.ReadOnly = True
        txtTorre.BackColor = Color.FromArgb(230, 230, 230)
        txtPiso.BackColor = Color.FromArgb(230, 230, 230)
        txtNumeroApartamento.BackColor = Color.FromArgb(230, 230, 230)
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)
        Try
            ' Deshabilitar botón para evitar doble clic
            btnGuardar.Enabled = False
            btnGuardar.Text = "Guardando..."

            ' Validar campos obligatorios
            If Not ValidarCampos() Then
                btnGuardar.Enabled = True
                btnGuardar.Text = If(esModoEdicion, "Actualizar", "Agregar")
                Return
            End If

            ' Actualizar objeto apartamento con los datos del formulario
            ActualizarObjetoApartamento()

            ' Guardar en la base de datos
            If esModoEdicion Then
                ' Actualizar apartamento existente
                Dim resultado As Boolean = ApartamentoDAL.ActualizarPropietario(apartamentoActual)

                If resultado Then
                    MessageBox.Show("Propietario actualizado exitosamente.", "Éxito",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    MessageBox.Show("No se pudo actualizar el propietario. Verifique que el apartamento existe.",
                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    btnGuardar.Enabled = True
                    btnGuardar.Text = "Actualizar"
                End If
            Else
                ' Crear nuevo apartamento
                Dim nuevoId As Integer = ApartamentoDAL.CrearApartamento(apartamentoActual)

                If nuevoId > 0 Then
                    apartamentoActual.IdApartamento = nuevoId
                    MessageBox.Show("Nuevo propietario agregado exitosamente.", "Éxito",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    MessageBox.Show("No se pudo agregar el nuevo propietario. Verifique que no exista un apartamento con los mismos datos.",
                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    btnGuardar.Enabled = True
                    btnGuardar.Text = "Agregar"
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error al guardar: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnGuardar.Enabled = True
            btnGuardar.Text = If(esModoEdicion, "Actualizar", "Agregar")
        End Try
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        ' Confirmar cancelación si hay cambios
        If HayCambios() Then
            Dim resultado As DialogResult = MessageBox.Show(
                "¿Está seguro que desea cancelar? Se perderán los cambios realizados.",
                "Confirmar Cancelación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If resultado = DialogResult.No Then
                Return
            End If
        End If

        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Function ValidarCampos() As Boolean
        ' Validar torre
        Dim torre As Integer
        If Not Integer.TryParse(txtTorre.Text.Trim(), torre) OrElse torre < 1 OrElse torre > 8 Then
            MessageBox.Show("La torre debe ser un número entre 1 y 8.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTorre.Focus()
            Return False
        End If

        ' Validar piso
        Dim piso As Integer
        If Not Integer.TryParse(txtPiso.Text.Trim(), piso) OrElse piso < 1 OrElse piso > 5 Then
            MessageBox.Show("El piso debe ser un número entre 1 y 5.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPiso.Focus()
            Return False
        End If

        ' Validar número de apartamento
        If String.IsNullOrWhiteSpace(txtNumeroApartamento.Text) Then
            MessageBox.Show("El número de apartamento es obligatorio.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNumeroApartamento.Focus()
            Return False
        End If

        ' Validar que el número de apartamento sea numérico
        Dim numeroApt As Integer
        If Not Integer.TryParse(txtNumeroApartamento.Text.Trim(), numeroApt) Then
            MessageBox.Show("El número de apartamento debe ser numérico.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNumeroApartamento.Focus()
            Return False
        End If

        ' Validar nombre del residente
        If String.IsNullOrWhiteSpace(txtNombreResidente.Text) Then
            MessageBox.Show("El nombre del residente es obligatorio.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreResidente.Focus()
            Return False
        End If

        ' Validar longitud del nombre
        If txtNombreResidente.Text.Trim().Length < 2 Then
            MessageBox.Show("El nombre del residente debe tener al menos 2 caracteres.", "Validación",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreResidente.Focus()
            Return False
        End If

        ' Validar correo si se proporciona
        If Not String.IsNullOrWhiteSpace(txtCorreo.Text) Then
            If Not ValidarFormatoCorreo(txtCorreo.Text.Trim()) Then
                MessageBox.Show("El formato del correo electrónico no es válido.", "Validación",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCorreo.Focus()
                Return False
            End If
        End If

        ' Validar teléfono si se proporciona
        If Not String.IsNullOrWhiteSpace(txtTelefono.Text) Then
            Dim telefono As String = txtTelefono.Text.Trim().Replace(" ", "").Replace("-", "")
            If telefono.Length < 7 OrElse Not IsNumeric(telefono) Then
                MessageBox.Show("El teléfono debe tener al menos 7 dígitos y solo contener números.",
                              "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtTelefono.Focus()
                Return False
            End If
        End If

        Return True
    End Function

    Private Function ValidarFormatoCorreo(correo As String) As Boolean
        Try
            Dim addr As New System.Net.Mail.MailAddress(correo)
            Return addr.Address = correo AndAlso correo.Contains("@") AndAlso correo.Contains(".")
        Catch
            Return False
        End Try
    End Function

    Private Sub ActualizarObjetoApartamento()
        apartamentoActual.Torre = Integer.Parse(txtTorre.Text.Trim())
        apartamentoActual.Piso = Integer.Parse(txtPiso.Text.Trim())
        apartamentoActual.NumeroApartamento = txtNumeroApartamento.Text.Trim()
        apartamentoActual.NombreResidente = txtNombreResidente.Text.Trim()
        apartamentoActual.Telefono = txtTelefono.Text.Trim()
        apartamentoActual.Correo = txtCorreo.Text.Trim()
        apartamentoActual.MatriculaInmobiliaria = txtMatriculaInmobiliaria.Text.Trim()

        ' Si es nuevo apartamento, establecer valores por defecto
        If Not esModoEdicion Then
            apartamentoActual.Activo = True
            apartamentoActual.FechaRegistro = DateTime.Now
        End If
    End Sub

    Private Function HayCambios() As Boolean
        If Not esModoEdicion Then
            ' Si es nuevo, verificar si hay algún dato ingresado
            Return Not String.IsNullOrWhiteSpace(txtNombreResidente.Text) OrElse
                   Not String.IsNullOrWhiteSpace(txtTelefono.Text) OrElse
                   Not String.IsNullOrWhiteSpace(txtCorreo.Text) OrElse
                   Not String.IsNullOrWhiteSpace(txtMatriculaInmobiliaria.Text) OrElse
                   txtTorre.Text <> "1" OrElse
                   txtPiso.Text <> "1" OrElse
                   Not String.IsNullOrWhiteSpace(txtNumeroApartamento.Text)
        Else
            ' Si es edición, comparar con valores originales
            Return txtNombreResidente.Text.Trim() <> If(apartamentoActual.NombreResidente, "") OrElse
                   txtTelefono.Text.Trim() <> If(apartamentoActual.Telefono, "") OrElse
                   txtCorreo.Text.Trim() <> If(apartamentoActual.Correo, "") OrElse
                   txtMatriculaInmobiliaria.Text.Trim() <> If(apartamentoActual.MatriculaInmobiliaria, "")
        End If
    End Function

    Private Sub FormDetallePropietario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Establecer foco inicial
        If esModoEdicion Then
            txtNombreResidente.Focus()
            txtNombreResidente.SelectAll()
        Else
            txtTorre.Focus()
            txtTorre.SelectAll()
        End If
    End Sub

    Private Sub FormDetallePropietario_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Permitir cerrar con Escape
        If e.KeyCode = Keys.Escape Then
            btnCancelar_Click(Nothing, Nothing)
        End If
    End Sub

End Class