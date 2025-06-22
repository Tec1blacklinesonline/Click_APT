Imports System.Windows.Forms
Imports System.Drawing
Imports BCrypt.Net
Imports System.Data.SQLite

Public Class FormRegistro
    Inherits Form

    Private tabControl As TabControl
    Private btnGuardar As Button
    Private btnCancelar As Button

    ' Controles para registro de usuarios
    Private txtNombreUsuario As TextBox
    Private txtNombreCompleto As TextBox
    Private txtEmail As TextBox
    Private txtContrasena As TextBox
    Private txtConfirmarContrasena As TextBox
    Private cboRol As ComboBox

    ' Controles para registro de cuotas
    Private cboAsamblea As ComboBox
    Private cboTipoPiso As ComboBox
    Private txtValorCuota As TextBox
    Private dtpFechaVencimiento As DateTimePicker
    Private txtDescripcionCuota As TextBox
    Private chkGenerarParaTodos As CheckBox

    ' Controles para gestión de usuarios
    Private lstUsuarios As ListBox
    Private txtNuevaContrasena As TextBox
    Private txtConfirmarNuevaContrasena As TextBox
    Private btnEliminarUsuario As Button
    Private btnCambiarContrasena As Button
    Private btnActivarDesactivar As Button
    Private lblUsuarioSeleccionado As Label

    Private Sub FormRegistro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarFormulario()
        CargarDatosIniciales()
    End Sub

    Private Sub ConfigurarFormulario()
        ' CORREGIDO: Formulario más grande y mejor centrado
        Me.Text = "Registro de Usuarios y Cuotas"
        Me.Size = New Size(700, 650)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(240, 240, 240)
        Me.ShowInTaskbar = False ' No aparece en la barra de tareas

        ' Panel superior - CORREGIDO: Altura ajustada
        Dim panelSuperior As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 70,
            .BackColor = Color.FromArgb(231, 76, 60)
        }

        Dim lblTitulo As New Label With {
            .Text = "📝 REGISTRO Y GESTIÓN DE USUARIOS",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        panelSuperior.Controls.Add(lblTitulo)
        Me.Controls.Add(panelSuperior)

        ' Panel inferior con botones - CORREGIDO: Botones centrados
        Dim panelInferior As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 70,
            .BackColor = Color.FromArgb(236, 240, 241)
        }

        ' Calcular posición centrada para los botones
        Dim anchoBoton As Integer = 120
        Dim espacioEntreBot As Integer = 20
        Dim anchoTotal As Integer = (anchoBoton * 2) + espacioEntreBot
        Dim xStart As Integer = (Me.Width - anchoTotal) \ 2

        btnGuardar = New Button With {
            .Text = "💾 Guardar",
            .Size = New Size(anchoBoton, 40),
            .Location = New Point(xStart, 15),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnGuardar.FlatAppearance.BorderSize = 0
        AddHandler btnGuardar.Click, AddressOf btnGuardar_Click
        panelInferior.Controls.Add(btnGuardar)

        btnCancelar = New Button With {
            .Text = "❌ Cancelar",
            .Size = New Size(anchoBoton, 40),
            .Location = New Point(xStart + anchoBoton + espacioEntreBot, 15),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf btnCancelar_Click
        panelInferior.Controls.Add(btnCancelar)

        Me.Controls.Add(panelInferior)

        ' TabControl - CORREGIDO: Configuración más explícita para mostrar pestañas
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Padding = New Point(12, 8),
            .Appearance = TabAppearance.Normal,
            .Alignment = TabAlignment.Top,
            .Multiline = False,
            .SizeMode = TabSizeMode.Normal,
            .DrawMode = TabDrawMode.Normal,
            .HotTrack = True
        }

        ' Tab Registro de Usuarios - CORREGIDO: Configuración más explícita
        Dim tabUsuarios As New TabPage With {
            .Text = "👤 Registrar Usuario",
            .BackColor = Color.White,
            .UseVisualStyleBackColor = True
        }
        CrearTabUsuarios(tabUsuarios)

        ' Tab Registro de Cuotas - CORREGIDO: Configuración más explícita
        Dim tabCuotas As New TabPage With {
            .Text = "💰 Generar Cuotas",
            .BackColor = Color.White,
            .UseVisualStyleBackColor = True
        }
        CrearTabCuotas(tabCuotas)

        ' Tab Gestión de Usuarios - NUEVO - Configuración más explícita
        Dim tabGestion As New TabPage With {
            .Text = "🔧 Gestionar Usuarios",
            .BackColor = Color.White,
            .UseVisualStyleBackColor = True
        }
        CrearTabGestionUsuarios(tabGestion)

        ' Agregar pestañas en orden
        tabControl.TabPages.Clear()
        tabControl.TabPages.Add(tabUsuarios)
        tabControl.TabPages.Add(tabCuotas)
        tabControl.TabPages.Add(tabGestion)

        ' Asegurar que la primera pestaña esté seleccionada
        tabControl.SelectedIndex = 0

        Me.Controls.Add(tabControl)

        ' CORREGIDO: Asegurar que el TabControl esté en el frente y sea visible
        tabControl.BringToFront()
        tabControl.Visible = True

        ' Evento para manejar cambios de pestaña
        AddHandler tabControl.SelectedIndexChanged, AddressOf tabControl_SelectedIndexChanged
    End Sub

    Private Sub CrearTabUsuarios(tab As TabPage)
        tab.BackColor = Color.White
        tab.Padding = New Padding(20)

        ' CORREGIDO: Layout mejorado con espaciado uniforme
        Dim inicioY As Integer = 30
        Dim espacioY As Integer = 50
        Dim labelX As Integer = 40
        Dim controlX As Integer = 200
        Dim controlWidth As Integer = 350
        Dim actualY As Integer = inicioY

        ' Nombre de usuario
        Dim lblNombreUsuario As New Label With {
            .Text = "Nombre de Usuario:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblNombreUsuario)

        txtNombreUsuario = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .Font = New Font("Segoe UI", 11),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtNombreUsuario)

        actualY += espacioY

        ' Nombre completo
        Dim lblNombreCompleto As New Label With {
            .Text = "Nombre Completo:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblNombreCompleto)

        txtNombreCompleto = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .Font = New Font("Segoe UI", 11),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtNombreCompleto)

        actualY += espacioY

        ' Email
        Dim lblEmail As New Label With {
            .Text = "Correo Electrónico:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblEmail)

        txtEmail = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .Font = New Font("Segoe UI", 11),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtEmail)

        actualY += espacioY

        ' Rol
        Dim lblRol As New Label With {
            .Text = "Rol:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblRol)

        cboRol = New ComboBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(250, 30),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 11)
        }
        cboRol.Items.AddRange({"Administrador", "Operador", "Consulta"})
        cboRol.SelectedIndex = 1 ' Operador por defecto
        tab.Controls.Add(cboRol)

        actualY += espacioY

        ' Contraseña
        Dim lblContrasena As New Label With {
            .Text = "Contraseña:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblContrasena)

        txtContrasena = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .PasswordChar = "*"c,
            .Font = New Font("Segoe UI", 11),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtContrasena)

        actualY += espacioY

        ' Confirmar contraseña
        Dim lblConfirmar As New Label With {
            .Text = "Confirmar Contraseña:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblConfirmar)

        txtConfirmarContrasena = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .PasswordChar = "*"c,
            .Font = New Font("Segoe UI", 11),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtConfirmarContrasena)

        actualY += espacioY + 20

        ' Nota informativa - CORREGIDO: Mejor posicionamiento
        Dim lblNota As New Label With {
            .Text = "💡 Nota: La contraseña debe tener al menos 8 caracteres.",
            .Location = New Point(labelX, actualY),
            .Size = New Size(500, 25),
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(52, 152, 219)
        }
        tab.Controls.Add(lblNota)
    End Sub

    Private Sub CrearTabCuotas(tab As TabPage)
        tab.BackColor = Color.White
        tab.Padding = New Padding(20)

        ' CORREGIDO: Layout mejorado con espaciado uniforme
        Dim inicioY As Integer = 30
        Dim espacioY As Integer = 50
        Dim labelX As Integer = 40
        Dim controlX As Integer = 200
        Dim controlWidth As Integer = 350
        Dim actualY As Integer = inicioY

        ' Asamblea
        Dim lblAsamblea As New Label With {
            .Text = "Asamblea:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblAsamblea)

        cboAsamblea = New ComboBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 11)
        }
        tab.Controls.Add(cboAsamblea)

        ' NUEVO: Botón para crear nueva asamblea
        Dim btnNuevaAsamblea As New Button With {
            .Text = "➕ Nueva Asamblea",
            .Location = New Point(controlX + controlWidth + 10, actualY),
            .Size = New Size(120, 30),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnNuevaAsamblea.FlatAppearance.BorderSize = 0
        AddHandler btnNuevaAsamblea.Click, AddressOf btnNuevaAsamblea_Click
        tab.Controls.Add(btnNuevaAsamblea)

        actualY += espacioY

        ' Tipo de piso
        Dim lblTipoPiso As New Label With {
            .Text = "Tipo de Piso:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblTipoPiso)

        cboTipoPiso = New ComboBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(250, 30),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 11)
        }
        cboTipoPiso.Items.AddRange({"Todos", "Primer Piso", "Segundo Piso", "Tercer Piso", "Cuarto Piso", "Quinto Piso"})
        cboTipoPiso.SelectedIndex = 0
        tab.Controls.Add(cboTipoPiso)

        actualY += espacioY

        ' Valor de la cuota
        Dim lblValor As New Label With {
            .Text = "Valor de la Cuota:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblValor)

        txtValorCuota = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(180, 30),
            .Font = New Font("Segoe UI", 11),
            .TextAlign = HorizontalAlignment.Right,
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtValorCuota)

        Dim lblPesos As New Label With {
            .Text = "COP",
            .Location = New Point(controlX + 190, actualY + 5),
            .Size = New Size(40, 20),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.Gray
        }
        tab.Controls.Add(lblPesos)

        actualY += espacioY

        ' Fecha de vencimiento
        Dim lblFechaVenc As New Label With {
            .Text = "Inicio a regir:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblFechaVenc)

        dtpFechaVencimiento = New DateTimePicker With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(200, 30),
            .Format = DateTimePickerFormat.Short,
            .Font = New Font("Segoe UI", 11),
            .Value = DateTime.Now.AddDays(30)
        }
        tab.Controls.Add(dtpFechaVencimiento)

        actualY += espacioY

        ' Descripción
        Dim lblDescripcion As New Label With {
            .Text = "Descripción:",
            .Location = New Point(labelX, actualY),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        tab.Controls.Add(lblDescripcion)

        txtDescripcionCuota = New TextBox With {
            .Location = New Point(controlX, actualY),
            .Size = New Size(controlWidth, 30),
            .Font = New Font("Segoe UI", 11),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tab.Controls.Add(txtDescripcionCuota)

        actualY += espacioY

        ' Checkbox generar para todos
        chkGenerarParaTodos = New CheckBox With {
            .Text = "✅ Generar cuota para todos los apartamentos",
            .Location = New Point(labelX, actualY),
            .Size = New Size(400, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.FromArgb(39, 174, 96),
            .Checked = True
        }
        tab.Controls.Add(chkGenerarParaTodos)

        actualY += 40

        ' Nota informativa - CORREGIDO: Mejor diseño
        Dim lblNotaCuotas As New Label With {
            .Text = "💡 Nota: Se generarán cuotas individuales para cada apartamento según el tipo de piso seleccionado.",
            .Location = New Point(labelX, actualY),
            .Size = New Size(550, 40),
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(52, 152, 219)
        }
        tab.Controls.Add(lblNotaCuotas)
    End Sub

    Private Sub CrearTabGestionUsuarios(tab As TabPage)
        tab.BackColor = Color.White
        tab.Padding = New Padding(20)

        ' Panel izquierdo - Lista de usuarios
        Dim panelIzquierdo As New Panel With {
            .Location = New Point(20, 20),
            .Size = New Size(280, 400),
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.FromArgb(248, 248, 248)
        }
        tab.Controls.Add(panelIzquierdo)

        Dim lblTituloLista As New Label With {
            .Text = "👥 USUARIOS DEL SISTEMA",
            .Location = New Point(10, 10),
            .Size = New Size(260, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White
        }
        panelIzquierdo.Controls.Add(lblTituloLista)

        lstUsuarios = New ListBox With {
            .Location = New Point(10, 45),
            .Size = New Size(260, 300),
            .Font = New Font("Segoe UI", 9),
            .BorderStyle = BorderStyle.FixedSingle
        }
        AddHandler lstUsuarios.SelectedIndexChanged, AddressOf lstUsuarios_SelectedIndexChanged
        panelIzquierdo.Controls.Add(lstUsuarios)

        Dim btnActualizarLista As New Button With {
            .Text = "🔄 Actualizar Lista",
            .Location = New Point(10, 355),
            .Size = New Size(260, 30),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnActualizarLista.FlatAppearance.BorderSize = 0
        AddHandler btnActualizarLista.Click, AddressOf CargarListaUsuarios
        panelIzquierdo.Controls.Add(btnActualizarLista)

        ' Panel derecho - Gestión
        Dim panelDerecho As New Panel With {
            .Location = New Point(320, 20),
            .Size = New Size(340, 400),
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White
        }
        tab.Controls.Add(panelDerecho)

        Dim lblTituloGestion As New Label With {
            .Text = "🔧 GESTIÓN DE USUARIO",
            .Location = New Point(10, 10),
            .Size = New Size(320, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.FromArgb(155, 89, 182),
            .ForeColor = Color.White
        }
        panelDerecho.Controls.Add(lblTituloGestion)

        ' Usuario seleccionado
        lblUsuarioSeleccionado = New Label With {
            .Text = "Seleccione un usuario de la lista",
            .Location = New Point(20, 50),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Italic),
            .ForeColor = Color.Gray
        }
        panelDerecho.Controls.Add(lblUsuarioSeleccionado)

        ' Sección cambio de contraseña
        Dim lblCambioPass As New Label With {
            .Text = "🔑 CAMBIAR CONTRASEÑA",
            .Location = New Point(20, 90),
            .Size = New Size(300, 20),
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(39, 174, 96)
        }
        panelDerecho.Controls.Add(lblCambioPass)

        Dim lblNuevaPass As New Label With {
            .Text = "Nueva Contraseña:",
            .Location = New Point(20, 120),
            .Size = New Size(120, 20),
            .Font = New Font("Segoe UI", 9)
        }
        panelDerecho.Controls.Add(lblNuevaPass)

        txtNuevaContrasena = New TextBox With {
            .Location = New Point(145, 118),
            .Size = New Size(170, 25),
            .PasswordChar = "*"c,
            .Font = New Font("Segoe UI", 10),
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelDerecho.Controls.Add(txtNuevaContrasena)

        Dim lblConfirmarNueva As New Label With {
            .Text = "Confirmar:",
            .Location = New Point(20, 150),
            .Size = New Size(120, 20),
            .Font = New Font("Segoe UI", 9)
        }
        panelDerecho.Controls.Add(lblConfirmarNueva)

        txtConfirmarNuevaContrasena = New TextBox With {
            .Location = New Point(145, 148),
            .Size = New Size(170, 25),
            .PasswordChar = "*"c,
            .Font = New Font("Segoe UI", 10),
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelDerecho.Controls.Add(txtConfirmarNuevaContrasena)

        btnCambiarContrasena = New Button With {
            .Text = "🔑 Cambiar Contraseña",
            .Location = New Point(20, 185),
            .Size = New Size(295, 35),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .Enabled = False
        }
        btnCambiarContrasena.FlatAppearance.BorderSize = 0
        AddHandler btnCambiarContrasena.Click, AddressOf btnCambiarContrasena_Click
        panelDerecho.Controls.Add(btnCambiarContrasena)

        ' Separador
        Dim separador As New Panel With {
            .Location = New Point(20, 240),
            .Size = New Size(295, 2),
            .BackColor = Color.LightGray
        }
        panelDerecho.Controls.Add(separador)

        ' Sección acciones de usuario
        Dim lblAcciones As New Label With {
            .Text = "⚠️ ACCIONES DE USUARIO",
            .Location = New Point(20, 255),
            .Size = New Size(300, 20),
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(231, 76, 60)
        }
        panelDerecho.Controls.Add(lblAcciones)

        btnActivarDesactivar = New Button With {
            .Text = "🔄 Activar/Desactivar",
            .Location = New Point(20, 285),
            .Size = New Size(140, 35),
            .BackColor = Color.FromArgb(243, 156, 18),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .Enabled = False
        }
        btnActivarDesactivar.FlatAppearance.BorderSize = 0
        AddHandler btnActivarDesactivar.Click, AddressOf btnActivarDesactivar_Click
        panelDerecho.Controls.Add(btnActivarDesactivar)

        btnEliminarUsuario = New Button With {
            .Text = "🗑️ Eliminar Usuario",
            .Location = New Point(175, 285),
            .Size = New Size(140, 35),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .Enabled = False
        }
        btnEliminarUsuario.FlatAppearance.BorderSize = 0
        AddHandler btnEliminarUsuario.Click, AddressOf btnEliminarUsuario_Click
        panelDerecho.Controls.Add(btnEliminarUsuario)

        ' Nota de seguridad
        Dim lblNotaSeguridad As New Label With {
            .Text = "⚠️ ATENCIÓN: Las acciones de eliminación son irreversibles. " &
                   "Se recomienda desactivar usuarios en lugar de eliminarlos.",
            .Location = New Point(20, 335),
            .Size = New Size(295, 50),
            .Font = New Font("Segoe UI", 8, FontStyle.Italic),
            .ForeColor = Color.FromArgb(231, 76, 60)
        }
        panelDerecho.Controls.Add(lblNotaSeguridad)
    End Sub

    Private Sub CargarDatosIniciales()
        Try
            ' Cargar asambleas
            CargarAsambleas()

            If cboAsamblea.Items.Count > 0 Then
                cboAsamblea.SelectedIndex = 0
            End If

            ' Cargar lista de usuarios
            CargarListaUsuarios()

        Catch ex As Exception
            MessageBox.Show($"Error al cargar datos iniciales: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargarListaUsuarios()
        Try
            lstUsuarios.Items.Clear()
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT id_usuario, nombre_usuario, nombre_completo, rol, activo FROM Usuarios ORDER BY nombre_usuario"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using lector As SQLiteDataReader = comando.ExecuteReader()
                        While lector.Read()
                            Dim estado As String = If(Convert.ToBoolean(lector("activo")), "✅", "❌")
                            Dim item As String = $"{estado} {lector("nombre_usuario")} - {lector("nombre_completo")} ({lector("rol")})"

                            ' Crear objeto anónimo con los datos del usuario
                            Dim usuarioData = New With {
                                .Text = item,
                                .Id = Convert.ToInt32(lector("id_usuario")),
                                .NombreUsuario = lector("nombre_usuario").ToString(),
                                .Activo = Convert.ToBoolean(lector("activo"))
                            }

                            lstUsuarios.Items.Add(usuarioData)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al cargar usuarios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub lstUsuarios_SelectedIndexChanged(sender As Object, e As EventArgs)
        If lstUsuarios.SelectedIndex >= 0 Then
            Dim usuarioSeleccionado = lstUsuarios.SelectedItem
            lblUsuarioSeleccionado.Text = $"Usuario: {usuarioSeleccionado.NombreUsuario}"
            lblUsuarioSeleccionado.ForeColor = Color.Black

            ' Habilitar botones
            btnCambiarContrasena.Enabled = True
            btnActivarDesactivar.Enabled = True
            btnEliminarUsuario.Enabled = True

            ' Actualizar texto del botón activar/desactivar
            If usuarioSeleccionado.Activo Then
                btnActivarDesactivar.Text = "❌ Desactivar Usuario"
                btnActivarDesactivar.BackColor = Color.FromArgb(243, 156, 18)
            Else
                btnActivarDesactivar.Text = "✅ Activar Usuario"
                btnActivarDesactivar.BackColor = Color.FromArgb(39, 174, 96)
            End If
        Else
            lblUsuarioSeleccionado.Text = "Seleccione un usuario de la lista"
            lblUsuarioSeleccionado.ForeColor = Color.Gray
            btnCambiarContrasena.Enabled = False
            btnActivarDesactivar.Enabled = False
            btnEliminarUsuario.Enabled = False
        End If
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)
        Try
            If tabControl.SelectedTab.Text.Contains("Usuario") Then
                GuardarUsuario()
            ElseIf tabControl.SelectedTab.Text.Contains("Cuotas") Then
                GenerarCuotas()
            Else
                MessageBox.Show("Use los botones específicos de la pestaña de gestión.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GuardarUsuario()
        ' Validaciones
        If String.IsNullOrWhiteSpace(txtNombreUsuario.Text) Then
            MessageBox.Show("El nombre de usuario es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreUsuario.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtNombreCompleto.Text) Then
            MessageBox.Show("El nombre completo es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNombreCompleto.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtEmail.Text) OrElse Not txtEmail.Text.Contains("@") Then
            MessageBox.Show("Debe ingresar un correo electrónico válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtEmail.Focus()
            Return
        End If

        If txtContrasena.Text.Length < 8 Then
            MessageBox.Show("La contraseña debe tener al menos 8 caracteres.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContrasena.Focus()
            Return
        End If

        If txtContrasena.Text <> txtConfirmarContrasena.Text Then
            MessageBox.Show("Las contraseñas no coinciden.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtConfirmarContrasena.Focus()
            Return
        End If

        Try
            ' Encriptar contraseña
            Dim hashContrasena As String = BCrypt.Net.BCrypt.HashPassword(txtContrasena.Text)

            ' Insertar usuario en base de datos
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Verificar que no exista el usuario
                Dim consultaExiste As String = "SELECT COUNT(*) FROM Usuarios WHERE nombre_usuario = @usuario"
                Using comandoExiste As New SQLiteCommand(consultaExiste, conexion)
                    comandoExiste.Parameters.AddWithValue("@usuario", txtNombreUsuario.Text.Trim())
                    If Convert.ToInt32(comandoExiste.ExecuteScalar()) > 0 Then
                        MessageBox.Show("Ya existe un usuario con ese nombre.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using

                ' Insertar nuevo usuario - CORREGIDO: Usar nombres exactos de tu tabla
                Dim consultaInsert As String = "INSERT INTO Usuarios (nombre_usuario, contrasena_hash, nombre_completo, email, rol, fecha_creacion, ultimo_acceso, activo) VALUES (@usuario, @contrasena, @nombre, @email, @rol, @fechaCreacion, @ultimoAcceso, @activo)"
                Using comando As New SQLiteCommand(consultaInsert, conexion)
                    comando.Parameters.AddWithValue("@usuario", txtNombreUsuario.Text.Trim())
                    comando.Parameters.AddWithValue("@contrasena", hashContrasena)
                    comando.Parameters.AddWithValue("@nombre", txtNombreCompleto.Text.Trim())
                    comando.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
                    comando.Parameters.AddWithValue("@rol", cboRol.SelectedItem.ToString())
                    comando.Parameters.AddWithValue("@fechaCreacion", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    comando.Parameters.AddWithValue("@ultimoAcceso", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    comando.Parameters.AddWithValue("@activo", 1)

                    If comando.ExecuteNonQuery() > 0 Then
                        MessageBox.Show("Usuario registrado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LimpiarFormularioUsuario()
                        CargarListaUsuarios() ' Actualizar lista
                    Else
                        MessageBox.Show("No se pudo registrar el usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al registrar usuario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GenerarCuotas()
        ' Validaciones
        If cboAsamblea.SelectedIndex = -1 Then
            MessageBox.Show("Debe seleccionar una asamblea.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cboAsamblea.Focus()
            Return
        End If

        Dim valorCuotaDecimal As Decimal
        If Not Decimal.TryParse(txtValorCuota.Text, valorCuotaDecimal) OrElse valorCuotaDecimal <= 0 Then
            MessageBox.Show("Debe ingresar un valor de cuota válido.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtValorCuota.Focus()
            Return
        End If

        If String.IsNullOrWhiteSpace(txtDescripcionCuota.Text) Then
            MessageBox.Show("La descripción de la cuota es obligatoria.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDescripcionCuota.Focus()
            Return
        End If

        Try
            ' Obtener ID de la asamblea seleccionada
            Dim idAsamblea As Integer = 0
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consultaAsamblea As String = "SELECT id_asamblea FROM Asambleas ORDER BY fecha_asamblea DESC LIMIT 1 OFFSET @indice"
                Using comandoAsamblea As New SQLiteCommand(consultaAsamblea, conexion)
                    comandoAsamblea.Parameters.AddWithValue("@indice", cboAsamblea.SelectedIndex)
                    Dim resultado = comandoAsamblea.ExecuteScalar()
                    If resultado IsNot Nothing Then
                        idAsamblea = Convert.ToInt32(resultado)
                    End If
                End Using
            End Using

            ' Generar cuotas
            Dim apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()
            Dim cuotasGeneradas As Integer = 0

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        For Each apartamento In apartamentos
                            ' Filtrar por tipo de piso si no es "Todos"
                            If cboTipoPiso.SelectedIndex > 0 Then
                                Dim pisoSeleccionado As Integer = cboTipoPiso.SelectedIndex
                                If apartamento.Piso <> pisoSeleccionado Then
                                    Continue For
                                End If
                            End If

                            ' Insertar cuota generada para el apartamento
                            Dim consultaInsert As String = "INSERT INTO cuotas_generadas_apartamento (id_apartamentos, matricula_inmobiliaria, fecha_cuota, valor_cuota, fecha_inicio, estado, tipo_cuota, tipo_piso, id_asamblea) VALUES (@idApt, @matricula, date('now'), @valor, @fechaVenc, 'pendiente', @descripcion, @tipoPiso, @idAsamblea)"

                            Using comando As New SQLiteCommand(consultaInsert, conexion, transaccion)
                                comando.Parameters.AddWithValue("@idApt", apartamento.IdApartamento)
                                comando.Parameters.AddWithValue("@matricula", apartamento.MatriculaInmobiliaria)
                                comando.Parameters.AddWithValue("@valor", valorCuotaDecimal)
                                comando.Parameters.AddWithValue("@fechaVenc", dtpFechaVencimiento.Value.ToString("yyyy-MM-dd"))
                                comando.Parameters.AddWithValue("@descripcion", txtDescripcionCuota.Text.Trim())
                                comando.Parameters.AddWithValue("@tipoPiso", cboTipoPiso.SelectedItem.ToString())
                                comando.Parameters.AddWithValue("@idAsamblea", idAsamblea)

                                If comando.ExecuteNonQuery() > 0 Then
                                    cuotasGeneradas += 1
                                End If
                            End Using
                        Next

                        transaccion.Commit()
                        MessageBox.Show($"Se generaron {cuotasGeneradas} cuotas exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LimpiarFormularioCuotas()

                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al generar cuotas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCambiarContrasena_Click(sender As Object, e As EventArgs)
        If lstUsuarios.SelectedIndex = -1 Then
            MessageBox.Show("Debe seleccionar un usuario.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If txtNuevaContrasena.Text.Length < 8 Then
            MessageBox.Show("La nueva contraseña debe tener al menos 8 caracteres.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNuevaContrasena.Focus()
            Return
        End If

        If txtNuevaContrasena.Text <> txtConfirmarNuevaContrasena.Text Then
            MessageBox.Show("Las contraseñas no coinciden.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtConfirmarNuevaContrasena.Focus()
            Return
        End If

        Dim usuarioSeleccionado = lstUsuarios.SelectedItem
        Dim resultado = MessageBox.Show($"¿Está seguro de cambiar la contraseña del usuario '{usuarioSeleccionado.NombreUsuario}'?",
                                      "Confirmar Cambio", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If resultado = DialogResult.Yes Then
            Try
                Dim hashContrasena As String = BCrypt.Net.BCrypt.HashPassword(txtNuevaContrasena.Text)

                Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                    conexion.Open()
                    Dim consulta As String = "UPDATE Usuarios SET contrasena_hash = @contrasena WHERE id_usuario = @id"
                    Using comando As New SQLiteCommand(consulta, conexion)
                        comando.Parameters.AddWithValue("@contrasena", hashContrasena)
                        comando.Parameters.AddWithValue("@id", usuarioSeleccionado.Id)

                        If comando.ExecuteNonQuery() > 0 Then
                            MessageBox.Show("Contraseña cambiada exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            txtNuevaContrasena.Clear()
                            txtConfirmarNuevaContrasena.Clear()
                        Else
                            MessageBox.Show("No se pudo cambiar la contraseña.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error al cambiar contraseña: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnActivarDesactivar_Click(sender As Object, e As EventArgs)
        If lstUsuarios.SelectedIndex = -1 Then
            MessageBox.Show("Debe seleccionar un usuario.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim usuarioSeleccionado = lstUsuarios.SelectedItem
        Dim nuevoEstado As Boolean = Not usuarioSeleccionado.Activo
        Dim accion As String = If(nuevoEstado, "activar", "desactivar")

        Dim resultado = MessageBox.Show($"¿Está seguro de {accion} el usuario '{usuarioSeleccionado.NombreUsuario}'?",
                                      "Confirmar Acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If resultado = DialogResult.Yes Then
            Try
                Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                    conexion.Open()
                    Dim consulta As String = "UPDATE Usuarios SET activo = @activo WHERE id_usuario = @id"
                    Using comando As New SQLiteCommand(consulta, conexion)
                        comando.Parameters.AddWithValue("@activo", nuevoEstado)
                        comando.Parameters.AddWithValue("@id", usuarioSeleccionado.Id)

                        If comando.ExecuteNonQuery() > 0 Then
                            MessageBox.Show($"Usuario {accion} exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            CargarListaUsuarios()
                            lstUsuarios.SelectedIndex = -1
                        Else
                            MessageBox.Show($"No se pudo {accion} el usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error al {accion} usuario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnEliminarUsuario_Click(sender As Object, e As EventArgs)
        If lstUsuarios.SelectedIndex = -1 Then
            MessageBox.Show("Debe seleccionar un usuario.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim usuarioSeleccionado = lstUsuarios.SelectedItem

        ' Verificar que no sea el usuario actual
        Dim usuarioActual = ConexionBD.ObtenerUsuarioActual()
        If usuarioActual IsNot Nothing AndAlso usuarioActual.NombreUsuario = usuarioSeleccionado.NombreUsuario Then
            MessageBox.Show("No puede eliminar su propio usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Doble confirmación para eliminar
        Dim resultado1 = MessageBox.Show($"⚠️ ATENCIÓN: ¿Está seguro de ELIMINAR PERMANENTEMENTE el usuario '{usuarioSeleccionado.NombreUsuario}'?" & Environment.NewLine & Environment.NewLine & "Esta acción NO SE PUEDE DESHACER.",
                                       "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If resultado1 = DialogResult.Yes Then
            Dim resultado2 = MessageBox.Show("¿REALMENTE desea eliminar este usuario? Esta acción es IRREVERSIBLE.",
                                           "Confirmación Final", MessageBoxButtons.YesNo, MessageBoxIcon.Stop)

            If resultado2 = DialogResult.Yes Then
                Try
                    Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                        conexion.Open()
                        Dim consulta As String = "DELETE FROM Usuarios WHERE id_usuario = @id"
                        Using comando As New SQLiteCommand(consulta, conexion)
                            comando.Parameters.AddWithValue("@id", usuarioSeleccionado.Id)

                            If comando.ExecuteNonQuery() > 0 Then
                                MessageBox.Show("Usuario eliminado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                CargarListaUsuarios()
                                lstUsuarios.SelectedIndex = -1
                            Else
                                MessageBox.Show("No se pudo eliminar el usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show($"Error al eliminar usuario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    Private Sub LimpiarFormularioUsuario()
        txtNombreUsuario.Clear()
        txtNombreCompleto.Clear()
        txtEmail.Clear()
        txtContrasena.Clear()
        txtConfirmarContrasena.Clear()
        cboRol.SelectedIndex = 1
        txtNombreUsuario.Focus()
    End Sub

    Private Sub LimpiarFormularioCuotas()
        txtValorCuota.Clear()
        txtDescripcionCuota.Clear()
        dtpFechaVencimiento.Value = DateTime.Now.AddDays(30)
        cboTipoPiso.SelectedIndex = 0
        chkGenerarParaTodos.Checked = True
        txtValorCuota.Focus()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub btnNuevaAsamblea_Click(sender As Object, e As EventArgs)
        ' Crear formulario modal para nueva asamblea
        Dim formAsamblea As New Form With {
            .Text = "Nueva Asamblea",
            .Size = New Size(500, 350),
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .MaximizeBox = False,
            .MinimizeBox = False,
            .ShowInTaskbar = False,
            .BackColor = Color.White
        }

        ' Panel superior
        Dim panelTitulo As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.FromArgb(52, 152, 219)
        }

        Dim lblTituloAsamblea As New Label With {
            .Text = "➕ CREAR NUEVA ASAMBLEA",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(20, 15),
            .AutoSize = True
        }
        panelTitulo.Controls.Add(lblTituloAsamblea)
        formAsamblea.Controls.Add(panelTitulo)

        ' Campos del formulario
        Dim yPos As Integer = 80

        ' Nombre de la asamblea
        Dim lblNombre As New Label With {
            .Text = "Nombre de la Asamblea:",
            .Location = New Point(30, yPos),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        formAsamblea.Controls.Add(lblNombre)

        Dim txtNombreAsamblea As New TextBox With {
            .Location = New Point(190, yPos),
            .Size = New Size(250, 25),
            .Font = New Font("Segoe UI", 10)
        }
        formAsamblea.Controls.Add(txtNombreAsamblea)

        yPos += 40

        ' Fecha de la asamblea
        Dim lblFecha As New Label With {
            .Text = "Fecha de la Asamblea:",
            .Location = New Point(30, yPos),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        formAsamblea.Controls.Add(lblFecha)

        Dim dtpFechaAsamblea As New DateTimePicker With {
            .Location = New Point(190, yPos),
            .Size = New Size(200, 25),
            .Format = DateTimePickerFormat.Short,
            .Font = New Font("Segoe UI", 10)
        }
        formAsamblea.Controls.Add(dtpFechaAsamblea)

        yPos += 40

        ' Descripción
        Dim lblDescripcionAsamblea As New Label With {
            .Text = "Descripción:",
            .Location = New Point(30, yPos),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        formAsamblea.Controls.Add(lblDescripcionAsamblea)

        Dim txtDescripcionAsamblea As New TextBox With {
            .Location = New Point(190, yPos),
            .Size = New Size(250, 60),
            .Font = New Font("Segoe UI", 10),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical
        }
        formAsamblea.Controls.Add(txtDescripcionAsamblea)

        yPos += 80

        ' Botones
        Dim btnGuardarAsamblea As New Button With {
            .Text = "💾 Guardar Asamblea",
            .Location = New Point(190, yPos),
            .Size = New Size(130, 35),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnGuardarAsamblea.FlatAppearance.BorderSize = 0
        formAsamblea.Controls.Add(btnGuardarAsamblea)

        Dim btnCancelarAsamblea As New Button With {
            .Text = "❌ Cancelar",
            .Location = New Point(330, yPos),
            .Size = New Size(100, 35),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        btnCancelarAsamblea.FlatAppearance.BorderSize = 0
        formAsamblea.Controls.Add(btnCancelarAsamblea)

        ' Eventos de los botones
        AddHandler btnGuardarAsamblea.Click, Sub()
                                                 If String.IsNullOrWhiteSpace(txtNombreAsamblea.Text) Then
                                                     MessageBox.Show("El nombre de la asamblea es obligatorio.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                     txtNombreAsamblea.Focus()
                                                     Return
                                                 End If

                                                 Try
                                                     ' Crear nueva asamblea usando inserción directa en BD con nombres correctos
                                                     Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                                                         conexion.Open()
                                                         Dim consulta As String = "INSERT INTO Asambleas (nombre_asamblea, fecha_asamblea, descripcion) VALUES (@nombre, @fecha, @descripcion)"
                                                         Using comando As New SQLiteCommand(consulta, conexion)
                                                             comando.Parameters.AddWithValue("@nombre", txtNombreAsamblea.Text.Trim())
                                                             comando.Parameters.AddWithValue("@fecha", dtpFechaAsamblea.Value.ToString("yyyy-MM-dd"))
                                                             comando.Parameters.AddWithValue("@descripcion", txtDescripcionAsamblea.Text.Trim())

                                                             If comando.ExecuteNonQuery() > 0 Then
                                                                 MessageBox.Show("Asamblea creada exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                 formAsamblea.DialogResult = DialogResult.OK
                                                                 formAsamblea.Close()
                                                             Else
                                                                 MessageBox.Show("No se pudo crear la asamblea.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                             End If
                                                         End Using
                                                     End Using

                                                 Catch ex As Exception
                                                     MessageBox.Show($"Error al crear asamblea: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                 End Try
                                             End Sub

        AddHandler btnCancelarAsamblea.Click, Sub()
                                                  formAsamblea.DialogResult = DialogResult.Cancel
                                                  formAsamblea.Close()
                                              End Sub

        ' Mostrar el formulario modal
        If formAsamblea.ShowDialog() = DialogResult.OK Then
            ' Recargar lista de asambleas y seleccionar la nueva
            CargarAsambleas()
            If cboAsamblea.Items.Count > 0 Then
                cboAsamblea.SelectedIndex = cboAsamblea.Items.Count - 1 ' Seleccionar la última (recién creada)
            End If
        End If

        formAsamblea.Dispose()
    End Sub

    Private Sub CargarAsambleas()
        Try
            cboAsamblea.Items.Clear()
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT id_asamblea, nombre_asamblea, fecha_asamblea FROM Asambleas ORDER BY fecha_asamblea DESC"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using lector As SQLiteDataReader = comando.ExecuteReader()
                        While lector.Read()
                            Dim fechaAsamblea As DateTime = DateTime.Parse(lector("fecha_asamblea").ToString())
                            Dim item As String = $"{lector("nombre_asamblea")} ({fechaAsamblea:dd/MM/yyyy})"
                            cboAsamblea.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al cargar asambleas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub tabControl_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' Actualizar título del botón Guardar según la pestaña activa
        If tabControl.SelectedTab.Text.Contains("Usuario") Then
            btnGuardar.Text = "💾 Registrar Usuario"
            btnGuardar.Visible = True
        ElseIf tabControl.SelectedTab.Text.Contains("Cuotas") Then
            btnGuardar.Text = "💾 Generar Cuotas"
            btnGuardar.Visible = True
        Else
            btnGuardar.Text = "💾 Guardar"
            btnGuardar.Visible = False ' Ocultar en la pestaña de gestión
        End If
    End Sub

End Class