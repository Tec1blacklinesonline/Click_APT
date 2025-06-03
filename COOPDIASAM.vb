Imports System.Windows.Forms
Imports System.Drawing

Public Class COOPDIASAM
    ' Variables a nivel de clase para acceder a los elementos
    Private panelMenu As Panel
    Private botonMenu As Button
    Private panelContenido As Panel
    Private labelTitulo As Label
    Private botonesMenuItems As New List(Of Button)()
    Private botonesTorres As New List(Of Button)()

    ' Colores personalizados para la interfaz
    Private colorPrimario As Color = Color.FromArgb(41, 128, 185)    ' Azul
    Private colorSecundario As Color = Color.FromArgb(52, 152, 219)  ' Azul claro
    Private colorFondo As Color = Color.FromArgb(236, 240, 241)      ' Gris muy claro
    Private colorMenu As Color = Color.FromArgb(44, 62, 80)          ' Azul oscuro
    Private colorBoton As Color = Color.FromArgb(52, 73, 94)         ' Gris azulado
    Private colorPagos As Color = Color.FromArgb(39, 174, 96)        ' Verde para pagos
    Private colorPagosOscuro As Color = Color.FromArgb(34, 139, 34)  ' Verde oscuro para pagos

    Private Sub COOPDIASAM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Ajustes de la ventana
        Me.Text = "CONJUNTO RESIDENCIAL COOPDIASAMA"
        Me.Size = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ControlBox = False
        Me.BackColor = colorFondo

        ' Crear el panel superior
        CrearPanelSuperior()

        ' Crear el panel lateral (menú desplegable)
        CrearPanelMenu()

        ' Crear panel de contenido principal
        CrearPanelContenido()

        ' Crear los botones de las torres en el panel de contenido (vista por defecto)
        CrearTorres()

        ' Manejar clics fuera del panel para cerrar el menú
        AddHandler Me.MouseDown, AddressOf Form_MouseDown
    End Sub

    Private Sub CrearPanelSuperior()
        ' Panel superior que contiene título y botón de menú
        Dim panelSuperior As New Panel With {
            .Size = New Size(Me.ClientSize.Width, 60),
            .Location = New Point(0, 0),
            .BackColor = colorPrimario,
            .Dock = DockStyle.Top
        }
        Me.Controls.Add(panelSuperior)

        ' Botón de menú hamburguesa
        botonMenu = New Button With {
            .Size = New Size(40, 40),
            .Location = New Point(10, 10),
            .Text = "≡",
            .Font = New Font("Arial", 16, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = colorPrimario,
            .ForeColor = Color.White
        }
        botonMenu.FlatAppearance.BorderSize = 0
        botonMenu.FlatAppearance.MouseOverBackColor = colorSecundario

        AddHandler botonMenu.Click, AddressOf ToggleMenu
        panelSuperior.Controls.Add(botonMenu)

        ' Título del sistema
        labelTitulo = New Label With {
            .Text = "ADMINISTRACIÓN COOPDIASAMA",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(70, 15)
        }
        panelSuperior.Controls.Add(labelTitulo)

        ' Información de usuario
        Dim labelUsuario As New Label With {
            .Text = "Usuario: Admin",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(panelSuperior.Width - 150, 20),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        panelSuperior.Controls.Add(labelUsuario)
    End Sub

    Private Sub CrearPanelMenu()
        ' Panel de menú lateral
        panelMenu = New Panel With {
            .Size = New Size(220, Me.ClientSize.Height - 60),
            .Location = New Point(0, 60),
            .BackColor = colorMenu,
            .Visible = False
        }
        Me.Controls.Add(panelMenu)

        ' Título del menú
        Dim lblMenuTitulo As New Label With {
            .Text = "MENÚ PRINCIPAL",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(50, 20)
        }
        panelMenu.Controls.Add(lblMenuTitulo)

        ' Botones del menú con íconos
        Dim botonesMenu() As String = {"TABLERO", "TORRES", "PROPIETARIOS", "PAGOS", "INFORMES", "CONFIGURACIÓN", "CERRAR SESIÓN"}
        Dim iconos() As String = {"📊", "🏢", "👥", "💰", "📋", "⚙️", "🚪"}

        For i = 0 To botonesMenu.Length - 1
            Dim btn As New Button With {
                .Text = iconos(i) & " " & botonesMenu(i),
                .Size = New Size(200, 50),
                .Location = New Point(10, 60 + i * 60),
                .BackColor = colorBoton,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 10),
                .ForeColor = Color.White,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(10, 0, 0, 0),
                .Tag = botonesMenu(i).ToLower()
            }
            btn.FlatAppearance.BorderSize = 0
            btn.FlatAppearance.MouseOverBackColor = colorSecundario

            AddHandler btn.Click, AddressOf BotonMenu_Click
            panelMenu.Controls.Add(btn)
            botonesMenuItems.Add(btn)
        Next
    End Sub

    Private Sub CrearPanelContenido()
        ' Panel principal que contiene el contenido
        panelContenido = New Panel With {
            .Location = New Point(0, 60),
            .Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 60),
            .BackColor = colorFondo,
            .Dock = DockStyle.Fill
        }
        Me.Controls.Add(panelContenido)
    End Sub

    Private Sub CrearTorres()
        ' Limpiar panel antes de crear torres
        panelContenido.Controls.Clear()

        ' Título de sección
        Dim lblSeccion As New Label With {
            .Text = "TORRES DEL CONJUNTO",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = colorMenu,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelContenido.Controls.Add(lblSeccion)

        ' Línea divisoria
        Dim lineaDivisoria As New Panel With {
            .BackColor = colorPrimario,
            .Size = New Size(panelContenido.Width - 40, 2),
            .Location = New Point(20, lblSeccion.Location.Y + 30)
        }
        panelContenido.Controls.Add(lineaDivisoria)

        ' Crear torres
        CrearTorresLayout("Ver Apartamentos", colorSecundario, colorPrimario, AddressOf Torre_Click)
    End Sub

    Private Sub CrearTorresLayout(textoBoton As String, colorTorre As Color, colorEncabezado As Color, eventoClick As EventHandler)
        Dim nombres() As String = {"Torre 1", "Torre 2", "Torre 3", "Torre 4", "Torre 5", "Torre 6", "Torre 7", "Torre 8"}
        Dim torresPorFila As Integer = 4
        Dim torresWidth As Integer = 200
        Dim torresHeight As Integer = 150
        Dim espacioHorizontal As Integer = 30
        Dim espacioVertical As Integer = 30

        ' Calcular posiciones
        Dim anchoTotalTorres As Integer = (torresWidth * torresPorFila) + (espacioHorizontal * (torresPorFila - 1))
        Dim xStart As Integer = Math.Max(20, (panelContenido.Width - anchoTotalTorres) \ 2)
        Dim yStart As Integer = 100

        For i As Integer = 0 To nombres.Length - 1
            Dim fila As Integer = i \ torresPorFila
            Dim columna As Integer = i Mod torresPorFila

            ' Panel contenedor para cada torre
            Dim panelTorre As New Panel With {
                .Size = New Size(torresWidth, torresHeight),
                .Location = New Point(xStart + columna * (torresWidth + espacioHorizontal),
                                    yStart + fila * (torresHeight + espacioVertical)),
                .BackColor = colorTorre,
                .Tag = i + 1
            }

            ' Etiqueta superior con el nombre de la torre
            Dim lblTorre As New Label With {
                .Text = nombres(i),
                .Font = New Font("Segoe UI", 12, FontStyle.Bold),
                .ForeColor = Color.White,
                .BackColor = colorEncabezado,
                .Size = New Size(torresWidth, 30),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Top
            }
            panelTorre.Controls.Add(lblTorre)

            ' Botón de acción
            Dim btn As New Button With {
                .Text = textoBoton,
                .Size = New Size(torresWidth - 40, 40),
                .Location = New Point((torresWidth - (torresWidth - 40)) \ 2, torresHeight - 60),
                .BackColor = colorBoton,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Tag = i + 1,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold)
            }
            btn.FlatAppearance.BorderSize = 0
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(44, 62, 80)

            AddHandler btn.Click, eventoClick
            panelTorre.Controls.Add(btn)

            ' Información de la torre
            Dim lblInfo As New Label With {
                .Text = "5 Pisos" & Environment.NewLine & "20 Apartamentos",
                .Font = New Font("Segoe UI", 9),
                .ForeColor = Color.White,
                .Location = New Point((torresWidth - 160) \ 2, 50),
                .Size = New Size(160, 40),
                .TextAlign = ContentAlignment.MiddleCenter
            }
            panelTorre.Controls.Add(lblInfo)

            panelContenido.Controls.Add(panelTorre)

            ' Efectos hover dinámicos
            Dim colorHover As Color = If(colorTorre = colorSecundario, colorPrimario, Color.FromArgb(46, 204, 113))
            AddHandler panelTorre.MouseEnter, Sub(sender, e) panelTorre.BackColor = colorHover
            AddHandler panelTorre.MouseLeave, Sub(sender, e) panelTorre.BackColor = colorTorre
        Next
    End Sub

    Private Sub ToggleMenu(sender As Object, e As EventArgs)
        ' Animación simple para el menú
        If panelMenu.Visible Then
            panelContenido.Left = 0
            panelMenu.Visible = False
        Else
            panelMenu.BringToFront()
            panelMenu.Visible = True
            panelContenido.Left = panelMenu.Width
        End If
    End Sub

    Private Sub BotonMenu_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)

        ' Destacar el botón seleccionado
        For Each btn As Button In botonesMenuItems
            btn.BackColor = colorBoton
        Next
        boton.BackColor = colorSecundario

        ' Manejar la opción seleccionada
        Select Case boton.Tag.ToString()
            Case "tablero"
                labelTitulo.Text = "TABLERO PRINCIPAL"
                MostrarTablero()

            Case "torres"
                labelTitulo.Text = "GESTIÓN DE TORRES"
                MostrarSeccionTorres()

            Case "propietarios"
                labelTitulo.Text = "GESTIÓN DE PROPIETARIOS"
                MostrarSeccionPropietarios()

            Case "pagos"
                labelTitulo.Text = "CONTROL DE PAGOS"
                MostrarSeccionPagos()

            Case "informes"
                labelTitulo.Text = "INFORMES Y REPORTES"
                MostrarSeccionInformes()

            Case "configuración"
                labelTitulo.Text = "CONFIGURACIÓN DEL SISTEMA"
                MostrarSeccionConfiguracion()

            Case "cerrar sesión"
                CerrarSesion()
        End Select

        ' Cerrar el menú después de seleccionar
        panelMenu.Visible = False
        panelContenido.Left = 0
    End Sub

    Private Sub MostrarTablero()
        ' Mostrar vista de torres por defecto (dashboard principal)
        CrearTorres()
    End Sub

    Private Sub MostrarSeccionTorres()
        ' Reutilizar la función de crear torres para gestión
        CrearTorres()
    End Sub

    Private Sub MostrarSeccionPropietarios()
        ' Limpiar panel
        panelContenido.Controls.Clear()

        ' Título de sección
        Dim lblSeccion As New Label With {
            .Text = "GESTIÓN DE PROPIETARIOS",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = colorMenu,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelContenido.Controls.Add(lblSeccion)

        ' Mensaje temporal
        Dim lblMensaje As New Label With {
            .Text = "Sección en desarrollo..." & Environment.NewLine & "Aquí se mostrará la gestión de propietarios",
            .Font = New Font("Segoe UI", 12),
            .ForeColor = Color.Gray,
            .Location = New Point(20, 80),
            .Size = New Size(400, 60)
        }
        panelContenido.Controls.Add(lblMensaje)
    End Sub

    Private Sub MostrarSeccionPagos()
        ' Limpiar el panel de contenido
        panelContenido.Controls.Clear()

        ' Título de sección
        Dim lblSeccion As New Label With {
            .Text = "SELECCIONE UNA TORRE PARA REGISTRAR PAGOS",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = colorMenu,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelContenido.Controls.Add(lblSeccion)

        ' Línea divisoria
        Dim lineaDivisoria As New Panel With {
            .BackColor = colorPagos,
            .Size = New Size(panelContenido.Width - 40, 2),
            .Location = New Point(20, lblSeccion.Location.Y + 30)
        }
        panelContenido.Controls.Add(lineaDivisoria)

        ' Crear torres para pagos con diseño verde
        CrearTorresLayout("💰 Registrar Pagos", colorPagos, colorPagosOscuro, AddressOf TorrePagos_Click)
    End Sub

    Private Sub MostrarSeccionInformes()
        ' Limpiar panel
        panelContenido.Controls.Clear()

        ' Título de sección
        Dim lblSeccion As New Label With {
            .Text = "INFORMES Y REPORTES",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = colorMenu,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelContenido.Controls.Add(lblSeccion)

        ' Mensaje temporal
        Dim lblMensaje As New Label With {
            .Text = "Sección en desarrollo..." & Environment.NewLine & "Aquí se mostrarán los informes y reportes",
            .Font = New Font("Segoe UI", 12),
            .ForeColor = Color.Gray,
            .Location = New Point(20, 80),
            .Size = New Size(400, 60)
        }
        panelContenido.Controls.Add(lblMensaje)
    End Sub

    Private Sub MostrarSeccionConfiguracion()
        ' Limpiar panel
        panelContenido.Controls.Clear()

        ' Título de sección
        Dim lblSeccion As New Label With {
            .Text = "CONFIGURACIÓN DEL SISTEMA",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = colorMenu,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelContenido.Controls.Add(lblSeccion)

        ' Mensaje temporal
        Dim lblMensaje As New Label With {
            .Text = "Sección en desarrollo..." & Environment.NewLine & "Aquí se mostrará la configuración del sistema",
            .Font = New Font("Segoe UI", 12),
            .ForeColor = Color.Gray,
            .Location = New Point(20, 80),
            .Size = New Size(400, 60)
        }
        panelContenido.Controls.Add(lblMensaje)
    End Sub

    Private Sub CerrarSesion()
        If MessageBox.Show("¿Desea cerrar la sesión?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                ' Ocultar formulario actual
                Me.Hide()

                ' Abrir formulario de login
                Dim formLogin As New Inicio()
                formLogin.Show()
            Catch ex As Exception
                MessageBox.Show($"Error al cerrar sesión: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End Try
        End If
    End Sub

    Private Sub Torre_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)
        Dim numeroTorre As Integer = CInt(boton.Tag)

        Try
            ' Abrir el formulario de apartamentos por torre
            Dim formApartamentos As New FormApartamentosTorre(numeroTorre)
            formApartamentos.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir la torre {numeroTorre}: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TorrePagos_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)
        Dim numeroTorre As Integer = CInt(boton.Tag)

        Try
            ' Abrir el formulario de pagos para la torre seleccionada
            Dim formPagos As New FormPagos(numeroTorre)
            formPagos.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir pagos de la torre {numeroTorre}: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs)
        ' Detectar clicks fuera del menú para cerrarlo
        If panelMenu.Visible AndAlso Not panelMenu.Bounds.Contains(PointToClient(Cursor.Position)) Then
            If Not botonMenu.Bounds.Contains(PointToClient(Cursor.Position)) Then
                panelMenu.Visible = False
                panelContenido.Left = 0
            End If
        End If
    End Sub

    Private Sub COOPDIASAM_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ' Manejar redimensionamiento si es necesario
        If panelContenido IsNot Nothing Then
            ' Actualizar elementos responsivos aquí si fuera necesario
        End If
    End Sub

    Private Sub COOPDIASAM_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ' Limpiar recursos si es necesario
        Application.Exit()
    End Sub

End Class