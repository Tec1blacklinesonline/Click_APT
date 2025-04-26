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

    Private Sub COOPDIASAM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Ajustes de la ventana
        Me.Text = "CONJUNTO RESIDENCIAL COOPDIASAMA"
        Me.Size = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle ' Esto ya lo tienes
        Me.MaximizeBox = False ' Esto ya lo tienes
        Me.MinimizeBox = False ' Deshabilita el botón minimizar
        Me.ControlBox = False ' Elimina TODA la barra de controles (min, max, cerrar)
        Me.BackColor = colorFondo
        ' Eliminar esta línea si no tienes un icono en los recursos
        ' Me.Icon = My.Resources.IconoApp

        ' Crear el panel superior
        CrearPanelSuperior()

        ' Crear el panel lateral (menú desplegable)
        CrearPanelMenu()

        ' Crear panel de contenido principal
        CrearPanelContenido()

        ' Crear los botones de las torres en el panel de contenido
        CrearTorres()

        ' Manejar clics fuera del panel para cerrar el menú con un enfoque más eficiente
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
        ' Configurar FlatAppearance después de crear el botón
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

        ' Agregar información de usuario (puedes personalizar esto)
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
        ' Panel de menú lateral mejorado
        panelMenu = New Panel With {
            .Size = New Size(220, Me.ClientSize.Height - 60),
            .Location = New Point(0, 60),
            .BackColor = colorMenu,
            .Visible = False
        }
        Me.Controls.Add(panelMenu)

        ' Agregar título al menú
        Dim lblMenuTitulo As New Label With {
            .Text = "MENÚ PRINCIPAL",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(50, 20)
        }
        panelMenu.Controls.Add(lblMenuTitulo)

        ' Botones del menú con íconos (puedes agregar íconos si tienes recursos)
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
            ' Configurar FlatAppearance después de crear el botón
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
    End Sub

    Private Sub CrearTorres()
        Dim nombres() As String = {"Torre 1", "Torre 2", "Torre 3", "Torre 4", "Torre 5", "Torre 6", "Torre 7", "Torre 8"}
        Dim torresPorFila As Integer = 4 ' Cambia este valor según cuántas torres quieras por fila
        Dim torresWidth As Integer = 200
        Dim torresHeight As Integer = 150
        Dim espacioHorizontal As Integer = 30 ' Espacio entre torres horizontalmente
        Dim espacioVertical As Integer = 30 ' Espacio entre torres verticalmente

        ' Calcular el ancho total que ocuparán todas las torres en una fila
        Dim anchoTotalTorres As Integer = (torresWidth * torresPorFila) + (espacioHorizontal * (torresPorFila - 1))

        ' Calcular la posición inicial X para centrar las torres horizontalmente
        Dim xStart As Integer = (panelContenido.Width - anchoTotalTorres) \ 2

        ' Posición inicial Y (después del título y la línea divisoria)
        Dim yStart As Integer = 100

        For i As Integer = 0 To nombres.Length - 1
            ' Calcular posición basada en el índice
            Dim fila As Integer = i \ torresPorFila
            Dim columna As Integer = i Mod torresPorFila

            ' Panel contenedor para cada torre
            Dim panelTorre As New Panel With {
            .Size = New Size(torresWidth, torresHeight),
            .Location = New Point(xStart + columna * (torresWidth + espacioHorizontal),
                                yStart + fila * (torresHeight + espacioVertical)),
            .BackColor = colorSecundario,
            .Tag = i + 1
        }

            ' Etiqueta superior con el nombre de la torre
            Dim lblTorre As New Label With {
            .Text = nombres(i),
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = colorPrimario,
            .Size = New Size(torresWidth, 30),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Top
        }
            panelTorre.Controls.Add(lblTorre)

            ' Crear botón dentro del panel - ahora centrado horizontalmente
            Dim btn As New Button With {
            .Text = "Ver Apartamentos",
            .Size = New Size(torresWidth - 40, 40),
            .Location = New Point((torresWidth - (torresWidth - 40)) \ 2, torresHeight - 60), ' Centrado horizontalmente
            .BackColor = colorBoton,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Tag = i + 1,
            .Font = New Font("Segoe UI", 9)
        }
            btn.FlatAppearance.BorderSize = 0
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(44, 62, 80)

            AddHandler btn.Click, AddressOf Torre_Click
            panelTorre.Controls.Add(btn)

            ' Etiqueta con información de la torre - ahora centrada
            Dim lblInfo As New Label With {
            .Text = "5 Pisos" & Environment.NewLine & "20 Apartamentos",
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.White,
            .Location = New Point((torresWidth - 160) \ 2, 50), ' Centrado horizontalmente
            .Size = New Size(160, 40),
            .TextAlign = ContentAlignment.MiddleCenter
        }
            panelTorre.Controls.Add(lblInfo)

            panelContenido.Controls.Add(panelTorre)
            botonesTorres.Add(btn)

            ' Efectos hover
            AddHandler panelTorre.MouseEnter, Sub(sender, e) panelTorre.BackColor = Color.FromArgb(41, 128, 185)
            AddHandler panelTorre.MouseLeave, Sub(sender, e) panelTorre.BackColor = colorSecundario
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
                MessageBox.Show("Sección Tablero en desarrollo", "Información")

            Case "torres"
                labelTitulo.Text = "GESTIÓN DE TORRES"
                MostrarSeccionTorres()

            Case "propietarios"
                labelTitulo.Text = "GESTIÓN DE PROPIETARIOS"
                MessageBox.Show("Sección Propietarios en desarrollo", "Información")

            Case "pagos"
                labelTitulo.Text = "CONTROL DE PAGOS"
                MessageBox.Show("Sección Pagos en desarrollo", "Información")

            Case "informes"
                labelTitulo.Text = "INFORMES Y REPORTES"
                MessageBox.Show("Sección Informes en desarrollo", "Información")

            Case "configuración"
                labelTitulo.Text = "CONFIGURACIÓN DEL SISTEMA"
                MessageBox.Show("Sección Configuración en desarrollo", "Información")

            Case "cerrar sesión"
                If MessageBox.Show("¿Desea cerrar la sesión?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    ' Cerrar el formulario actual
                    Me.Hide() ' Ocultar en lugar de cerrar para evitar que se cierre toda la aplicación

                    ' Abrir el formulario de login
                    Dim formLogin As New Inicio() ' Reemplaza "FormLogin" con el nombre de tu formulario de login
                    formLogin.Show()
                End If
        End Select

        ' Cerrar el menú después de seleccionar
        panelMenu.Visible = False
        panelContenido.Left = 0
    End Sub
    Private Sub MostrarSeccionTorres()
        ' Esta función ya está implementada por defecto (los botones de torres)
        ' Aquí podrías recargar los datos de las torres desde la base de datos
        MessageBox.Show("Visualizando todas las torres", "Torres")
    End Sub

    Private Sub Torre_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)
        Dim numeroTorre As Integer = CInt(boton.Tag)

        ' Aquí deberías abrir un nuevo formulario para mostrar los apartamentos de esta torre
        MessageBox.Show($"Abriendo detalles de la Torre {numeroTorre}" & Environment.NewLine &
                       "Pisos: 5" & Environment.NewLine &
                       "Apartamentos por piso: 4" & Environment.NewLine &
                       "Total apartamentos: 20",
                       $"Torre {numeroTorre}")

        ' Para implementar en el futuro:
        ' Dim formTorre As New FormDetalleTorre(numeroTorre)
        ' formTorre.ShowDialog()
    End Sub

    Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs)
        ' Método más eficiente para detectar clicks fuera del menú
        If panelMenu.Visible AndAlso Not panelMenu.Bounds.Contains(PointToClient(Cursor.Position)) Then
            If Not botonMenu.Bounds.Contains(PointToClient(Cursor.Position)) Then
                panelMenu.Visible = False
                panelContenido.Left = 0
            End If
        End If
    End Sub

    Private Sub COOPDIASAM_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ' Manejar el redimensionamiento del formulario si es necesario
        If panelContenido IsNot Nothing Then
            ' Actualizar anchura de la línea divisoria y otros controles si fuera necesario
        End If
    End Sub


    Private Sub COOPDIASAM_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ' Este método ya no es necesario con el nuevo enfoque
        ' Pero puedes dejarlo si tiene otra funcionalidad importante
    End Sub


End Class