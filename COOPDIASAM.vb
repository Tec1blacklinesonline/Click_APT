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
    Private lblUsuarioActual As Label
    Private lblEstadisticas As Label

    ' Colores personalizados para la interfaz
    Private colorPrimario As Color = Color.FromArgb(41, 128, 185)    ' Azul
    Private colorSecundario As Color = Color.FromArgb(52, 152, 219)  ' Azul claro
    Private colorFondo As Color = Color.FromArgb(236, 240, 241)      ' Gris muy claro
    Private colorMenu As Color = Color.FromArgb(44, 62, 80)          ' Azul oscuro
    Private colorBoton As Color = Color.FromArgb(52, 73, 94)         ' Gris azulado
    Private colorPagos As Color = Color.FromArgb(39, 174, 96)        ' Verde para pagos
    Private colorPagosOscuro As Color = Color.FromArgb(34, 139, 34)  ' Verde oscuro para pagos
    Private colorEstados As Color = Color.FromArgb(155, 89, 182)     ' Morado para estados
    Private colorHistorial As Color = Color.FromArgb(231, 76, 60)    ' Rojo para historial

    Private Sub COOPDIASAM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Verificar integridad de la base de datos al inicio
        If Not ConexionBD.VerificarIntegridadBD() Then
            MessageBox.Show("Se detectaron problemas en la base de datos. El sistema puede no funcionar correctamente.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        ' Ajustes de la ventana
        Me.Text = "CONJUNTO RESIDENCIAL COOPDIASAM - v2025.1"
        Me.Size = New Size(1400, 800)  ' Ventana más grande
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
        CrearDashboardPrincipal()

        ' Manejar clics fuera del panel para cerrar el menú
        AddHandler Me.MouseDown, AddressOf Form_MouseDown

        ' Cargar estadísticas iniciales
        CargarEstadisticasGenerales()
    End Sub

    Private Sub CrearPanelSuperior()
        ' Panel superior que contiene título y botón de menú
        Dim panelSuperior As New Panel With {
            .Size = New Size(Me.ClientSize.Width, 80),  ' Más alto para más información
            .Location = New Point(0, 0),
            .BackColor = colorPrimario,
            .Dock = DockStyle.Top
        }
        Me.Controls.Add(panelSuperior)

        ' Botón de menú hamburguesa
        botonMenu = New Button With {
            .Size = New Size(40, 40),
            .Location = New Point(10, 20),
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
            .Text = "ADMINISTRACIÓN COOPDIASAM",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(70, 15)
        }
        panelSuperior.Controls.Add(labelTitulo)

        ' Información de usuario actual
        lblUsuarioActual = New Label With {
            .Text = "Cargando usuario...",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(panelSuperior.Width - 200, 15),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        panelSuperior.Controls.Add(lblUsuarioActual)

        ' Estadísticas generales
        lblEstadisticas = New Label With {
            .Text = "",
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.LightGray,
            .AutoSize = True,
            .Location = New Point(70, 45),
            .MaximumSize = New Size(800, 50)
        }
        panelSuperior.Controls.Add(lblEstadisticas)

        ' Botón de backup rápido
        Dim btnBackup As New Button With {
            .Text = "💾",
            .Size = New Size(30, 30),
            .Location = New Point(panelSuperior.Width - 50, 45),
            .BackColor = colorPrimario,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 12),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        btnBackup.FlatAppearance.BorderSize = 0
        AddHandler btnBackup.Click, AddressOf btnBackup_Click
        panelSuperior.Controls.Add(btnBackup)
    End Sub

    Private Sub CrearPanelMenu()
        ' Panel de menú lateral
        panelMenu = New Panel With {
            .Size = New Size(250, Me.ClientSize.Height - 80),  ' Más ancho
            .Location = New Point(0, 80),
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
            .Location = New Point(60, 20)
        }
        panelMenu.Controls.Add(lblMenuTitulo)

        ' Botones del menú con íconos (AMPLIADO)
        Dim botonesMenu() As String = {"DASHBOARD", "TORRES", "PROPIETARIOS", "PAGOS", "ESTADOS", "HISTORIAL", "REGISTRO", "CONFIGURACIÓN", "CERRAR SESIÓN"}
        Dim iconos() As String = {"📊", "🏢", "👥", "💰", "📋", "📜", "📝", "⚙️", "🚪"}

        For i = 0 To botonesMenu.Length - 1
            Dim btn As New Button With {
                .Text = iconos(i) & " " & botonesMenu(i),
                .Size = New Size(230, 45),  ' Botones más grandes
                .Location = New Point(10, 60 + i * 50),
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
            .Location = New Point(0, 80),
            .Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 80),
            .BackColor = colorFondo,
            .Dock = DockStyle.Fill
        }
        Me.Controls.Add(panelContenido)
    End Sub

    Private Sub CrearDashboardPrincipal()
        ' Limpiar panel antes de crear dashboard
        panelContenido.Controls.Clear()

        ' Título de sección
        Dim lblSeccion As New Label With {
            .Text = "DASHBOARD PRINCIPAL - CONJUNTO RESIDENCIAL COOPDIASAM",
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

        ' Panel de estadísticas resumidas
        CrearPanelEstadisticas()

        ' Panel de accesos rápidos
        CrearPanelAccesosRapidos()

        ' Panel de torres
        CrearTorresEnDashboard()
    End Sub

    Private Sub CrearPanelEstadisticas()
        Dim panelStats As New Panel With {
            .Location = New Point(20, 70),
            .Size = New Size(panelContenido.Width - 40, 120),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelContenido.Controls.Add(panelStats)

        ' Título del panel
        Dim lblTituloStats As New Label With {
            .Text = "📊 ESTADÍSTICAS GENERALES",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = colorMenu,
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        panelStats.Controls.Add(lblTituloStats)

        ' Labels para mostrar estadísticas
        Dim lblTotalApartamentos As New Label With {
            .Text = "Total Apartamentos: --",
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(20, 40),
            .AutoSize = True,
            .Tag = "total_apartamentos"
        }
        panelStats.Controls.Add(lblTotalApartamentos)

        Dim lblPagosMes As New Label With {
            .Text = "Pagos del Mes: --",
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(20, 60),
            .AutoSize = True,
            .Tag = "pagos_mes"
        }
        panelStats.Controls.Add(lblPagosMes)

        Dim lblRecaudacion As New Label With {
            .Text = "Recaudación del Mes: --",
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(300, 40),
            .AutoSize = True,
            .Tag = "recaudacion_mes"
        }
        panelStats.Controls.Add(lblRecaudacion)

        Dim lblCuotasPendientes As New Label With {
            .Text = "Cuotas Pendientes: --",
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(300, 60),
            .AutoSize = True,
            .Tag = "cuotas_pendientes"
        }
        panelStats.Controls.Add(lblCuotasPendientes)

        ' Botón actualizar estadísticas
        Dim btnActualizar As New Button With {
            .Text = "🔄 Actualizar",
            .Size = New Size(100, 30),
            .Location = New Point(panelStats.Width - 120, 80),
            .BackColor = colorPrimario,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnActualizar.FlatAppearance.BorderSize = 0
        AddHandler btnActualizar.Click, AddressOf CargarEstadisticasGenerales
        panelStats.Controls.Add(btnActualizar)
    End Sub

    Private Sub CrearPanelAccesosRapidos()
        Dim panelAccesos As New Panel With {
            .Location = New Point(20, 210),
            .Size = New Size(panelContenido.Width - 40, 80),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelContenido.Controls.Add(panelAccesos)

        Dim lblTituloAccesos As New Label With {
            .Text = "⚡ ACCESOS RÁPIDOS",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = colorMenu,
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        panelAccesos.Controls.Add(lblTituloAccesos)

        ' Botones de acceso rápido
        Dim accesosRapidos() As (String, Color, String) = {
            ("👥 Propietarios", colorSecundario, "propietarios"),
            ("📋 Estados", colorEstados, "estados"),
            ("📜 Historial", colorHistorial, "historial"),
            ("📝 Registro", colorPagosOscuro, "registro")
        }

        For i = 0 To accesosRapidos.Length - 1
            Dim btnAcceso As New Button With {
                .Text = accesosRapidos(i).Item1,
                .Size = New Size(150, 35),
                .Location = New Point(20 + i * 160, 35),
                .BackColor = accesosRapidos(i).Item2,
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold),
                .Tag = accesosRapidos(i).Item3
            }
            btnAcceso.FlatAppearance.BorderSize = 0
            AddHandler btnAcceso.Click, AddressOf AccesoRapido_Click
            panelAccesos.Controls.Add(btnAcceso)
        Next
    End Sub

    Private Sub CrearTorresEnDashboard()
        ' Título para torres
        Dim lblTorres As New Label With {
            .Text = "🏢 TORRES DEL CONJUNTO",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = colorMenu,
            .Location = New Point(20, 310),
            .AutoSize = True
        }
        panelContenido.Controls.Add(lblTorres)

        ' Crear torres con el layout original pero optimizado
        CrearTorresLayout("Ver Apartamentos", colorSecundario, colorPrimario, AddressOf Torre_Click, 330)
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
        CrearTorresLayout("Ver Apartamentos", colorSecundario, colorPrimario, AddressOf Torre_Click, 100)
    End Sub

    Private Sub CrearTorresLayout(textoBoton As String, colorTorre As Color, colorEncabezado As Color, eventoClick As EventHandler, yStart As Integer)
        Dim nombres() As String = {"Torre 1", "Torre 2", "Torre 3", "Torre 4", "Torre 5", "Torre 6", "Torre 7", "Torre 8"}
        Dim torresPorFila As Integer = 4
        Dim torresWidth As Integer = 200
        Dim torresHeight As Integer = 150
        Dim espacioHorizontal As Integer = 30
        Dim espacioVertical As Integer = 30

        ' Calcular posiciones
        Dim anchoTotalTorres As Integer = (torresWidth * torresPorFila) + (espacioHorizontal * (torresPorFila - 1))
        Dim xStart As Integer = Math.Max(20, (panelContenido.Width - anchoTotalTorres) \ 2)

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

            ' Información de la torre (con datos reales)
            Try
                Dim resumenTorre = ApartamentoDAL.ObtenerResumenTorre(i + 1)
                Dim totalApartamentos As Integer = Convert.ToInt32(resumenTorre("total_apartamentos"))
                Dim apartamentosAlDia As Integer = Convert.ToInt32(resumenTorre("apartamentos_al_dia"))
                Dim apartamentosPendientes As Integer = Convert.ToInt32(resumenTorre("apartamentos_pendientes"))

                Dim lblInfo As New Label With {
                    .Text = $"Total: {totalApartamentos}" & Environment.NewLine &
                           $"Al día: {apartamentosAlDia}" & Environment.NewLine &
                           $"Pendientes: {apartamentosPendientes}",
                    .Font = New Font("Segoe UI", 9),
                    .ForeColor = Color.White,
                    .Location = New Point(10, 40),
                    .Size = New Size(180, 60),
                    .TextAlign = ContentAlignment.TopLeft
                }
                panelTorre.Controls.Add(lblInfo)
            Catch
                Dim lblInfo As New Label With {
                    .Text = "5 Pisos" & Environment.NewLine & "20 Apartamentos",
                    .Font = New Font("Segoe UI", 9),
                    .ForeColor = Color.White,
                    .Location = New Point((torresWidth - 160) \ 2, 50),
                    .Size = New Size(160, 40),
                    .TextAlign = ContentAlignment.MiddleCenter
                }
                panelTorre.Controls.Add(lblInfo)
            End Try

            ' Botón de acción
            Dim btn As New Button With {
                .Text = textoBoton,
                .Size = New Size(torresWidth - 20, 35),
                .Location = New Point(10, torresHeight - 45),
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

            panelContenido.Controls.Add(panelTorre)

            ' Efectos hover dinámicos
            Dim colorHover As Color = If(colorTorre = colorSecundario, colorPrimario, Color.FromArgb(46, 204, 113))
            AddHandler panelTorre.MouseEnter, Sub(sender, e) panelTorre.BackColor = colorHover
            AddHandler panelTorre.MouseLeave, Sub(sender, e) panelTorre.BackColor = colorTorre
        Next
    End Sub

    Private Sub CargarEstadisticasGenerales()
        Try
            Dim estadisticas = ConexionBD.ObtenerEstadisticasGenerales()

            ' Actualizar labels en el panel de estadísticas si existe
            For Each control As Control In panelContenido.Controls
                If TypeOf control Is Panel Then
                    For Each subControl As Control In control.Controls
                        If TypeOf subControl Is Label AndAlso subControl.Tag IsNot Nothing Then
                            Select Case subControl.Tag.ToString()
                                Case "total_apartamentos"
                                    subControl.Text = $"Total Apartamentos: {estadisticas("total_apartamentos")}"
                                Case "pagos_mes"
                                    subControl.Text = $"Pagos del Mes: {estadisticas("pagos_mes_actual")}"
                                Case "recaudacion_mes"
                                    subControl.Text = $"Recaudación del Mes: {Convert.ToDecimal(estadisticas("recaudacion_mes_actual")):C}"
                                Case "cuotas_pendientes"
                                    subControl.Text = $"Cuotas Pendientes: {estadisticas("cuotas_pendientes")}"
                            End Select
                        End If
                    Next
                End If
            Next

            ' Actualizar estadísticas en el header
            lblEstadisticas.Text = $"📊 Apartamentos: {estadisticas("total_apartamentos")} | 💰 Pagos del mes: {estadisticas("pagos_mes_actual")} | 💵 Recaudado: {Convert.ToDecimal(estadisticas("recaudacion_mes_actual")):C}"

            ' Actualizar información del usuario
            Dim usuarioActual = ConexionBD.ObtenerUsuarioActual()
            If usuarioActual IsNot Nothing Then
                lblUsuarioActual.Text = $"👤 {usuarioActual.NombreCompleto} ({usuarioActual.Rol})"
            End If

        Catch ex As Exception
            lblEstadisticas.Text = "Error al cargar estadísticas"
        End Try
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
            Case "dashboard"
                labelTitulo.Text = "DASHBOARD PRINCIPAL"
                CrearDashboardPrincipal()

            Case "torres"
                labelTitulo.Text = "GESTIÓN DE TORRES"
                CrearTorres()

            Case "propietarios"
                labelTitulo.Text = "GESTIÓN DE PROPIETARIOS"
                MostrarSeccionPropietarios()

            Case "pagos"
                labelTitulo.Text = "CONTROL DE PAGOS"
                MostrarSeccionPagos()

            Case "estados"
                labelTitulo.Text = "ESTADOS DE CUENTA"
                MostrarSeccionEstados()

            Case "historial"
                labelTitulo.Text = "HISTORIAL Y AUDITORÍA"
                MostrarSeccionHistorial()

            Case "registro"
                labelTitulo.Text = "REGISTRO DE USUARIOS Y CUOTAS"
                MostrarSeccionRegistro()

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

    ' NUEVOS MÉTODOS PARA LAS NUEVAS SECCIONES
    Private Sub MostrarSeccionEstados()
        Try
            Dim formEstados As New FormEstados()
            formEstados.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir estados de cuenta: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MostrarSeccionHistorial()
        Try
            Dim formHistorial As New FormHistorial()
            formHistorial.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir historial: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MostrarSeccionRegistro()
        Try
            ' Verificar permisos
            Dim usuarioActual = ConexionBD.ObtenerUsuarioActual()
            If usuarioActual IsNot Nothing AndAlso usuarioActual.Rol.ToString() <> "Administrador" Then
                MessageBox.Show("No tiene permisos para acceder a esta sección.", "Acceso Denegado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim formRegistro As New FormRegistro()
            formRegistro.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir registro: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' MÉTODOS EXISTENTES MEJORADOS
    Private Sub MostrarSeccionPropietarios()
        Try
            Dim formPropietarios As New FormPropietarios()
            formPropietarios.ShowDialog()
            ' Actualizar estadísticas después de cerrar el formulario
            CargarEstadisticasGenerales()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir el formulario de propietarios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
        CrearTorresLayout("💰 Registrar Pagos", colorPagos, colorPagosOscuro, AddressOf TorrePagos_Click, 100)
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

        ' Panel de configuraciones
        Dim panelConfig As New Panel With {
            .Location = New Point(20, 70),
            .Size = New Size(panelContenido.Width - 40, 400),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        panelContenido.Controls.Add(panelConfig)

        ' Botón configuración SMTP
        Dim btnSMTP As New Button With {
            .Text = "📧 Configurar Correo SMTP",
            .Size = New Size(200, 40),
            .Location = New Point(20, 20),
            .BackColor = colorPrimario,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnSMTP.FlatAppearance.BorderSize = 0
        AddHandler btnSMTP.Click, AddressOf btnSMTP_Click
        panelConfig.Controls.Add(btnSMTP)

        ' Botón backup manual
        Dim btnBackupManual As New Button With {
            .Text = "💾 Realizar Backup",
            .Size = New Size(200, 40),
            .Location = New Point(240, 20),
            .BackColor = colorEstados,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnBackupManual.FlatAppearance.BorderSize = 0
        AddHandler btnBackupManual.Click, AddressOf btnBackupManual_Click
        panelConfig.Controls.Add(btnBackupManual)

        ' Botón verificar integridad
        Dim btnIntegridad As New Button With {
            .Text = "🔍 Verificar Integridad BD",
            .Size = New Size(200, 40),
            .Location = New Point(20, 80),
            .BackColor = colorHistorial,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnIntegridad.FlatAppearance.BorderSize = 0
        AddHandler btnIntegridad.Click, AddressOf btnIntegridad_Click
        panelConfig.Controls.Add(btnIntegridad)
    End Sub

    ' NUEVOS EVENTOS
    Private Sub AccesoRapido_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)
        Select Case boton.Tag.ToString()
            Case "propietarios"
                MostrarSeccionPropietarios()
            Case "estados"
                MostrarSeccionEstados()
            Case "historial"
                MostrarSeccionHistorial()
            Case "registro"
                MostrarSeccionRegistro()
        End Select
    End Sub

    Private Sub btnBackup_Click(sender As Object, e As EventArgs)
        RealizarBackupRapido()
    End Sub

    Private Sub btnSMTP_Click(sender As Object, e As EventArgs)
        Try
            Dim formSMTP As New FormConfiguracionSMTP()
            formSMTP.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir configuración SMTP: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBackupManual_Click(sender As Object, e As EventArgs)
        RealizarBackupRapido()
    End Sub

    Private Sub btnIntegridad_Click(sender As Object, e As EventArgs)
        Try
            Me.Cursor = Cursors.WaitCursor
            If ConexionBD.VerificarIntegridadBD() Then
                MessageBox.Show("La base de datos está íntegra y funcionando correctamente.", "Verificación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al verificar integridad: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub RealizarBackupRapido()
        Try
            Dim saveDialog As New SaveFileDialog With {
                .Filter = "Base de Datos|*.db",
                .Title = "Guardar Backup",
                .FileName = $"CONJUNTO_2025_backup_{DateTime.Now:yyyyMMdd_HHmmss}.db"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                If ConexionBD.RealizarBackup(saveDialog.FileName) Then
                    MessageBox.Show("Backup realizado exitosamente.", "Backup", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al realizar backup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CerrarSesion()
        If MessageBox.Show("¿Desea cerrar la sesión?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                ' Limpiar sesión actual
                ConexionBD.LimpiarSesion()

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

    ' EVENTOS EXISTENTES
    Private Sub Torre_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)
        Dim numeroTorre As Integer = CInt(boton.Tag)

        Try
            Dim formApartamentos As New FormApartamentosTorre(numeroTorre)
            formApartamentos.ShowDialog()
            ' Actualizar estadísticas después de cerrar el formulario
            CargarEstadisticasGenerales()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir la torre {numeroTorre}: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TorrePagos_Click(sender As Object, e As EventArgs)
        Dim boton As Button = CType(sender, Button)
        Dim numeroTorre As Integer = CInt(boton.Tag)

        Try
            Dim formPagos As New FormPagos(numeroTorre)
            formPagos.ShowDialog()
            ' Actualizar estadísticas después de cerrar el formulario
            CargarEstadisticasGenerales()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir pagos de la torre {numeroTorre}: {ex.Message}", "Error",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs)
        If panelMenu.Visible AndAlso Not panelMenu.Bounds.Contains(PointToClient(Cursor.Position)) Then
            If Not botonMenu.Bounds.Contains(PointToClient(Cursor.Position)) Then
                panelMenu.Visible = False
                panelContenido.Left = 0
            End If
        End If
    End Sub

    Private Sub COOPDIASAM_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If panelContenido IsNot Nothing Then
            ' Actualizar elementos responsivos aquí si fuera necesario
        End If
    End Sub

    Private Sub COOPDIASAM_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ' Limpiar recursos si es necesario
        ConexionBD.LimpiarSesion()
        Application.Exit()
    End Sub

End Class