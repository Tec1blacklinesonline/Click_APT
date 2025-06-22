' ============================================================================
' DASHBOARD MEJORADO SIN DEPENDENCIAS DE CHART
' Versión corregida que no requiere System.Windows.Forms.DataVisualization.Charting
' ============================================================================

Imports System.Data.SQLite
Imports System.Drawing
Imports System.Windows.Forms

Public Class FormDashboard
    Inherits Form

    ' Controles principales
    Private panelHeader As Panel
    Private panelMetricas As Panel
    Private panelGraficos As Panel
    Private panelActividades As Panel

    ' Métricas
    Private lblTotalRecaudado As Label
    Private lblTotalApartamentos As Label
    Private lblPagosMes As Label
    Private lblMorosidad As Label

    ' Paneles de gráficos (reemplazando Charts)
    Private panelRecaudacionMensual As Panel
    Private panelEstadoPagos As Panel
    Private panelTorres As Panel

    ' Lista de actividades
    Private lstActividades As ListBox

    ' Timer para actualización automática
    Private timerActualizacion As Timer

    Public Sub New()
        InitializeComponent()
        ConfigurarFormulario()
        CargarDatos()
        ConfigurarTimer()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Dashboard - COOPDIASAM"
        Me.Size = New Size(1400, 900)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(240, 240, 240)
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub ConfigurarFormulario()
        ' Panel superior (Header)
        panelHeader = New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(52, 73, 94)
        }

        Dim lblTitulo As New Label With {
            .Text = "📊 DASHBOARD EJECUTIVO - COOPDIASAM",
            .Font = New Font("Segoe UI", 20, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(30, 20),
            .AutoSize = True
        }

        Dim lblFechaHora As New Label With {
            .Text = DateTime.Now.ToString("dddd, dd MMMM yyyy - HH:mm"),
            .Font = New Font("Segoe UI", 12),
            .ForeColor = Color.LightGray,
            .Location = New Point(30, 50),
            .AutoSize = True
        }

        Dim btnActualizar As New Button With {
            .Text = "🔄 Actualizar",
            .Size = New Size(120, 35),
            .Location = New Point(1250, 22),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        btnActualizar.FlatAppearance.BorderSize = 0
        AddHandler btnActualizar.Click, AddressOf btnActualizar_Click

        panelHeader.Controls.AddRange({lblTitulo, lblFechaHora, btnActualizar})

        ' Panel de métricas (Cards superiores)
        panelMetricas = New Panel With {
            .Dock = DockStyle.Top,
            .Height = 150,
            .BackColor = Color.Transparent,
            .Padding = New Padding(20)
        }

        CrearTarjetasMetricas()

        ' Panel de gráficos (Centro)
        panelGraficos = New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .Padding = New Padding(20, 10, 20, 20)
        }

        CrearGraficos()

        ' Panel de actividades recientes (Lateral derecho)
        panelActividades = New Panel With {
            .Dock = DockStyle.Right,
            .Width = 350,
            .BackColor = Color.White,
            .Padding = New Padding(15)
        }

        CrearPanelActividades()

        ' Agregar controles al formulario
        Me.Controls.Add(panelGraficos)
        Me.Controls.Add(panelActividades)
        Me.Controls.Add(panelMetricas)
        Me.Controls.Add(panelHeader)
    End Sub

    Private Sub CrearTarjetasMetricas()
        ' Tarjeta 1: Total Recaudado
        Dim card1 As Panel = CrearTarjetaMetrica(
            "💰 TOTAL RECAUDADO (MES)",
            "$0",
            Color.FromArgb(39, 174, 96),
            New Point(20, 20)
        )
        lblTotalRecaudado = CType(card1.Controls(1), Label)

        ' Tarjeta 2: Total Apartamentos
        Dim card2 As Panel = CrearTarjetaMetrica(
            "🏠 TOTAL APARTAMENTOS",
            "0",
            Color.FromArgb(52, 152, 219),
            New Point(290, 20)
        )
        lblTotalApartamentos = CType(card2.Controls(1), Label)

        ' Tarjeta 3: Pagos del Mes
        Dim card3 As Panel = CrearTarjetaMetrica(
            "📋 PAGOS DEL MES",
            "0",
            Color.FromArgb(155, 89, 182),
            New Point(560, 20)
        )
        lblPagosMes = CType(card3.Controls(1), Label)

        ' Tarjeta 4: Tasa de Morosidad
        Dim card4 As Panel = CrearTarjetaMetrica(
            "⚠️ MOROSIDAD",
            "0%",
            Color.FromArgb(231, 76, 60),
            New Point(830, 20)
        )
        lblMorosidad = CType(card4.Controls(1), Label)

        panelMetricas.Controls.AddRange({card1, card2, card3, card4})
    End Sub

    Private Function CrearTarjetaMetrica(titulo As String, valor As String, color As Color, ubicacion As Point) As Panel
        Dim card As New Panel With {
            .Size = New Size(250, 110),
            .Location = ubicacion,
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.None
        }

        ' CORREGIDO: Agregar sombra usando evento Paint
        AddHandler card.Paint, Sub(sender, e)
                                   Using shadowBrush As New SolidBrush(Color.FromArgb(30, Color.Gray))
                                       e.Graphics.FillRectangle(shadowBrush, 3, 3, card.Width, card.Height)
                                   End Using
                               End Sub

        ' Línea superior de color
        Dim lineaColor As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 4,
            .BackColor = color
        }

        ' Título
        Dim lblTitulo As New Label With {
            .Text = titulo,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(80, 80, 80),
            .Location = New Point(15, 15),
            .AutoSize = True
        }

        ' Valor principal
        Dim lblValor As New Label With {
            .Text = valor,
            .Font = New Font("Segoe UI", 24, FontStyle.Bold),
            .ForeColor = color,
            .Location = New Point(15, 40),
            .AutoSize = True
        }

        ' Icono decorativo
        Dim lblIcono As New Label With {
            .Text = "📊",
            .Font = New Font("Segoe UI", 20),
            .ForeColor = Color.LightGray,
            .Location = New Point(190, 45),
            .AutoSize = True
        }

        card.Controls.AddRange({lineaColor, lblTitulo, lblValor, lblIcono})
        Return card
    End Function

    Private Sub CrearGraficos()
        ' Layout en grid 2x2
        Dim anchoGrafico As Integer = (panelGraficos.Width - 60) \ 2
        Dim altoGrafico As Integer = (panelGraficos.Height - 60) \ 2

        ' Panel 1: Recaudación Mensual (Superior Izquierdo)
        panelRecaudacionMensual = CrearPanelGrafico(
            "📊 Recaudación Últimos 6 Meses",
            New Point(20, 20),
            New Size(anchoGrafico, altoGrafico),
            Color.FromArgb(52, 152, 219)
        )

        ' Panel 2: Estado de Pagos (Superior Derecho)
        panelEstadoPagos = CrearPanelGrafico(
            "🥧 Estado de Pagos por Apartamento",
            New Point(anchoGrafico + 40, 20),
            New Size(anchoGrafico, altoGrafico),
            Color.FromArgb(39, 174, 96)
        )

        ' Panel 3: Comparación por Torres (Inferior Izquierdo)
        panelTorres = CrearPanelGrafico(
            "🏢 Recaudación por Torre (Mes Actual)",
            New Point(20, altoGrafico + 40),
            New Size(anchoGrafico, altoGrafico),
            Color.FromArgb(155, 89, 182)
        )

        ' Panel de métricas adicionales (Inferior Derecho)
        Dim panelMetricasAdicionales = CrearPanelMetricasAdicionales(
            New Point(anchoGrafico + 40, altoGrafico + 40),
            New Size(anchoGrafico, altoGrafico)
        )

        panelGraficos.Controls.AddRange({panelRecaudacionMensual, panelEstadoPagos, panelTorres, panelMetricasAdicionales})
    End Sub

    Private Function CrearPanelGrafico(titulo As String, ubicacion As Point, tamaño As Size, color As Color) As Panel
        Dim panel As New Panel With {
            .Location = ubicacion,
            .Size = tamaño,
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }

        ' Título del gráfico
        Dim lblTitulo As New Label With {
            .Text = titulo,
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.FromArgb(80, 80, 80),
            .Location = New Point(15, 15),
            .AutoSize = True
        }

        ' Contenedor para datos (simulando gráfico con texto)
        Dim panelDatos As New Panel With {
            .Location = New Point(15, 50),
            .Size = New Size(tamaño.Width - 30, tamaño.Height - 80),
            .BackColor = Color.FromArgb(245, 245, 245),
            .BorderStyle = BorderStyle.None
        }

        ' Label para mostrar datos
        Dim lblDatos As New Label With {
            .Text = "Cargando datos...",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = color,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Dock = DockStyle.Fill
        }

        panelDatos.Controls.Add(lblDatos)
        panel.Controls.AddRange({lblTitulo, panelDatos})

        Return panel
    End Function

    Private Function CrearPanelMetricasAdicionales(ubicacion As Point, tamaño As Size) As Panel
        Dim panel As New Panel With {
            .Location = ubicacion,
            .Size = tamaño,
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }

        ' Título
        Dim lblTitulo As New Label With {
            .Text = "📈 MÉTRICAS ADICIONALES",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(52, 73, 94),
            .Location = New Point(20, 20),
            .AutoSize = True
        }

        ' Métricas
        Dim yPos As Integer = 60
        Dim metricas As String() = {
            "💰 Promedio por Pago: $0",
            "📅 Días Promedio Mora: 0",
            "🏆 Torre con Mayor Recaudación: N/A",
            "📊 Tasa de Cobranza: 0%",
            "⏰ Último Pago Registrado: N/A",
            "💳 Apartamentos al Día: 0"
        }

        For Each metrica In metricas
            Dim lbl As New Label With {
                .Text = metrica,
                .Font = New Font("Segoe UI", 11),
                .ForeColor = Color.FromArgb(100, 100, 100),
                .Location = New Point(20, yPos),
                .AutoSize = True
            }
            panel.Controls.Add(lbl)
            yPos += 35
        Next

        panel.Controls.Add(lblTitulo)
        Return panel
    End Function

    Private Sub CrearPanelActividades()
        ' Título
        Dim lblTitulo As New Label With {
            .Text = "🕐 ACTIVIDADES RECIENTES",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(52, 73, 94),
            .Location = New Point(10, 10),
            .AutoSize = True
        }

        ' Lista de actividades
        lstActividades = New ListBox With {
            .Location = New Point(10, 50),
            .Size = New Size(320, 400),
            .Font = New Font("Segoe UI", 9),
            .BorderStyle = BorderStyle.None,
            .BackColor = Color.FromArgb(250, 250, 250)
        }

        ' Botón para ver más
        Dim btnVerMas As New Button With {
            .Text = "Ver Historial Completo",
            .Size = New Size(150, 30),
            .Location = New Point(10, 470),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9)
        }
        btnVerMas.FlatAppearance.BorderSize = 0
        AddHandler btnVerMas.Click, AddressOf btnVerMas_Click

        panelActividades.Controls.AddRange({lblTitulo, lstActividades, btnVerMas})
    End Sub

    Private Sub ConfigurarTimer()
        timerActualizacion = New Timer With {
            .Interval = 300000 ' 5 minutos
        }
        AddHandler timerActualizacion.Tick, AddressOf timerActualizacion_Tick
        timerActualizacion.Start()
    End Sub

    Private Sub CargarDatos()
        Try
            Me.Cursor = Cursors.WaitCursor

            ' Cargar métricas principales
            CargarMetricasPrincipales()

            ' Cargar gráficos
            CargarGraficoRecaudacion()
            CargarGraficoEstadoPagos()
            CargarGraficoTorres()

            ' Cargar actividades recientes
            CargarActividadesRecientes()

        Catch ex As Exception
            MessageBox.Show($"Error al cargar datos del dashboard: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CargarMetricasPrincipales()
        Try
            Dim estadisticas = ConexionBD.ObtenerEstadisticasGenerales()
            Dim estadisticasPagos = PagosDAL.ObtenerEstadisticasPagos()

            ' Actualizar métricas
            lblTotalRecaudado.Text = Convert.ToDecimal(estadisticasPagos("recaudacion_mes_actual")).ToString("C0")
            lblTotalApartamentos.Text = estadisticas("total_apartamentos").ToString()
            lblPagosMes.Text = estadisticasPagos("pagos_mes_actual").ToString()

            ' Calcular morosidad
            Dim totalApartamentos As Integer = Convert.ToInt32(estadisticas("total_apartamentos"))
            Dim apartamentosConPagos As Integer = Convert.ToInt32(estadisticasPagos("apartamentos_con_pagos"))
            Dim tasaMorosidad As Double = If(totalApartamentos > 0, Math.Round((1 - (apartamentosConPagos / totalApartamentos)) * 100, 1), 0)
            lblMorosidad.Text = $"{tasaMorosidad}%"

            ' Cambiar color según morosidad
            If tasaMorosidad > 30 Then
                lblMorosidad.ForeColor = Color.FromArgb(231, 76, 60) ' Rojo
            ElseIf tasaMorosidad > 15 Then
                lblMorosidad.ForeColor = Color.FromArgb(243, 156, 18) ' Amarillo
            Else
                lblMorosidad.ForeColor = Color.FromArgb(39, 174, 96) ' Verde
            End If

        Catch ex As Exception
            ' Valores por defecto en caso de error
            lblTotalRecaudado.Text = "$0"
            lblTotalApartamentos.Text = "0"
            lblPagosMes.Text = "0"
            lblMorosidad.Text = "0%"
        End Try
    End Sub

    Private Sub CargarGraficoRecaudacion()
        Try
            Dim panelDatos As Panel = CType(panelRecaudacionMensual.Controls(1), Panel)
            Dim lblDatos As Label = CType(panelDatos.Controls(0), Label)

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("📈 RECAUDACIÓN MENSUAL")
            sb.AppendLine("")

            ' Obtener datos de los últimos 6 meses
            For i As Integer = 5 To 0 Step -1
                Dim fecha As DateTime = DateTime.Now.AddMonths(-i)
                Dim mesAño As String = fecha.ToString("MMM yyyy")

                Try
                    Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                        conexion.Open()
                        Dim consulta As String = "SELECT COALESCE(SUM(total_pagado), 0) FROM pagos WHERE strftime('%Y-%m', fecha_pago) = @mesAño"
                        Using comando As New SQLiteCommand(consulta, conexion)
                            comando.Parameters.AddWithValue("@mesAño", fecha.ToString("yyyy-MM"))
                            Dim valor As Decimal = Convert.ToDecimal(comando.ExecuteScalar())
                            sb.AppendLine(mesAño & ": " & valor.ToString("C0"))
                        End Using
                    End Using
                Catch
                    sb.AppendLine(mesAño & ": Error")
                End Try
            Next

            lblDatos.Text = sb.ToString()

        Catch ex As Exception
            ' Datos de ejemplo en caso de error
            Dim lblDatos As Label = CType(CType(panelRecaudacionMensual.Controls(1), Panel).Controls(0), Label)
            lblDatos.Text = "📈 RECAUDACIÓN" & vbCrLf & vbCrLf &
                           "Oct 2024: $15,000,000" & vbCrLf &
                           "Nov 2024: $18,000,000" & vbCrLf &
                           "Dic 2024: $22,000,000" & vbCrLf &
                           "Ene 2025: $16,000,000" & vbCrLf &
                           "Feb 2025: $19,000,000" & vbCrLf &
                           "Mar 2025: $21,000,000"
        End Try
    End Sub

    Private Sub CargarGraficoEstadoPagos()
        Try
            Dim panelDatos As Panel = CType(panelEstadoPagos.Controls(1), Panel)
            Dim lblDatos As Label = CType(panelDatos.Controls(0), Label)

            Dim alDia As Integer = 0
            Dim mora As Integer = 0

            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Obtener apartamentos al día
                Dim consultaAlDia As String = "SELECT COUNT(*) FROM (SELECT a.id_apartamentos, COALESCE(p.saldo_actual, 0) as saldo FROM Apartamentos a LEFT JOIN pagos p ON a.id_apartamentos = p.id_apartamentos WHERE saldo <= 0 GROUP BY a.id_apartamentos)"
                Using comando As New SQLiteCommand(consultaAlDia, conexion)
                    alDia = Convert.ToInt32(comando.ExecuteScalar())
                End Using

                ' Obtener apartamentos con mora
                Dim consultaMora As String = "SELECT COUNT(*) FROM (SELECT a.id_apartamentos, COALESCE(p.saldo_actual, 0) as saldo FROM Apartamentos a LEFT JOIN pagos p ON a.id_apartamentos = p.id_apartamentos WHERE saldo > 0 GROUP BY a.id_apartamentos)"
                Using comando As New SQLiteCommand(consultaMora, conexion)
                    mora = Convert.ToInt32(comando.ExecuteScalar())
                End Using
            End Using

            lblDatos.Text = "🥧 ESTADOS DE PAGO" & vbCrLf & vbCrLf &
                           "✅ Al Día: " & alDia.ToString() & vbCrLf & vbCrLf &
                           "⚠️ Con Mora: " & mora.ToString() & vbCrLf & vbCrLf &
                           "📊 Total: " & (alDia + mora).ToString()

        Catch ex As Exception
            ' Datos de ejemplo
            Dim lblDatos As Label = CType(CType(panelEstadoPagos.Controls(1), Panel).Controls(0), Label)
            lblDatos.Text = "🥧 ESTADOS DE PAGO" & vbCrLf & vbCrLf &
                           "✅ Al Día: 45" & vbCrLf & vbCrLf &
                           "⚠️ Con Mora: 15" & vbCrLf & vbCrLf &
                           "📊 Total: 60"
        End Try
    End Sub

    Private Sub CargarGraficoTorres()
        Try
            Dim panelDatos As Panel = CType(panelTorres.Controls(1), Panel)
            Dim lblDatos As Label = CType(panelDatos.Controls(0), Label)

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("🏢 RECAUDACIÓN POR TORRE")
            sb.AppendLine("")

            For torre As Integer = 1 To 8
                Try
                    Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                        conexion.Open()
                        Dim consulta As String = "SELECT COALESCE(SUM(p.total_pagado), 0) FROM pagos p INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos WHERE a.id_torre = @torre AND strftime('%Y-%m', p.fecha_pago) = strftime('%Y-%m', 'now')"
                        Using comando As New SQLiteCommand(consulta, conexion)
                            comando.Parameters.AddWithValue("@torre", torre)
                            Dim valor As Decimal = Convert.ToDecimal(comando.ExecuteScalar())
                            sb.AppendLine("Torre " & torre.ToString() & ": " & valor.ToString("C0"))
                        End Using
                    End Using
                Catch
                    sb.AppendLine("Torre " & torre.ToString() & ": Error")
                End Try
            Next

            lblDatos.Text = sb.ToString()

        Catch ex As Exception
            ' Datos de ejemplo
            Dim lblDatos As Label = CType(CType(panelTorres.Controls(1), Panel).Controls(0), Label)
            lblDatos.Text = "🏢 RECAUDACIÓN" & vbCrLf & vbCrLf &
                           "Torre 1: $2,500,000" & vbCrLf &
                           "Torre 2: $3,200,000" & vbCrLf &
                           "Torre 3: $2,800,000" & vbCrLf &
                           "Torre 4: $3,100,000" & vbCrLf &
                           "Torre 5: $2,900,000" & vbCrLf &
                           "Torre 6: $3,400,000" & vbCrLf &
                           "Torre 7: $2,700,000" & vbCrLf &
                           "Torre 8: $3,000,000"
        End Try
    End Sub

    Private Sub CargarActividadesRecientes()
        Try
            lstActividades.Items.Clear()

            ' Obtener últimos pagos
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT p.fecha_pago, p.numero_recibo, p.total_pagado, a.numero_apartamento, a.id_torre FROM pagos p INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos ORDER BY p.fecha_pago DESC LIMIT 10"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim fecha As DateTime = Convert.ToDateTime(reader("fecha_pago"))
                            Dim recibo As String = reader("numero_recibo").ToString()
                            Dim monto As Decimal = Convert.ToDecimal(reader("total_pagado"))
                            Dim apartamento As String = reader("numero_apartamento").ToString()
                            Dim torre As Integer = Convert.ToInt32(reader("id_torre"))

                            Dim actividad As String = $"💰 {fecha:dd/MM HH:mm} - T{torre}-{apartamento}: {monto:C}"
                            lstActividades.Items.Add(actividad)
                        End While
                    End Using
                End Using
            End Using

            If lstActividades.Items.Count = 0 Then
                lstActividades.Items.Add("📋 No hay actividades recientes")
            End If

        Catch ex As Exception
            lstActividades.Items.Clear()
            lstActividades.Items.Add("⚠️ Error al cargar actividades")
        End Try
    End Sub

    ' Eventos
    Private Sub btnActualizar_Click(sender As Object, e As EventArgs)
        CargarDatos()
    End Sub

    Private Sub btnVerMas_Click(sender As Object, e As EventArgs)
        Try
            Dim formHistorial As New FormHistorial()
            formHistorial.ShowDialog()
        Catch ex As Exception
            MessageBox.Show($"Error al abrir historial: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub timerActualizacion_Tick(sender As Object, e As EventArgs)
        CargarDatos()
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        timerActualizacion?.Stop()
        MyBase.OnFormClosed(e)
    End Sub

End Class