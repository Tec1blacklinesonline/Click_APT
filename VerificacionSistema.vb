' ============================================================================
' VERIFICACION AUTOMÁTICA DEL SISTEMA COOPDIASAM
' Este archivo verifica que todas las funcionalidades estén correctamente instaladas
' ============================================================================

Imports System.Data.SQLite
Imports System.IO
Imports System.Configuration
Imports System.Windows.Forms

Public Class VerificacionSistema

    ''' <summary>
    ''' Ejecuta verificación completa del sistema
    ''' </summary>
    Public Shared Function EjecutarVerificacionCompleta() As VerificacionResultado
        Dim resultado As New VerificacionResultado()

        Try
            ' 1. Verificar configuración básica
            VerificarConfiguracion(resultado)

            ' 2. Verificar base de datos
            VerificarBaseDatos(resultado)

            ' 3. Verificar archivos y carpetas
            VerificarArchivosYCarpetas(resultado)

            ' 4. Verificar librerías
            VerificarLibrerias(resultado)

            ' 5. Verificar funcionalidades clave
            VerificarFuncionalidades(resultado)

        Catch ex As Exception
            resultado.AgregarError($"Error general en verificación: {ex.Message}")
        End Try

        Return resultado
    End Function

    ''' <summary>
    ''' Verifica la configuración básica del App.config
    ''' </summary>
    Private Shared Sub VerificarConfiguracion(resultado As VerificacionResultado)
        Try
            ' Verificar cadena de conexión
            Dim cadenaConexion As String = ConfigurationManager.ConnectionStrings("MiConexionSQLite")?.ConnectionString
            If String.IsNullOrEmpty(cadenaConexion) Then
                resultado.AgregarError("❌ Cadena de conexión 'MiConexionSQLite' no encontrada en App.config")
            Else
                resultado.AgregarExito("✅ Cadena de conexión configurada correctamente")
            End If

            ' Verificar configuraciones clave
            Dim configuracionesClave As String() = {"NombreConjunto", "VersionAplicacion", "RutaRecibos", "RutaBackups"}
            For Each config In configuracionesClave
                Dim valor As String = ConfigurationManager.AppSettings(config)
                If String.IsNullOrEmpty(valor) Then
                    resultado.AgregarAdvertencia($"⚠️ Configuración '{config}' no encontrada")
                Else
                    resultado.AgregarExito($"✅ Configuración '{config}': {valor}")
                End If
            Next

        Catch ex As Exception
            resultado.AgregarError($"❌ Error al verificar configuración: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica la base de datos y su estructura
    ''' </summary>
    Private Shared Sub VerificarBaseDatos(resultado As VerificacionResultado)
        Try
            ' Verificar conexión
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                resultado.AgregarExito("✅ Conexión a base de datos exitosa")

                ' Verificar tablas principales
                Dim tablasRequeridas As String() = {
                    "Apartamentos", "pagos", "Usuarios", "Torres", "Asambleas",
                    "cuotas_generadas_apartamento", "parametros_interes", "historico_cambios",
                    "cuentas", "pagos_simplificados", "calculos_interes"
                }

                For Each tabla In tablasRequeridas
                    If VerificarTablaExiste(conexion, tabla) Then
                        ' Contar registros
                        Dim conteo As Integer = ContarRegistrosTabla(conexion, tabla)
                        resultado.AgregarExito($"✅ Tabla '{tabla}': {conteo} registros")
                    Else
                        resultado.AgregarError($"❌ Tabla '{tabla}' no existe")
                    End If
                Next

                ' Verificar integridad
                If ConexionBD.VerificarIntegridadBD() Then
                    resultado.AgregarExito("✅ Integridad de base de datos verificada")
                Else
                    resultado.AgregarError("❌ Problemas de integridad en base de datos")
                End If

                ' Verificar datos básicos
                VerificarDatosBásicos(conexion, resultado)

            End Using

        Catch ex As Exception
            resultado.AgregarError($"❌ Error al verificar base de datos: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica la existencia de archivos y carpetas necesarios
    ''' </summary>
    Private Shared Sub VerificarArchivosYCarpetas(resultado As VerificacionResultado)
        Try
            ' Verificar carpetas
            Dim carpetas As String() = {
                ConfigurationManager.AppSettings("RutaRecibos"),
                ConfigurationManager.AppSettings("RutaBackups"),
                ConfigurationManager.AppSettings("RutaLogs"),
                ConfigurationManager.AppSettings("RutaExportaciones")
            }

            For Each carpeta In carpetas
                If Not String.IsNullOrEmpty(carpeta) Then
                    If Directory.Exists(carpeta) Then
                        resultado.AgregarExito($"✅ Carpeta existe: {carpeta}")
                    Else
                        Try
                            Directory.CreateDirectory(carpeta)
                            resultado.AgregarAdvertencia($"⚠️ Carpeta creada: {carpeta}")
                        Catch
                            resultado.AgregarError($"❌ No se pudo crear carpeta: {carpeta}")
                        End Try
                    End If
                End If
            Next

            ' Verificar archivos de imagen
            Dim archivosImagen As String() = {
                ConfigurationManager.AppSettings("LogoConjunto"),
                ConfigurationManager.AppSettings("FirmaAdministrador")
            }

            For Each archivo In archivosImagen
                If Not String.IsNullOrEmpty(archivo) Then
                    If File.Exists(archivo) Then
                        resultado.AgregarExito($"✅ Archivo existe: {archivo}")
                    Else
                        resultado.AgregarAdvertencia($"⚠️ Archivo opcional no encontrado: {archivo}")
                    End If
                End If
            Next

        Catch ex As Exception
            resultado.AgregarError($"❌ Error al verificar archivos: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica que las librerías necesarias estén disponibles
    ''' </summary>
    Private Shared Sub VerificarLibrerias(resultado As VerificacionResultado)
        Try
            ' Verificar System.Data.SQLite
            Try
                Dim versionSQLite As String = GetType(SQLiteConnection).Assembly.GetName().Version.ToString()
                resultado.AgregarExito($"✅ System.Data.SQLite versión: {versionSQLite}")
            Catch
                resultado.AgregarError("❌ System.Data.SQLite no disponible")
            End Try

            ' Verificar BCrypt
            Try
                Dim hashTest As String = BCrypt.Net.BCrypt.HashPassword("test")
                resultado.AgregarExito("✅ BCrypt.Net disponible")
            Catch
                resultado.AgregarError("❌ BCrypt.Net no disponible")
            End Try

            ' Verificar iTextSharp
            Try
                Dim doc As New iTextSharp.text.Document()
                resultado.AgregarExito("✅ iTextSharp disponible")
            Catch
                resultado.AgregarError("❌ iTextSharp no disponible")
            End Try

        Catch ex As Exception
            resultado.AgregarError($"❌ Error al verificar librerías: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica funcionalidades clave del sistema
    ''' </summary>
    Private Shared Sub VerificarFuncionalidades(resultado As VerificacionResultado)
        Try
            ' Verificar clases DAL
            Try
                Dim apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()
                resultado.AgregarExito($"✅ ApartamentoDAL funcional: {apartamentos.Count} apartamentos")
            Catch ex As Exception
                resultado.AgregarError($"❌ Error en ApartamentoDAL: {ex.Message}")
            End Try

            Try
                Dim torres = TorresDAL.ObtenerTodasLasTorres()
                resultado.AgregarExito($"✅ TorresDAL funcional: {torres.Count} torres")
            Catch ex As Exception
                resultado.AgregarError($"❌ Error en TorresDAL: {ex.Message}")
            End Try

            Try
                Dim tasaInteres = ParametrosDAL.ObtenerTasaInteresMoraActual()
                resultado.AgregarExito($"✅ ParametrosDAL funcional: Tasa {tasaInteres}%")
            Catch ex As Exception
                resultado.AgregarError($"❌ Error en ParametrosDAL: {ex.Message}")
            End Try

            ' Verificar funciones de usuario
            Try
                Dim estadisticas = ConexionBD.ObtenerEstadisticasGenerales()
                resultado.AgregarExito("✅ Estadísticas generales funcionando")
            Catch ex As Exception
                resultado.AgregarError($"❌ Error en estadísticas: {ex.Message}")
            End Try

        Catch ex As Exception
            resultado.AgregarError($"❌ Error al verificar funcionalidades: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica datos básicos necesarios para el funcionamiento
    ''' </summary>
    Private Shared Sub VerificarDatosBásicos(conexion As SQLiteConnection, resultado As VerificacionResultado)
        Try
            ' Verificar que existe al menos un usuario administrador
            Dim consultaAdmin As String = "SELECT COUNT(*) FROM Usuarios WHERE rol = 'Administrador' AND activo = 1"
            Using comando As New SQLiteCommand(consultaAdmin, conexion)
                Dim adminCount As Integer = Convert.ToInt32(comando.ExecuteScalar())
                If adminCount > 0 Then
                    resultado.AgregarExito($"✅ {adminCount} usuario(s) administrador(es) activo(s)")
                Else
                    resultado.AgregarError("❌ No hay usuarios administradores activos")
                End If
            End Using

            ' Verificar parámetros de interés
            Dim consultaParametros As String = "SELECT COUNT(*) FROM parametros_interes WHERE activo = 1"
            Using comando As New SQLiteCommand(consultaParametros, conexion)
                Dim paramCount As Integer = Convert.ToInt32(comando.ExecuteScalar())
                If paramCount > 0 Then
                    resultado.AgregarExito($"✅ {paramCount} parámetro(s) de interés activo(s)")
                Else
                    resultado.AgregarAdvertencia("⚠️ No hay parámetros de interés activos")
                End If
            End Using

            ' Verificar torres
            Dim consultaTorres As String = "SELECT COUNT(*) FROM Torres"
            Using comando As New SQLiteCommand(consultaTorres, conexion)
                Dim torreCount As Integer = Convert.ToInt32(comando.ExecuteScalar())
                If torreCount >= 8 Then
                    resultado.AgregarExito($"✅ {torreCount} torres configuradas")
                Else
                    resultado.AgregarAdvertencia($"⚠️ Solo {torreCount} torres configuradas (esperadas: 8)")
                End If
            End Using

        Catch ex As Exception
            resultado.AgregarError($"❌ Error al verificar datos básicos: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica si una tabla existe
    ''' </summary>
    Private Shared Function VerificarTablaExiste(conexion As SQLiteConnection, nombreTabla As String) As Boolean
        Try
            Dim consulta As String = "SELECT name FROM sqlite_master WHERE type='table' AND name=@tabla"
            Using comando As New SQLiteCommand(consulta, conexion)
                comando.Parameters.AddWithValue("@tabla", nombreTabla)
                Return comando.ExecuteScalar() IsNot Nothing
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Cuenta registros en una tabla
    ''' </summary>
    Private Shared Function ContarRegistrosTabla(conexion As SQLiteConnection, nombreTabla As String) As Integer
        Try
            Dim consulta As String = $"SELECT COUNT(*) FROM [{nombreTabla}]"
            Using comando As New SQLiteCommand(consulta, conexion)
                Return Convert.ToInt32(comando.ExecuteScalar())
            End Using
        Catch
            Return -1
        End Try
    End Function

End Class

''' <summary>
''' Clase para almacenar resultados de la verificación
''' </summary>
Public Class VerificacionResultado
    Public Property Exitosos As New List(Of String)
    Public Property Errores As New List(Of String)
    Public Property Advertencias As New List(Of String)

    Public Sub AgregarExito(mensaje As String)
        Exitosos.Add(mensaje)
    End Sub

    Public Sub AgregarError(mensaje As String)
        Errores.Add(mensaje)
    End Sub

    Public Sub AgregarAdvertencia(mensaje As String)
        Advertencias.Add(mensaje)
    End Sub

    Public ReadOnly Property EsExitoso As Boolean
        Get
            Return Errores.Count = 0
        End Get
    End Property

    Public ReadOnly Property TieneAdvertencias As Boolean
        Get
            Return Advertencias.Count > 0
        End Get
    End Property

    Private Function ObtenerValorCeldaSeguro(valor As Object, valorPorDefecto As String) As String
        If valor Is Nothing OrElse IsDBNull(valor) Then
            Return valorPorDefecto
        End If
        Return valor.ToString()
    End Function

    Public Function ObtenerResumenCompleto() As String
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("=== REPORTE DE VERIFICACIÓN DEL SISTEMA COOPDIASAM ===")
        sb.AppendLine($"Fecha: {DateTime.Now:dd/MM/yyyy HH:mm:ss}")
        sb.AppendLine()

        If Exitosos.Count > 0 Then
            sb.AppendLine("🟢 VERIFICACIONES EXITOSAS:")
            For Each exitoso In Exitosos
                sb.AppendLine($"   {exitoso}")
            Next
            sb.AppendLine()
        End If

        If Advertencias.Count > 0 Then
            sb.AppendLine("🟡 ADVERTENCIAS:")
            For Each advertencia In Advertencias
                sb.AppendLine($"   {advertencia}")
            Next
            sb.AppendLine()
        End If

        If Errores.Count > 0 Then
            sb.AppendLine("🔴 ERRORES ENCONTRADOS:")
            For Each errorMsg In Errores
                sb.AppendLine($"   {errorMsg}")
            Next
            sb.AppendLine()
        End If

        sb.AppendLine("=== RESUMEN ===")
        sb.AppendLine($"✅ Exitosos: {Exitosos.Count}")
        sb.AppendLine($"⚠️ Advertencias: {Advertencias.Count}")
        sb.AppendLine($"❌ Errores: {Errores.Count}")
        sb.AppendLine($"📊 Estado general: {If(EsExitoso, "EXITOSO", "CON ERRORES")}")

        Return sb.ToString()
    End Function

End Class

''' <summary>
''' Formulario para mostrar resultados de verificación
''' </summary>
Public Class FormVerificacion
    Inherits Form

    Private txtResultados As TextBox
    Private btnCerrar As Button
    Private btnGuardarReporte As Button

    Public Sub New(resultado As VerificacionResultado)
        InitializeComponent()
        MostrarResultados(resultado)
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Verificación del Sistema COOPDIASAM"
        Me.Size = New Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' TextBox para resultados
        txtResultados = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .ReadOnly = True,
            .Font = New Font("Consolas", 9),
            .Dock = DockStyle.Fill,
            .BackColor = Color.Black,
            .ForeColor = Color.LightGreen
        }
        Me.Controls.Add(txtResultados)

        ' Panel de botones
        Dim panelBotones As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 60
        }

        btnCerrar = New Button With {
            .Text = "Cerrar",
            .Size = New Size(100, 35),
            .Location = New Point(680, 12),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnCerrar.FlatAppearance.BorderSize = 0
        AddHandler btnCerrar.Click, Sub() Me.Close()
        panelBotones.Controls.Add(btnCerrar)

        btnGuardarReporte = New Button With {
            .Text = "Guardar Reporte",
            .Size = New Size(120, 35),
            .Location = New Point(540, 12),
            .BackColor = Color.FromArgb(39, 174, 96),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnGuardarReporte.FlatAppearance.BorderSize = 0
        AddHandler btnGuardarReporte.Click, AddressOf GuardarReporte
        panelBotones.Controls.Add(btnGuardarReporte)

        Me.Controls.Add(panelBotones)
    End Sub

    Private Sub MostrarResultados(resultado As VerificacionResultado)
        txtResultados.Text = resultado.ObtenerResumenCompleto()
    End Sub

    Private Sub GuardarReporte()
        Try
            Dim saveDialog As New SaveFileDialog With {
                .Filter = "Archivo de Texto|*.txt",
                .Title = "Guardar Reporte de Verificación",
                .FileName = $"VerificacionSistema_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                File.WriteAllText(saveDialog.FileName, txtResultados.Text)
                MessageBox.Show("Reporte guardado exitosamente.", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al guardar reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class