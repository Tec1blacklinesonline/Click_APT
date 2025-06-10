Imports System.Data.SQLite
Imports System.Configuration
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports BCrypt.Net

Public Class ConexionBD
    ' 🔄 Obtiene la cadena de conexión desde App.config
    Private Shared ReadOnly cadenaConexion As String = ConfigurationManager.ConnectionStrings("MiConexionSQLite").ConnectionString

    ' 🌐 Retorna un objeto SQLiteConnection
    Public Shared Function ObtenerConexion() As SQLiteConnection
        Return New SQLiteConnection(cadenaConexion)
    End Function

    ' 🔄 Prueba si hay conexión a la base de datos
    Public Shared Function ProbarConexion() As Boolean
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()

                ' NUEVO: Verificar si las tablas principales existen
                Dim tablasRequeridas As String() = {"Apartamentos", "pagos", "Usuarios", "Torres", "cuotas_generadas_apartamento"}
                For Each tabla In tablasRequeridas
                    If Not VerificarTablaExiste(conexion, tabla) Then
                        MessageBox.Show($"Error: La tabla '{tabla}' no existe en la base de datos.", "Error de Estructura", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                Next

                Return True
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al conectar con la base de datos: " & ex.Message, "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 🆕 NUEVO: Verifica si una tabla existe en la base de datos
    Private Shared Function VerificarTablaExiste(conexion As SQLiteConnection, nombreTabla As String) As Boolean
        Try
            Dim consulta As String = "SELECT name FROM sqlite_master WHERE type='table' AND name=@tabla"
            Using comando As New SQLiteCommand(consulta, conexion)
                comando.Parameters.AddWithValue("@tabla", nombreTabla)
                Dim resultado = comando.ExecuteScalar()
                Return resultado IsNot Nothing
            End Using
        Catch
            Return False
        End Try
    End Function

    ' 🔐 Genera el hash SHA256 de un texto (MANTENIDO para compatibilidad)
    Public Shared Function GenerarHashSHA256(texto As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(texto)
            Dim hash As Byte() = sha256.ComputeHash(bytes)
            Return BitConverter.ToString(hash).Replace("-", "").ToLower()
        End Using
    End Function

    ' 👤 MEJORADO: Valida el usuario con su contraseña usando BCrypt
    Public Shared Function ValidarUsuario(usuario As String, contrasena As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT contrasena_hash, activo FROM Usuarios WHERE nombre_usuario = @usuario"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@usuario", usuario)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            Dim hashGuardado As String = reader("contrasena_hash").ToString()
                            Dim usuarioActivo As Boolean = Convert.ToBoolean(reader("activo"))

                            ' Verificar si el usuario está activo
                            If Not usuarioActivo Then
                                MessageBox.Show("El usuario está desactivado. Contacte al administrador.", "Usuario Inactivo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                Return False
                            End If

                            ' Verificar contraseña con BCrypt
                            Dim esValida As Boolean = BCrypt.Net.BCrypt.Verify(contrasena, hashGuardado)

                            ' Si la validación es exitosa, actualizar último acceso
                            If esValida Then
                                ActualizarUltimoAcceso(usuario)
                            End If

                            Return esValida
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al validar usuario: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return False
    End Function

    ' 🆕 NUEVO: Actualiza la fecha de último acceso del usuario
    Private Shared Sub ActualizarUltimoAcceso(nombreUsuario As String)
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "UPDATE Usuarios SET ultimo_acceso = datetime('now') WHERE nombre_usuario = @usuario"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@usuario", nombreUsuario)
                    comando.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            ' No es crítico si falla la actualización del último acceso
            System.Diagnostics.Debug.WriteLine($"Error al actualizar último acceso: {ex.Message}")
        End Try
    End Sub

    ' 🆕 NUEVO: Obtiene información del usuario autenticado
    Public Shared Function ObtenerInformacionUsuario(nombreUsuario As String) As Object
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT id_usuario, nombre_usuario, nombre_completo, email, rol, fecha_creacion, ultimo_acceso FROM Usuarios WHERE nombre_usuario = @usuario AND activo = 1"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@usuario", nombreUsuario)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            Return New With {
                                .IdUsuario = Convert.ToInt32(reader("id_usuario")),
                                .NombreUsuario = reader("nombre_usuario").ToString(),
                                .NombreCompleto = reader("nombre_completo").ToString(),
                                .Email = reader("email").ToString(),
                                .Rol = reader("rol").ToString(),
                                .FechaCreacion = Convert.ToDateTime(reader("fecha_creacion")),
                                .UltimoAcceso = If(IsDBNull(reader("ultimo_acceso")), Nothing, Convert.ToDateTime(reader("ultimo_acceso")))
                            }
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener información del usuario: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    ' 🆕 NUEVO: Verifica si un usuario tiene permisos para una acción específica
    Public Shared Function VerificarPermisos(nombreUsuario As String, accionRequerida As String) As Boolean
        Try
            Dim infoUsuario = ObtenerInformacionUsuario(nombreUsuario)
            If infoUsuario IsNot Nothing Then
                Dim rol As String = infoUsuario.Rol.ToString()

                Select Case rol.ToUpper()
                    Case "ADMINISTRADOR"
                        Return True ' Administrador tiene todos los permisos

                    Case "OPERADOR"
                        ' Operador puede realizar la mayoría de acciones excepto gestión de usuarios
                        Return Not accionRequerida.ToUpper().Contains("USUARIO")

                    Case "CONSULTA"
                        ' Usuario de consulta solo puede ver información, no modificar
                        Return accionRequerida.ToUpper().Contains("VER") OrElse accionRequerida.ToUpper().Contains("CONSULTAR")

                    Case Else
                        Return False
                End Select
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al verificar permisos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return False
    End Function

    ' 🆕 NUEVO: Registra una actividad en el histórico de cambios
    Public Shared Sub RegistrarActividad(nombreUsuario As String, tablaAfectada As String, idRegistro As Integer, tipoCambio As String, detalle As String)
        Try
            HistoricoCambiosDAL.RegistrarCambio(tablaAfectada, idRegistro, tipoCambio, nombreUsuario, detalle)
        Catch ex As Exception
            ' No es crítico si falla el registro de actividad
            System.Diagnostics.Debug.WriteLine($"Error al registrar actividad: {ex.Message}")
        End Try
    End Sub

    ' 🆕 NUEVO: Verifica la integridad de la base de datos
    Public Shared Function VerificarIntegridadBD() As Boolean
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()

                ' Verificar integridad de SQLite
                Dim consulta As String = "PRAGMA integrity_check"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Dim resultado = comando.ExecuteScalar()
                    If resultado.ToString() = "ok" Then
                        Return True
                    Else
                        MessageBox.Show($"Error de integridad en la base de datos: {resultado}", "Error de Integridad", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al verificar integridad: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 🆕 NUEVO: Obtiene estadísticas generales de la base de datos
    Public Shared Function ObtenerEstadisticasGenerales() As Dictionary(Of String, Object)
        Dim estadisticas As New Dictionary(Of String, Object)

        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()

                ' Total de apartamentos
                Dim consultaApartamentos As String = "SELECT COUNT(*) FROM Apartamentos"
                Using comando As New SQLiteCommand(consultaApartamentos, conexion)
                    estadisticas("total_apartamentos") = Convert.ToInt32(comando.ExecuteScalar())
                End Using

                ' Total de usuarios activos
                Dim consultaUsuarios As String = "SELECT COUNT(*) FROM Usuarios WHERE activo = 1"
                Using comando As New SQLiteCommand(consultaUsuarios, conexion)
                    estadisticas("usuarios_activos") = Convert.ToInt32(comando.ExecuteScalar())
                End Using

                ' Total de pagos del mes actual
                Dim consultaPagos As String = "SELECT COUNT(*) FROM pagos WHERE strftime('%Y-%m', fecha_pago) = strftime('%Y-%m', 'now')"
                Using comando As New SQLiteCommand(consultaPagos, conexion)
                    estadisticas("pagos_mes_actual") = Convert.ToInt32(comando.ExecuteScalar())
                End Using

                ' Total recaudado del mes actual
                Dim consultaRecaudacion As String = "SELECT COALESCE(SUM(total_pagado), 0) FROM pagos WHERE strftime('%Y-%m', fecha_pago) = strftime('%Y-%m', 'now')"
                Using comando As New SQLiteCommand(consultaRecaudacion, conexion)
                    estadisticas("recaudacion_mes_actual") = Convert.ToDecimal(comando.ExecuteScalar())
                End Using

                ' Cuotas pendientes
                Dim consultaCuotasPendientes As String = "SELECT COUNT(*) FROM cuotas_generadas_apartamento WHERE estado = 'pendiente'"
                Using comando As New SQLiteCommand(consultaCuotasPendientes, conexion)
                    estadisticas("cuotas_pendientes") = Convert.ToInt32(comando.ExecuteScalar())
                End Using

            End Using
        Catch ex As Exception
            ' En caso de error, devolver valores por defecto
            estadisticas("total_apartamentos") = 0
            estadisticas("usuarios_activos") = 0
            estadisticas("pagos_mes_actual") = 0
            estadisticas("recaudacion_mes_actual") = 0
            estadisticas("cuotas_pendientes") = 0
        End Try

        Return estadisticas
    End Function

    ' 🆕 NUEVO: Realiza backup de la base de datos
    Public Shared Function RealizarBackup(rutaDestino As String) As Boolean
        Try
            Dim rutaOrigen As String = cadenaConexion.Split("="c)(1).Split(";"c)(0)
            System.IO.File.Copy(rutaOrigen, rutaDestino, True)
            Return True
        Catch ex As Exception
            MessageBox.Show($"Error al realizar backup: {ex.Message}", "Error de Backup", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 🆕 NUEVO: Variable para almacenar información del usuario actual de la sesión
    Private Shared usuarioActual As Object = Nothing

    ' 🆕 NUEVO: Establece el usuario actual de la sesión
    Public Shared Sub EstablecerUsuarioActual(nombreUsuario As String)
        usuarioActual = ObtenerInformacionUsuario(nombreUsuario)
    End Sub

    ' 🆕 NUEVO: Obtiene el usuario actual de la sesión
    Public Shared Function ObtenerUsuarioActual() As Object
        Return usuarioActual
    End Function

    ' 🆕 NUEVO: Limpia la sesión del usuario actual
    Public Shared Sub LimpiarSesion()
        usuarioActual = Nothing
    End Sub

End Class