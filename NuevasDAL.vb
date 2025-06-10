' ============================================================================
' NUEVAS CLASES DAL PARA MANEJAR LAS TABLAS AGREGADAS EN LA BD ACTUALIZADA
' ============================================================================

Imports System.Data.SQLite
Imports System.Windows.Forms

' ============================================================================
' CLASE PARA MANEJAR HISTORICO DE CAMBIOS
' ============================================================================
Public Class HistoricoCambiosDAL

    ''' <summary>
    ''' Registra un cambio en el histórico
    ''' </summary>
    Public Shared Function RegistrarCambio(tablaAfectada As String, idRegistro As Integer, tipoCambio As String, usuarioResponsable As String, detalleCambio As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "INSERT INTO historico_cambios (tabla_afectada, id_registro, tipo_cambio, fecha_cambio, usuario_responsable, detalle_cambio) VALUES (@tabla, @id, @tipo, datetime('now'), @usuario, @detalle)"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@tabla", tablaAfectada)
                    comando.Parameters.AddWithValue("@id", idRegistro)
                    comando.Parameters.AddWithValue("@tipo", tipoCambio)
                    comando.Parameters.AddWithValue("@usuario", usuarioResponsable)
                    comando.Parameters.AddWithValue("@detalle", detalleCambio)
                    Return comando.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al registrar cambio en histórico: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Obtiene el histórico de cambios para una tabla específica
    ''' </summary>
    Public Shared Function ObtenerHistoricoPorTabla(tablaAfectada As String) As List(Of Object)
        Dim cambios As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM historico_cambios WHERE tabla_afectada = @tabla ORDER BY fecha_cambio DESC"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@tabla", tablaAfectada)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            cambios.Add(New With {
                                .IdCambio = Convert.ToInt32(reader("id_cambio")),
                                .TablaAfectada = reader("tabla_afectada").ToString(),
                                .IdRegistro = Convert.ToInt32(reader("id_registro")),
                                .TipoCambio = reader("tipo_cambio").ToString(),
                                .FechaCambio = Convert.ToDateTime(reader("fecha_cambio")),
                                .UsuarioResponsable = reader("usuario_responsable").ToString(),
                                .DetalleCambio = reader("detalle_cambio").ToString()
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener histórico: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return cambios
    End Function

End Class

' ============================================================================
' CLASE PARA MANEJAR TORRES
' ============================================================================
Public Class TorresDAL

    ''' <summary>
    ''' Obtiene información de todas las torres
    ''' </summary>
    Public Shared Function ObtenerTodasLasTorres() As List(Of Object)
        Dim torres As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM Torres ORDER BY id_torre"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            torres.Add(New With {
                                .IdTorre = Convert.ToInt32(reader("id_torre")),
                                .Nombre = reader("nombre").ToString(),
                                .Pisos = Convert.ToInt32(reader("pisos")),
                                .TotalApartamentos = Convert.ToInt32(reader("total_apartamentos"))
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener torres: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return torres
    End Function

    ''' <summary>
    ''' Obtiene información de una torre específica
    ''' </summary>
    Public Shared Function ObtenerTorrePorId(idTorre As Integer) As Object
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM Torres WHERE id_torre = @id"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idTorre)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            Return New With {
                                .IdTorre = Convert.ToInt32(reader("id_torre")),
                                .Nombre = reader("nombre").ToString(),
                                .Pisos = Convert.ToInt32(reader("pisos")),
                                .TotalApartamentos = Convert.ToInt32(reader("total_apartamentos"))
                            }
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener torre: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

End Class

' ============================================================================
' CLASE PARA MANEJAR ASAMBLEAS
' ============================================================================
Public Class AsambleasDAL

    ''' <summary>
    ''' Obtiene todas las asambleas
    ''' </summary>
    Public Shared Function ObtenerTodasLasAsambleas() As List(Of Object)
        Dim asambleas As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM Asambleas ORDER BY fecha_asamblea DESC"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            asambleas.Add(New With {
                                .IdAsamblea = Convert.ToInt32(reader("id_asamblea")),
                                .NombreAsamblea = reader("nombre_asamblea").ToString(),
                                .FechaAsamblea = Convert.ToDateTime(reader("fecha_asamblea")),
                                .Descripcion = reader("descripcion").ToString()
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener asambleas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return asambleas
    End Function

    ''' <summary>
    ''' Crea una nueva asamblea
    ''' </summary>
    Public Shared Function CrearAsamblea(nombreAsamblea As String, fechaAsamblea As DateTime, descripcion As String) As Integer
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "INSERT INTO Asambleas (nombre_asamblea, fecha_asamblea, descripcion) VALUES (@nombre, @fecha, @descripcion); SELECT last_insert_rowid();"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@nombre", nombreAsamblea)
                    comando.Parameters.AddWithValue("@fecha", fechaAsamblea)
                    comando.Parameters.AddWithValue("@descripcion", descripcion)
                    Return Convert.ToInt32(comando.ExecuteScalar())
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al crear asamblea: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
    End Function

End Class

' ============================================================================
' CLASE PARA MANEJAR USUARIOS DEL SISTEMA
' ============================================================================
Public Class UsuariosDAL

    ''' <summary>
    ''' Obtiene todos los usuarios activos
    ''' </summary>
    Public Shared Function ObtenerUsuariosActivos() As List(Of Object)
        Dim usuarios As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT id_usuario, nombre_usuario, nombre_completo, email, rol, fecha_creacion, ultimo_acceso FROM Usuarios WHERE activo = 1 ORDER BY nombre_usuario"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            usuarios.Add(New With {
                                .IdUsuario = Convert.ToInt32(reader("id_usuario")),
                                .NombreUsuario = reader("nombre_usuario").ToString(),
                                .NombreCompleto = reader("nombre_completo").ToString(),
                                .Email = reader("email").ToString(),
                                .Rol = reader("rol").ToString(),
                                .FechaCreacion = Convert.ToDateTime(reader("fecha_creacion")),
                                .UltimoAcceso = If(IsDBNull(reader("ultimo_acceso")), Nothing, Convert.ToDateTime(reader("ultimo_acceso")))
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener usuarios: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return usuarios
    End Function

    ''' <summary>
    ''' Actualiza la fecha de último acceso de un usuario
    ''' </summary>
    Public Shared Function ActualizarUltimoAcceso(nombreUsuario As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "UPDATE Usuarios SET ultimo_acceso = datetime('now') WHERE nombre_usuario = @usuario"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@usuario", nombreUsuario)
                    Return comando.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al actualizar último acceso: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

End Class

' ============================================================================
' CLASE PARA MANEJAR CUENTAS POR APARTAMENTO
' ============================================================================
Public Class CuentasDAL

    ''' <summary>
    ''' Obtiene las cuentas de un apartamento
    ''' </summary>
    Public Shared Function ObtenerCuentasPorApartamento(idApartamento As Integer) As List(Of Object)
        Dim cuentas As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM cuentas WHERE id_apartamento = @id AND activo = 1 ORDER BY fecha_creacion DESC"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            cuentas.Add(New With {
                                .IdCuenta = Convert.ToInt32(reader("id_cuenta")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamento")),
                                .Concepto = reader("concepto").ToString(),
                                .Saldo = Convert.ToDecimal(reader("saldo")),
                                .FechaCreacion = Convert.ToDateTime(reader("fecha_creacion")),
                                .Activo = Convert.ToBoolean(reader("activo"))
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener cuentas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return cuentas
    End Function

    ''' <summary>
    ''' Crea una nueva cuenta para un apartamento
    ''' </summary>
    Public Shared Function CrearCuenta(idApartamento As Integer, concepto As String, saldo As Decimal) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "INSERT INTO cuentas (id_apartamento, concepto, saldo, fecha_creacion, activo) VALUES (@id, @concepto, @saldo, datetime('now'), 1)"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)
                    comando.Parameters.AddWithValue("@concepto", concepto)
                    comando.Parameters.AddWithValue("@saldo", saldo)
                    Return comando.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al crear cuenta: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

End Class

' ============================================================================
' CLASE PARA MANEJAR PAGOS SIMPLIFICADOS
' ============================================================================
Public Class PagosSimplificadosDAL

    ''' <summary>
    ''' Registra un pago simplificado
    ''' </summary>
    Public Shared Function RegistrarPagoSimplificado(idApartamento As Integer, monto As Decimal, concepto As String, metodoPago As String, Optional referencia As String = "", Optional observaciones As String = "") As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "INSERT INTO pagos_simplificados (id_apartamento, monto, fecha_pago, concepto, metodo_pago, referencia, observaciones, fecha_registro) VALUES (@id, @monto, datetime('now'), @concepto, @metodo, @referencia, @obs, datetime('now'))"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)
                    comando.Parameters.AddWithValue("@monto", monto)
                    comando.Parameters.AddWithValue("@concepto", concepto)
                    comando.Parameters.AddWithValue("@metodo", metodoPago)
                    comando.Parameters.AddWithValue("@referencia", referencia)
                    comando.Parameters.AddWithValue("@obs", observaciones)
                    Return comando.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al registrar pago simplificado: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Obtiene los pagos simplificados de un apartamento
    ''' </summary>
    Public Shared Function ObtenerPagosSimplificadosPorApartamento(idApartamento As Integer) As List(Of Object)
        Dim pagos As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM pagos_simplificados WHERE id_apartamento = @id ORDER BY fecha_pago DESC"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            pagos.Add(New With {
                                .IdPago = Convert.ToInt32(reader("id_pago")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamento")),
                                .Monto = Convert.ToDecimal(reader("monto")),
                                .FechaPago = Convert.ToDateTime(reader("fecha_pago")),
                                .Concepto = reader("concepto").ToString(),
                                .MetodoPago = reader("metodo_pago").ToString(),
                                .Referencia = reader("referencia").ToString(),
                                .Observaciones = reader("observaciones").ToString(),
                                .FechaRegistro = Convert.ToDateTime(reader("fecha_registro"))
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener pagos simplificados: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return pagos
    End Function

End Class