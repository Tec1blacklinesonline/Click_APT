Imports System.Data.SQLite
Imports System.Windows.Forms ' Solo para demostración de errores, se recomienda un sistema de logging en producción.

Public Class ParametrosDAL

    ''' <summary>
    ''' Obtiene la tasa de interés de mora más reciente y activa de la tabla parametros_interes.
    ''' Esta tasa debe ser la "Tasa de Usura" o el "Interés Bancario Corriente" * 1.5.
    ''' CORREGIDO: Nombre de tabla cambiado a minúsculas
    ''' </summary>
    ''' <returns>La tasa de interés anual como Decimal (ej. 25.0 para 25%) o 0 si no se encuentra un parámetro válido.</returns>
    Public Shared Function ObtenerTasaInteresMoraActual() As Decimal
        Dim tasaInteres As Decimal = 0D ' Valor por defecto si no se encuentra ninguna tasa

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' CORREGIDO: Consulta tabla parametros_interes (minúsculas)
                ' que sea válida hoy.
                ' Asume que `tasa_interes_mora` en la DB ya está en porcentaje (ej. 25.0)
                Dim consulta As String = "SELECT tasa_interes_mora FROM parametros_interes " &
                                         "WHERE date(fecha_vigencia_desde) <= date('now') " &
                                         "AND (fecha_vigencia_hasta IS NULL OR date(fecha_vigencia_hasta) >= date('now')) " &
                                         "AND activo = 1 " &
                                         "ORDER BY fecha_vigencia_desde DESC, id_parametro DESC LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    Dim resultado As Object = comando.ExecuteScalar()

                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        tasaInteres = Convert.ToDecimal(resultado)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al obtener la tasa de interés de mora de la base de datos: " & ex.Message, "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tasaInteres
    End Function

    ''' <summary>
    ''' NUEVO: Obtiene todos los parámetros de interés históricos
    ''' </summary>
    Public Shared Function ObtenerHistorialParametros() As List(Of Object)
        Dim parametros As New List(Of Object)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT * FROM parametros_interes ORDER BY fecha_vigencia_desde DESC"
                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            parametros.Add(New With {
                                .IdParametro = Convert.ToInt32(reader("id_parametro")),
                                .TasaInteresMora = Convert.ToDecimal(reader("tasa_interes_mora")),
                                .Descripcion = reader("descripcion").ToString(),
                                .FechaVigenciaDesde = Convert.ToDateTime(reader("fecha_vigencia_desde")),
                                .FechaVigenciaHasta = If(IsDBNull(reader("fecha_vigencia_hasta")), Nothing, Convert.ToDateTime(reader("fecha_vigencia_hasta"))),
                                .Activo = Convert.ToBoolean(reader("activo"))
                            })
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al obtener historial de parámetros: " & ex.Message, "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return parametros
    End Function

    ''' <summary>
    ''' NUEVO: Inserta un nuevo parámetro de interés
    ''' </summary>
    Public Shared Function InsertarParametroInteres(tasaInteres As Decimal, descripcion As String, fechaVigenciaDesde As DateTime, Optional fechaVigenciaHasta As DateTime? = Nothing) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' Desactivar parámetros anteriores si es necesario
                        Dim consultaDesactivar As String = "UPDATE parametros_interes SET activo = 0 WHERE activo = 1"
                        Using comandoDesactivar As New SQLiteCommand(consultaDesactivar, conexion, transaccion)
                            comandoDesactivar.ExecuteNonQuery()
                        End Using

                        ' Insertar nuevo parámetro
                        Dim consultaInsertar As String = "INSERT INTO parametros_interes (tasa_interes_mora, descripcion, fecha_vigencia_desde, fecha_vigencia_hasta, activo, fecha_actualizacion_ts) VALUES (@tasa, @descripcion, @fechaDesde, @fechaHasta, 1, datetime('now'))"
                        Using comandoInsertar As New SQLiteCommand(consultaInsertar, conexion, transaccion)
                            comandoInsertar.Parameters.AddWithValue("@tasa", tasaInteres)
                            comandoInsertar.Parameters.AddWithValue("@descripcion", descripcion)
                            comandoInsertar.Parameters.AddWithValue("@fechaDesde", fechaVigenciaDesde)
                            comandoInsertar.Parameters.AddWithValue("@fechaHasta", If(fechaVigenciaHasta.HasValue, CObj(fechaVigenciaHasta.Value), DBNull.Value))
                            comandoInsertar.ExecuteNonQuery()
                        End Using

                        transaccion.Commit()
                        Return True
                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al insertar parámetro de interés: " & ex.Message, "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

End Class