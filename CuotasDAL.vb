Imports System.Data.SQLite
Imports System.Windows.Forms ' Solo para demostración de errores, se recomienda un sistema de logging en producción.

Public Class CuotasDAL

    ''' <summary>
    ''' Clase interna para encapsular la información de la cuota pendiente más antigua.
    ''' </summary>
    Public Class CuotaPendienteInfo
        Public Property IdCuota As Integer
        Public Property ValorCuota As Decimal ' El valor de la cuota (base para el interés)
        Public Property FechaVencimiento As Date
        Public Property ExisteCuotaPendiente As Boolean = False
    End Class

    ''' <summary>
    ''' Obtiene la información de la cuota pendiente más antigua para un apartamento dado.
    ''' Una cuota se considera pendiente si su 'estado' es 'pendiente' y su fecha de vencimiento es pasada o hoy.
    ''' </summary>
    ''' <param name="idApartamento">El ID del apartamento.</param>
    ''' <returns>Un objeto CuotaPendienteInfo con los detalles de la cuota más antigua pendiente,
    ''' o un objeto con ExisteCuotaPendiente = False si no hay cuotas pendientes y vencidas.</returns>
    Public Shared Function ObtenerCuotaPendienteMasAntigua(idApartamento As Integer) As CuotaPendienteInfo
        Dim cuotaInfo As New CuotaPendienteInfo()

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Consulta para obtener la cuota más antigua pendiente (ordenando por fecha_vencimiento y luego por id_cuota)
                ' Se busca la cuota cuyo estado es 'pendiente' y que ya venció (o vence hoy).
                Dim consulta As String = "SELECT id_cuota, valor_cuota, fecha_vencimiento " &
                                         "FROM cuotas " &
                                         "WHERE id_apartamentos = @idApartamento " &
                                         "AND estado = 'pendiente' " &
                                         "AND date(fecha_vencimiento) <= date('now') " &
                                         "ORDER BY fecha_vencimiento ASC, id_cuota ASC LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            cuotaInfo.IdCuota = Convert.ToInt32(reader("id_cuota"))
                            cuotaInfo.ValorCuota = Convert.ToDecimal(reader("valor_cuota"))
                            cuotaInfo.FechaVencimiento = Convert.ToDateTime(reader("fecha_vencimiento"))
                            cuotaInfo.ExisteCuotaPendiente = True
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al obtener la cuota pendiente más antigua: " & ex.Message, "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cuotaInfo.ExisteCuotaPendiente = False ' Asegurar que el flag esté en falso si hay un error
        End Try

        Return cuotaInfo
    End Function

    ' Puedes añadir aquí más métodos para gestionar las cuotas (ej. registrar, actualizar estado)
    ' Public Shared Function RegistrarCuota(cuota As CuotaModel) As Boolean
    ' Public Shared Function ActualizarEstadoCuota(idCuota As Integer, estado As String) As Boolean
    ' ...

End Class