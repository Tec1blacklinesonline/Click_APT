' ============================================================================
' CUOTASDAL.VB - VB.NET PURO SIMPLIFICADO
' ============================================================================

Imports System.Data.SQLite
Imports System.Windows.Forms

Public Class CuotasDAL

    ''' <summary>
    ''' Estructura para información de cuota pendiente
    ''' </summary>
    Public Structure CuotaPendienteInfo
        Public ExisteCuotaPendiente As Boolean
        Public FechaVencimiento As Date
        Public ValorCuota As Decimal
        Public IdCuota As Integer
        Public DiasVencida As Integer
    End Structure

    ''' <summary>
    ''' Obtiene la cuota pendiente más antigua de un apartamento
    ''' </summary>
    Public Shared Function ObtenerCuotaPendienteMasAntigua(idApartamento As Integer) As CuotaPendienteInfo
        Dim resultado As New CuotaPendienteInfo()
        resultado.ExisteCuotaPendiente = False

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT id_cuota, fecha_vencimiento, valor_cuota, fecha_cuota " &
                                       "FROM cuotas_generadas_apartamento " &
                                       "WHERE id_apartamentos = @idApartamento " &
                                       "AND estado = 'pendiente' " &
                                       "AND fecha_vencimiento < date('now') " &
                                       "ORDER BY fecha_vencimiento ASC " &
                                       "LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            resultado.ExisteCuotaPendiente = True
                            resultado.IdCuota = Convert.ToInt32(reader("id_cuota"))
                            resultado.FechaVencimiento = Convert.ToDateTime(reader("fecha_vencimiento"))
                            resultado.ValorCuota = Convert.ToDecimal(reader("valor_cuota"))
                            resultado.DiasVencida = (Date.Today - resultado.FechaVencimiento.Date).Days
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error en ObtenerCuotaPendienteMasAntigua: " & ex.Message)
        End Try

        Return resultado
    End Function

    ''' <summary>
    ''' Marca una cuota como pagada
    ''' </summary>
    Public Shared Function MarcarCuotaComoPagada(idCuota As Integer, idPago As Integer) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "UPDATE cuotas_generadas_apartamento " &
                                       "SET estado = 'pagada', " &
                                       "fecha_pago = datetime('now') " &
                                       "WHERE id_cuota = @idCuota"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idCuota", idCuota)
                    Return comando.ExecuteNonQuery() > 0
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al marcar cuota como pagada: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Obtiene el total de intereses calculados para un apartamento
    ''' </summary>
    Public Shared Function ObtenerTotalInteresesCalculados(idApartamento As Integer) As Decimal
        Dim totalIntereses As Decimal = 0

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT COALESCE(SUM(valor_interes), 0) " &
                                       "FROM calculos_interes " &
                                       "WHERE id_apartamentos = @idApartamento " &
                                       "AND pagado = 0"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Dim resultado = comando.ExecuteScalar()
                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        totalIntereses = Convert.ToDecimal(resultado)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al obtener intereses calculados: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return totalIntereses
    End Function

End Class