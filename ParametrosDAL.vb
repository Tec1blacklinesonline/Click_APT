Imports System.Data.SQLite
Imports System.Windows.Forms ' Solo para demostración de errores, se recomienda un sistema de logging en producción.

Public Class ParametrosDAL

    ''' <summary>
    ''' Obtiene la tasa de interés de mora más reciente y activa de la tabla Parametros_Interes.
    ''' Esta tasa debe ser la "Tasa de Usura" o el "Interés Bancario Corriente" * 1.5.
    ''' </summary>
    ''' <returns>La tasa de interés anual como Decimal (ej. 25.0 para 25%) o 0 si no se encuentra un parámetro válido.</returns>
    Public Shared Function ObtenerTasaInteresMoraActual() As Decimal
        Dim tasaInteres As Decimal = 0D ' Valor por defecto si no se encuentra ninguna tasa

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Consulta para obtener la tasa de interés más reciente (fecha_vigencia_desde)
                ' que sea válida hoy.
                ' Asume que `tasa_interes_mora` en la DB ya está en porcentaje (ej. 25.0)
                Dim consulta As String = "SELECT tasa_interes_mora FROM Parametros_Interes " &
                                         "WHERE date(fecha_vigencia_desde) <= date('now') " &
                                         "AND (fecha_vigencia_hasta IS NULL OR date(fecha_vigencia_hasta) >= date('now')) " &
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

End Class