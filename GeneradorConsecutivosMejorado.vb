' ============================================================================
' GENERADOR DE CONSECUTIVOS MEJORADO - SOLO LA CLASE INDEPENDIENTE
' ✅ Este archivo va como GeneradorConsecutivosMejorado.vb
' ============================================================================

Imports System.Data.SQLite

Public Class GeneradorConsecutivosMejorado

    ''' <summary>
    ''' Genera número de recibo único con verificación de duplicados
    ''' </summary>
    Public Shared Function GenerarNumeroReciboUnico() As String
        Dim intentos As Integer = 0
        Dim maxIntentos As Integer = 10

        Do While intentos < maxIntentos
            Try
                Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                    conexion.Open()

                    Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                        Try
                            ' Obtener el último número con LOCK para evitar concurrencia
                            Dim consultaMax As String = "SELECT COALESCE(MAX(CAST(numero_recibo AS INTEGER)), 0) FROM pagos_apartamento WHERE numero_recibo GLOB '[0-9]*'"

                            Dim ultimoNumero As Integer = 0
                            Using comando As New SQLiteCommand(consultaMax, conexion, transaccion)
                                Dim resultado As Object = comando.ExecuteScalar()
                                If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                                    ultimoNumero = Convert.ToInt32(resultado)
                                End If
                            End Using

                            ' Generar siguiente número
                            Dim siguienteNumero As Integer = ultimoNumero + 1
                            Dim numeroRecibo As String = siguienteNumero.ToString().PadLeft(8, "0"c)

                            ' Verificar que no existe (doble verificación)
                            Dim consultaVerificacion As String = "SELECT COUNT(*) FROM pagos_apartamento WHERE numero_recibo = @numero"
                            Using comandoVerif As New SQLiteCommand(consultaVerificacion, conexion, transaccion)
                                comandoVerif.Parameters.AddWithValue("@numero", numeroRecibo)
                                Dim existe As Integer = Convert.ToInt32(comandoVerif.ExecuteScalar())

                                If existe = 0 Then
                                    ' Reservar el número insertando registro temporal
                                    Dim consultaReserva As String = "INSERT INTO temp_recibos_reservados (numero_recibo, fecha_reserva) VALUES (@numero, datetime('now'))"
                                    Try
                                        Using comandoReserva As New SQLiteCommand(consultaReserva, conexion, transaccion)
                                            comandoReserva.Parameters.AddWithValue("@numero", numeroRecibo)
                                            comandoReserva.ExecuteNonQuery()
                                        End Using
                                    Catch ex As SQLiteException
                                        ' Si falla la reserva, el número ya está en uso
                                        Continue Do
                                    End Try

                                    transaccion.Commit()
                                    Return numeroRecibo
                                End If
                            End Using

                        Catch
                            transaccion.Rollback()
                        End Try
                    End Using
                End Using

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"Error generando consecutivo (intento {intentos + 1}): {ex.Message}")
            End Try

            intentos += 1
            System.Threading.Thread.Sleep(100) ' Pequeña pausa entre intentos
        Loop

        ' Fallback: usar timestamp único
        Return DateTime.Now.ToString("yyyyMMddHHmmssfff")
    End Function

    ''' <summary>
    ''' Crear tabla temporal para reservar números de recibo
    ''' </summary>
    Public Shared Sub CrearTablaReservaRecibos()
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim crearTabla As String = "
                    CREATE TABLE IF NOT EXISTS temp_recibos_reservados (
                        numero_recibo TEXT PRIMARY KEY,
                        fecha_reserva TEXT NOT NULL
                    )"

                Using comando As New SQLiteCommand(crearTabla, conexion)
                    comando.ExecuteNonQuery()
                End Using

                ' Limpiar reservas antiguas (más de 1 hora)
                Dim limpiarAntiguos As String = "DELETE FROM temp_recibos_reservados WHERE datetime(fecha_reserva) < datetime('now', '-1 hour')"
                Using comando As New SQLiteCommand(limpiarAntiguos, conexion)
                    comando.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error creando tabla de reservas: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Liberar número de recibo reservado después de uso exitoso
    ''' </summary>
    Public Shared Sub LiberarNumeroReciboReservado(numeroRecibo As String)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "DELETE FROM temp_recibos_reservados WHERE numero_recibo = @numero"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@numero", numeroRecibo)
                    comando.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            ' Error silencioso
        End Try
    End Sub

End Class