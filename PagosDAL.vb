Imports System.Data.SQLite

Public Class PagosDAL
    Public Shared Function ObtenerUltimoSaldo(idApartamento As Integer) As Decimal
        Try
            Using conn As New SQLiteConnection(ConexionDB.ObtenerCadenaConexion())
                conn.Open()

                Dim query As String = "SELECT SaldoActual FROM Pagos WHERE IdApartamento = @id ORDER BY FechaPago DESC, IdPago DESC LIMIT 1"

                Using cmd As New SQLiteCommand(query, conn)
                    cmd.Parameters.AddWithValue("@id", idApartamento)

                    Dim resultado = cmd.ExecuteScalar()
                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        Return Convert.ToDecimal(resultado)
                    End If
                End Using
            End Using

            Return 0
        Catch ex As Exception
            Throw New Exception($"Error al obtener último saldo: {ex.Message}")
        End Try
    End Function

    Public Shared Function ObtenerMatriculaInmobiliaria(idApartamento As Integer) As String
        Try
            Using conn As New SQLiteConnection(ConexionDB.ObtenerCadenaConexion())
                conn.Open()

                Dim query As String = "SELECT MatriculaInmobiliaria FROM Apartamentos WHERE IdApartamento = @id"

                Using cmd As New SQLiteCommand(query, conn)
                    cmd.Parameters.AddWithValue("@id", idApartamento)

                    Dim resultado = cmd.ExecuteScalar()
                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        Return resultado.ToString()
                    End If
                End Using
            End Using

            Return $"APT{idApartamento}"
        Catch ex As Exception
            Return $"APT{idApartamento}"
        End Try
    End Function

    Public Shared Function ExisteNumeroRecibo(numeroRecibo As String) As Boolean
        Try
            Using conn As New SQLiteConnection(ConexionDB.ObtenerCadenaConexion())
                conn.Open()

                Dim query As String = "SELECT COUNT(*) FROM Pagos WHERE NumeroRecibo = @numero"

                Using cmd As New SQLiteCommand(query, conn)
                    cmd.Parameters.AddWithValue("@numero", numeroRecibo)

                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function InsertarPago(pago As PagoModel) As Boolean
        Try
            Using conn As New SQLiteConnection(ConexionDB.ObtenerCadenaConexion())
                conn.Open()

                Dim query As String = "INSERT INTO Pagos (IdApartamento, NumeroRecibo, FechaPago, SaldoAnterior, " &
                                    "PagoAdministracion, PagoIntereses, TotalPagado, SaldoActual, Observaciones, " &
                                    "MatriculaInmobiliaria, CuotaActual, FechaRegistro) " &
                                    "VALUES (@idApartamento, @numeroRecibo, @fechaPago, @saldoAnterior, " &
                                    "@pagoAdmin, @pagoInteres, @totalPagado, @saldoActual, @observaciones, " &
                                    "@matricula, @cuotaActual, @fechaRegistro)"

                Using cmd As New SQLiteCommand(query, conn)
                    cmd.Parameters.AddWithValue("@idApartamento", pago.IdApartamento)
                    cmd.Parameters.AddWithValue("@numeroRecibo", pago.NumeroRecibo)
                    cmd.Parameters.AddWithValue("@fechaPago", pago.FechaPago)
                    cmd.Parameters.AddWithValue("@saldoAnterior", pago.SaldoAnterior)
                    cmd.Parameters.AddWithValue("@pagoAdmin", pago.PagoAdministracion)
                    cmd.Parameters.AddWithValue("@pagoInteres", pago.PagoIntereses)
                    cmd.Parameters.AddWithValue("@totalPagado", pago.TotalPagado)
                    cmd.Parameters.AddWithValue("@saldoActual", pago.SaldoActual)
                    cmd.Parameters.AddWithValue("@observaciones", If(String.IsNullOrEmpty(pago.Observaciones), "", pago.Observaciones))
                    cmd.Parameters.AddWithValue("@matricula", pago.MatriculaInmobiliaria)
                    cmd.Parameters.AddWithValue("@cuotaActual", pago.CuotaActual)
                    cmd.Parameters.AddWithValue("@fechaRegistro", pago.FechaRegistro)

                    cmd.ExecuteNonQuery()
                    Return True
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception($"Error al insertar pago: {ex.Message}")
        End Try
    End Function

    Public Shared Function ObtenerPagosPorTorre(numeroTorre As Integer) As List(Of PagoModel)
        Dim pagos As New List(Of PagoModel)

        Try
            Using conn As New SQLiteConnection(ConexionDB.ObtenerCadenaConexion())
                conn.Open()

                Dim query As String = "SELECT p.*, a.NumeroApartamento " &
                                    "FROM Pagos p " &
                                    "INNER JOIN Apartamentos a ON p.IdApartamento = a.IdApartamento " &
                                    "WHERE a.Torre = @torre " &
                                    "ORDER BY p.FechaPago DESC, a.NumeroApartamento"

                Using cmd As New SQLiteCommand(query, conn)
                    cmd.Parameters.AddWithValue("@torre", numeroTorre)

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim pago As New PagoModel With {
                                .IdPago = Convert.ToInt32(reader("IdPago")),
                                .IdApartamento = Convert.ToInt32(reader("IdApartamento")),
                                .NumeroRecibo = reader("NumeroRecibo").ToString(),
                                .FechaPago = Convert.ToDateTime(reader("FechaPago")),
                                .SaldoAnterior = Convert.ToDecimal(reader("SaldoAnterior")),
                                .PagoAdministracion = Convert.ToDecimal(reader("PagoAdministracion")),
                                .PagoIntereses = Convert.ToDecimal(reader("PagoIntereses")),
                                .TotalPagado = Convert.ToDecimal(reader("TotalPagado")),
                                .SaldoActual = Convert.ToDecimal(reader("SaldoActual")),
                                .Observaciones = reader("Observaciones").ToString(),
                                .MatriculaInmobiliaria = reader("MatriculaInmobiliaria").ToString(),
                                .CuotaActual = Convert.ToDecimal(reader("CuotaActual")),
                                .FechaRegistro = Convert.ToDateTime(reader("FechaRegistro"))
                            }
                            pagos.Add(pago)
                        End While
                    End Using
                End Using
            End Using

            Return pagos
        Catch ex As Exception
            Throw New Exception($"Error al obtener pagos: {ex.Message}")
        End Try
    End Function
End Class