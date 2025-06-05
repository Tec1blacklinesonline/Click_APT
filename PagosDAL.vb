Imports System.Data.SQLite

Public Class PagosDAL

    ' Método para registrar un nuevo pago
    Public Shared Function RegistrarPago(pago As PagoModel) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' Insertar en tabla pagos
                        Dim consultaInsert As String = "
                            INSERT INTO pagos (
                                id_apartamentos, 
                                matricula_inmobiliaria, 
                                id_cuota, 
                                fecha_pago, 
                                numero_recibo, 
                                saldo_anterior, 
                                vr_pagado_administracion, 
                                vr_pagado_intereses, 
                                cuota_actual, 
                                total_pagado, 
                                saldo_actual, 
                                detalle, 
                                observacion
                            ) VALUES (
                                @idApartamento, 
                                @matricula, 
                                @idCuota, 
                                @fechaPago, 
                                @numeroRecibo, 
                                @saldoAnterior, 
                                @pagoAdmin, 
                                @pagoInteres, 
                                @cuotaActual, 
                                @totalPagado, 
                                @saldoFinal, 
                                @detalle, 
                                @observaciones
                            )"

                        Using comando As New SQLiteCommand(consultaInsert, conexion, transaccion)
                            comando.Parameters.AddWithValue("@idApartamento", pago.IdApartamento)
                            comando.Parameters.AddWithValue("@matricula", pago.MatriculaInmobiliaria)
                            comando.Parameters.AddWithValue("@idCuota", If(pago.IdCuota.HasValue, CObj(pago.IdCuota.Value), DBNull.Value))
                            comando.Parameters.AddWithValue("@fechaPago", pago.FechaPago)
                            comando.Parameters.AddWithValue("@numeroRecibo", pago.NumeroRecibo)
                            comando.Parameters.AddWithValue("@saldoAnterior", pago.SaldoAnterior)
                            comando.Parameters.AddWithValue("@pagoAdmin", pago.PagoAdministracion)
                            comando.Parameters.AddWithValue("@pagoInteres", pago.PagoIntereses)
                            comando.Parameters.AddWithValue("@cuotaActual", pago.CuotaActual)
                            comando.Parameters.AddWithValue("@totalPagado", pago.TotalPagado)
                            comando.Parameters.AddWithValue("@saldoFinal", pago.SaldoActual)
                            comando.Parameters.AddWithValue("@detalle", If(String.IsNullOrEmpty(pago.Detalle), DBNull.Value, CObj(pago.Detalle)))
                            comando.Parameters.AddWithValue("@observaciones", If(String.IsNullOrEmpty(pago.Observaciones), DBNull.Value, CObj(pago.Observaciones)))

                            comando.ExecuteNonQuery()
                        End Using

                        transaccion.Commit()
                        Return True

                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw New Exception($"Error en la transacción: {ex.Message}")
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al registrar pago: {ex.Message}")
        End Try
    End Function

    ' Método para obtener el historial de pagos de un apartamento
    Public Shared Function ObtenerHistorialPagos(idApartamento As Integer) As List(Of PagoModel)
        Dim pagos As New List(Of PagoModel)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        id_pago,
                        id_apartamentos,
                        matricula_inmobiliaria,
                        id_cuota,
                        fecha_pago,
                        numero_recibo,
                        saldo_anterior,
                        vr_pagado_administracion,
                        vr_pagado_intereses,
                        cuota_actual,
                        total_pagado,
                        saldo_actual,
                        detalle,
                        observacion
                    FROM pagos 
                    WHERE id_apartamentos = @idApartamento
                    ORDER BY fecha_pago DESC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim pago As New PagoModel With {
                                .IdPago = Convert.ToInt32(reader("id_pago")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .IdCuota = If(IsDBNull(reader("id_cuota")), Nothing, Convert.ToInt32(reader("id_cuota"))),
                                .FechaPago = Convert.ToDateTime(reader("fecha_pago")),
                                .NumeroRecibo = If(IsDBNull(reader("numero_recibo")), "", reader("numero_recibo").ToString()),
                                .SaldoAnterior = If(IsDBNull(reader("saldo_anterior")), 0D, Convert.ToDecimal(reader("saldo_anterior"))),
                                .PagoAdministracion = If(IsDBNull(reader("vr_pagado_administracion")), 0D, Convert.ToDecimal(reader("vr_pagado_administracion"))),
                                .PagoIntereses = If(IsDBNull(reader("vr_pagado_intereses")), 0D, Convert.ToDecimal(reader("vr_pagado_intereses"))),
                                .CuotaActual = If(IsDBNull(reader("cuota_actual")), 0D, Convert.ToDecimal(reader("cuota_actual"))),
                                .TotalPagado = If(IsDBNull(reader("total_pagado")), 0D, Convert.ToDecimal(reader("total_pagado"))),
                                .SaldoActual = If(IsDBNull(reader("saldo_actual")), 0D, Convert.ToDecimal(reader("saldo_actual"))),
                                .Detalle = If(IsDBNull(reader("detalle")), "", reader("detalle").ToString()),
                                .Observaciones = If(IsDBNull(reader("observacion")), "", reader("observacion").ToString())
                            }

                            pagos.Add(pago)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al obtener historial de pagos: {ex.Message}")
        End Try

        Return pagos
    End Function

    ' Método para obtener el último saldo de un apartamento
    Public Shared Function ObtenerUltimoSaldo(idApartamento As Integer) As Decimal
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT saldo_actual 
                    FROM pagos 
                    WHERE id_apartamentos = @idApartamento 
                    ORDER BY fecha_pago DESC, id_pago DESC
                    LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Dim resultado = comando.ExecuteScalar()
                    Return If(resultado IsNot Nothing AndAlso Not IsDBNull(resultado), Convert.ToDecimal(resultado), 0)
                End Using
            End Using

        Catch ex As Exception
            Return 0
        End Try
    End Function

    ' Método para verificar si existe un número de recibo
    Public Shared Function ExisteNumeroRecibo(numeroRecibo As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT COUNT(*) FROM pagos WHERE numero_recibo = @numeroRecibo"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@numeroRecibo", numeroRecibo)

                    Dim count As Integer = Convert.ToInt32(comando.ExecuteScalar())
                    Return count > 0
                End Using
            End Using

        Catch ex As Exception
            Return False
        End Try
    End Function

    ' Método para obtener pagos por torre y periodo
    Public Shared Function ObtenerPagosPorTorre(torre As Integer, fechaInicio As DateTime?, fechaFin As DateTime?) As List(Of PagoModel)
        Dim pagos As New List(Of PagoModel)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        p.id_pago,
                        p.id_apartamentos,
                        p.matricula_inmobiliaria,
                        p.id_cuota,
                        p.fecha_pago,
                        p.numero_recibo,
                        p.saldo_anterior,
                        p.vr_pagado_administracion,
                        p.vr_pagado_intereses,
                        p.cuota_actual,
                        p.total_pagado,
                        p.saldo_actual,
                        p.detalle,
                        p.observacion,
                        a.numero_apartamento
                    FROM pagos p
                    INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos
                    WHERE a.id_torre = @torre"

                If fechaInicio.HasValue Then
                    consulta &= " AND p.fecha_pago >= @fechaInicio"
                End If

                If fechaFin.HasValue Then
                    consulta &= " AND p.fecha_pago <= @fechaFin"
                End If

                consulta &= " ORDER BY p.fecha_pago DESC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@torre", torre)

                    If fechaInicio.HasValue Then
                        comando.Parameters.AddWithValue("@fechaInicio", fechaInicio.Value)
                    End If

                    If fechaFin.HasValue Then
                        comando.Parameters.AddWithValue("@fechaFin", fechaFin.Value)
                    End If

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim pago As New PagoModel With {
                                .IdPago = Convert.ToInt32(reader("id_pago")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .IdCuota = If(IsDBNull(reader("id_cuota")), Nothing, Convert.ToInt32(reader("id_cuota"))),
                                .FechaPago = Convert.ToDateTime(reader("fecha_pago")),
                                .NumeroRecibo = If(IsDBNull(reader("numero_recibo")), "", reader("numero_recibo").ToString()),
                                .SaldoAnterior = If(IsDBNull(reader("saldo_anterior")), 0D, Convert.ToDecimal(reader("saldo_anterior"))),
                                .PagoAdministracion = If(IsDBNull(reader("vr_pagado_administracion")), 0D, Convert.ToDecimal(reader("vr_pagado_administracion"))),
                                .PagoIntereses = If(IsDBNull(reader("vr_pagado_intereses")), 0D, Convert.ToDecimal(reader("vr_pagado_intereses"))),
                                .CuotaActual = If(IsDBNull(reader("cuota_actual")), 0D, Convert.ToDecimal(reader("cuota_actual"))),
                                .TotalPagado = If(IsDBNull(reader("total_pagado")), 0D, Convert.ToDecimal(reader("total_pagado"))),
                                .SaldoActual = If(IsDBNull(reader("saldo_actual")), 0D, Convert.ToDecimal(reader("saldo_actual"))),
                                .Detalle = If(IsDBNull(reader("detalle")), "", reader("detalle").ToString()),
                                .Observaciones = If(IsDBNull(reader("observacion")), "", reader("observacion").ToString()),
                                .NumeroApartamento = If(IsDBNull(reader("numero_apartamento")), "", reader("numero_apartamento").ToString())
                            }

                            pagos.Add(pago)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al obtener pagos por torre: {ex.Message}")
        End Try

        Return pagos
    End Function

    ' Método para obtener estadísticas de pagos por torre
    Public Shared Function ObtenerEstadisticasPagosTorre(torre As Integer, mes As Integer, año As Integer) As Dictionary(Of String, Object)
        Dim estadisticas As New Dictionary(Of String, Object)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        COUNT(DISTINCT p.id_apartamentos) as apartamentos_pagaron,
                        COUNT(*) as total_pagos,
                        SUM(p.vr_pagado_administracion) as total_administracion,
                        SUM(p.vr_pagado_intereses) as total_intereses,
                        SUM(p.total_pagado) as total_recaudado,
                        AVG(p.total_pagado) as promedio_pago
                    FROM pagos p
                    INNER JOIN Apartamentos a ON p.id_apartamentos = a.id_apartamentos
                    WHERE a.id_torre = @torre 
                    AND strftime('%m', p.fecha_pago) = @mes 
                    AND strftime('%Y', p.fecha_pago) = @año"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@torre", torre)
                    comando.Parameters.AddWithValue("@mes", mes.ToString("00"))
                    comando.Parameters.AddWithValue("@año", año.ToString())

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("apartamentos_pagaron") = If(IsDBNull(reader("apartamentos_pagaron")), 0, Convert.ToInt32(reader("apartamentos_pagaron")))
                            estadisticas("total_pagos") = If(IsDBNull(reader("total_pagos")), 0, Convert.ToInt32(reader("total_pagos")))
                            estadisticas("total_administracion") = If(IsDBNull(reader("total_administracion")), 0D, Convert.ToDecimal(reader("total_administracion")))
                            estadisticas("total_intereses") = If(IsDBNull(reader("total_intereses")), 0D, Convert.ToDecimal(reader("total_intereses")))
                            estadisticas("total_recaudado") = If(IsDBNull(reader("total_recaudado")), 0D, Convert.ToDecimal(reader("total_recaudado")))
                            estadisticas("promedio_pago") = If(IsDBNull(reader("promedio_pago")), 0D, Convert.ToDecimal(reader("promedio_pago")))
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' En caso de error, devolver valores por defecto
            estadisticas("apartamentos_pagaron") = 0
            estadisticas("total_pagos") = 0
            estadisticas("total_administracion") = 0
            estadisticas("total_intereses") = 0
            estadisticas("total_recaudado") = 0
            estadisticas("promedio_pago") = 0
        End Try

        Return estadisticas
    End Function

    ' Método para obtener la matrícula inmobiliaria de un apartamento
    Public Shared Function ObtenerMatriculaInmobiliaria(idApartamento As Integer) As String
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT matricula_inmobiliaria FROM Apartamentos WHERE id_apartamentos = @id"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)

                    Dim resultado = comando.ExecuteScalar()
                    Return If(resultado IsNot Nothing AndAlso Not IsDBNull(resultado), resultado.ToString(), "185000")
                End Using
            End Using

        Catch ex As Exception
            Return "185000" ' Valor por defecto
        End Try
    End Function

End Class