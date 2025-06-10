' ============================================================================
' CLASE DAL COMPLETA PARA GESTIÓN DE PAGOS
' Implementa todos los métodos necesarios para el sistema de pagos
' ============================================================================

Imports System.Data.SQLite
Imports System.Windows.Forms

Public Class PagosDAL

    ''' <summary>
    ''' CRÍTICO: Registra un pago en la base de datos
    ''' Este método es llamado desde FormPagos.vb cuando se confirma un pago
    ''' </summary>
    Public Shared Function RegistrarPago(pago As PagoModel) As Boolean
        If pago Is Nothing OrElse Not pago.Validar() Then
            MessageBox.Show("Datos del pago inválidos", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' 1. Verificar que el número de recibo sea único
                        If ExisteNumeroRecibo(pago.NumeroRecibo, conexion, transaccion) Then
                            Throw New Exception($"El número de recibo {pago.NumeroRecibo} ya existe")
                        End If

                        ' 2. Insertar el pago
                        Dim consultaPago As String = "
                            INSERT INTO pagos (
                                id_apartamentos, id_cuota, fecha_pago, numero_recibo, 
                                saldo_anterior, vr_pagado_administracion, vr_pagado_intereses, 
                                cuota_actual, total_pagado, saldo_actual, detalle, 
                                observacion, estado_pago, registrado_por
                            ) VALUES (
                                @idApartamento, @idCuota, @fechaPago, @numeroRecibo,
                                @saldoAnterior, @pagoAdmin, @pagoIntereses,
                                @cuotaActual, @totalPagado, @saldoActual, @detalle,
                                @observaciones, @estadoPago, @registradoPor
                            )"

                        Using comandoPago As New SQLiteCommand(consultaPago, conexion, transaccion)
                            comandoPago.Parameters.AddWithValue("@idApartamento", pago.IdApartamento)
                            comandoPago.Parameters.AddWithValue("@idCuota", If(pago.IdCuota, DBNull.Value))
                            comandoPago.Parameters.AddWithValue("@fechaPago", pago.FechaPago)
                            comandoPago.Parameters.AddWithValue("@numeroRecibo", pago.NumeroRecibo)
                            comandoPago.Parameters.AddWithValue("@saldoAnterior", pago.SaldoAnterior)
                            comandoPago.Parameters.AddWithValue("@pagoAdmin", pago.PagoAdministracion)
                            comandoPago.Parameters.AddWithValue("@pagoIntereses", pago.PagoIntereses)
                            comandoPago.Parameters.AddWithValue("@cuotaActual", pago.CuotaActual)
                            comandoPago.Parameters.AddWithValue("@totalPagado", pago.TotalPagado)
                            comandoPago.Parameters.AddWithValue("@saldoActual", pago.SaldoActual)
                            comandoPago.Parameters.AddWithValue("@detalle", If(String.IsNullOrEmpty(pago.Detalle), DBNull.Value, pago.Detalle))
                            comandoPago.Parameters.AddWithValue("@observaciones", If(String.IsNullOrEmpty(pago.Observaciones), DBNull.Value, pago.Observaciones))
                            comandoPago.Parameters.AddWithValue("@estadoPago", pago.EstadoPago)
                            comandoPago.Parameters.AddWithValue("@registradoPor", If(String.IsNullOrEmpty(pago.UsuarioRegistro), DBNull.Value, pago.UsuarioRegistro))

                            comandoPago.ExecuteNonQuery()
                        End Using

                        ' 3. Si hay pago de administración, marcar cuotas como pagadas
                        If pago.PagoAdministracion > 0 Then
                            MarcarCuotasComoPagadas(pago.IdApartamento, pago.PagoAdministracion, conexion, transaccion)
                        End If

                        ' 4. Si hay pago de intereses, registrar en cálculos de interés
                        If pago.PagoIntereses > 0 Then
                            RegistrarPagoIntereses(pago.IdApartamento, pago.PagoIntereses, conexion, transaccion)
                        End If

                        ' 5. Registrar en histórico de cambios
                        ConexionBD.RegistrarActividad(
                            If(pago.UsuarioRegistro, "Sistema"),
                            "pagos",
                            pago.IdApartamento,
                            "INSERT",
                            "Pago registrado - Recibo: " & pago.NumeroRecibo & ", Total: " & pago.TotalPagado.ToString("C")
                        )

                        transaccion.Commit()
                        Return True

                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al registrar pago: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' CRÍTICO: Obtiene un pago por su número de recibo
    ''' Usado en FormPagos.vb para generar PDF y enviar correos
    ''' </summary>
    Public Shared Function ObtenerPagoPorNumeroRecibo(numeroRecibo As String) As PagoModel
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT * FROM pagos WHERE numero_recibo = @numeroRecibo"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@numeroRecibo", numeroRecibo)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            Dim pago As New PagoModel()
                            pago.IdPago = Convert.ToInt32(reader("id_pago"))
                            pago.IdApartamento = Convert.ToInt32(reader("id_apartamentos"))
                            pago.IdCuota = If(IsDBNull(reader("id_cuota")), Nothing, Convert.ToInt32(reader("id_cuota")))
                            pago.FechaPago = Convert.ToDateTime(reader("fecha_pago"))
                            pago.NumeroRecibo = reader("numero_recibo").ToString()
                            pago.SaldoAnterior = Convert.ToDecimal(reader("saldo_anterior"))
                            pago.PagoAdministracion = Convert.ToDecimal(reader("vr_pagado_administracion"))
                            pago.PagoIntereses = Convert.ToDecimal(reader("vr_pagado_intereses"))
                            pago.CuotaActual = Convert.ToDecimal(reader("cuota_actual"))
                            pago.TotalPagado = Convert.ToDecimal(reader("total_pagado"))
                            pago.SaldoActual = Convert.ToDecimal(reader("saldo_actual"))
                            pago.Detalle = If(IsDBNull(reader("detalle")), "", reader("detalle").ToString())
                            pago.Observaciones = If(IsDBNull(reader("observacion")), "", reader("observacion").ToString())
                            pago.EstadoPago = If(IsDBNull(reader("estado_pago")), "REGISTRADO", reader("estado_pago").ToString())

                            ' Obtener matrícula inmobiliaria para el pago
                            pago.MatriculaInmobiliaria = ObtenerMatriculaInmobiliaria(pago.IdApartamento)

                            Return pago
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al obtener pago: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' CRÍTICO: Obtiene el último saldo de un apartamento
    ''' Usado en FormPagos.vb para calcular el saldo anterior
    ''' </summary>
    Public Shared Function ObtenerUltimoSaldo(idApartamento As Integer) As Decimal
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Primero intentar obtener del último pago registrado
                Dim consultaUltimoPago As String = "
                    SELECT saldo_actual 
                    FROM pagos 
                    WHERE id_apartamentos = @idApartamento 
                    ORDER BY fecha_pago DESC, id_pago DESC 
                    LIMIT 1"

                Using comando As New SQLiteCommand(consultaUltimoPago, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Dim resultado = comando.ExecuteScalar()

                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        Return Convert.ToDecimal(resultado)
                    End If
                End Using

                ' Si no hay pagos, calcular saldo basado en cuotas pendientes
                Dim consultaCuotas As String = "
                    SELECT COALESCE(SUM(valor_cuota), 0) 
                    FROM cuotas_generadas_apartamento 
                    WHERE id_apartamentos = @idApartamento 
                      AND estado = 'pendiente'"

                Using comando As New SQLiteCommand(consultaCuotas, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Dim resultadoCuotas = comando.ExecuteScalar()
                    Return If(resultadoCuotas IsNot Nothing, Convert.ToDecimal(resultadoCuotas), 0D)
                End Using

            End Using

        Catch ex As Exception
            Return 0D
        End Try
    End Function

    ''' <summary>
    ''' CRÍTICO: Obtiene la matrícula inmobiliaria de un apartamento
    ''' Usado para generar el número de recibo según especificaciones
    ''' </summary>
    Public Shared Function ObtenerMatriculaInmobiliaria(idApartamento As Integer) As String
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT matricula_inmobiliaria FROM Apartamentos WHERE id_apartamentos = @idApartamento"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Dim resultado = comando.ExecuteScalar()

                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        Return resultado.ToString()
                    End If
                End Using
            End Using

        Catch ex As Exception
            Return ""
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' Obtiene el historial de pagos de un apartamento específico o de todos si idApartamento = 0
    ''' </summary>
    Public Shared Function ObtenerHistorialPagos(idApartamento As Integer) As List(Of PagoModel)
        Dim pagos As New List(Of PagoModel)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String
                If idApartamento > 0 Then
                    consulta = "SELECT * FROM pagos WHERE id_apartamentos = @idApartamento ORDER BY fecha_pago DESC"
                Else
                    consulta = "SELECT * FROM pagos ORDER BY fecha_pago DESC"
                End If

                Using comando As New SQLiteCommand(consulta, conexion)
                    If idApartamento > 0 Then
                        comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    End If

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim pago As New PagoModel()
                            pago.IdPago = Convert.ToInt32(reader("id_pago"))
                            pago.IdApartamento = Convert.ToInt32(reader("id_apartamentos"))
                            pago.IdCuota = If(IsDBNull(reader("id_cuota")), Nothing, Convert.ToInt32(reader("id_cuota")))
                            pago.FechaPago = Convert.ToDateTime(reader("fecha_pago"))
                            pago.NumeroRecibo = reader("numero_recibo").ToString()
                            pago.SaldoAnterior = Convert.ToDecimal(reader("saldo_anterior"))
                            pago.PagoAdministracion = Convert.ToDecimal(reader("vr_pagado_administracion"))
                            pago.PagoIntereses = Convert.ToDecimal(reader("vr_pagado_intereses"))
                            pago.CuotaActual = Convert.ToDecimal(reader("cuota_actual"))
                            pago.TotalPagado = Convert.ToDecimal(reader("total_pagado"))
                            pago.SaldoActual = Convert.ToDecimal(reader("saldo_actual"))
                            pago.Detalle = If(IsDBNull(reader("detalle")), String.Empty, reader("detalle").ToString())
                            pago.Observaciones = If(IsDBNull(reader("observacion")), String.Empty, reader("observacion").ToString())
                            pago.EstadoPago = If(IsDBNull(reader("estado_pago")), "REGISTRADO", reader("estado_pago").ToString())

                            pagos.Add(pago)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al obtener historial de pagos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return pagos
    End Function

    ''' <summary>
    ''' Obtiene el último pago realizado por un apartamento
    ''' </summary>
    Public Shared Function ObtenerUltimoPago(idApartamento As Integer) As PagoModel
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT * FROM pagos WHERE id_apartamentos = @idApartamento ORDER BY fecha_pago DESC LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            Dim pago As New PagoModel()
                            pago.IdPago = Convert.ToInt32(reader("id_pago"))
                            pago.IdApartamento = Convert.ToInt32(reader("id_apartamentos"))
                            pago.IdCuota = If(IsDBNull(reader("id_cuota")), Nothing, Convert.ToInt32(reader("id_cuota")))
                            pago.FechaPago = Convert.ToDateTime(reader("fecha_pago"))
                            pago.NumeroRecibo = reader("numero_recibo").ToString()
                            pago.SaldoAnterior = Convert.ToDecimal(reader("saldo_anterior"))
                            pago.PagoAdministracion = Convert.ToDecimal(reader("vr_pagado_administracion"))
                            pago.PagoIntereses = Convert.ToDecimal(reader("vr_pagado_intereses"))
                            pago.CuotaActual = Convert.ToDecimal(reader("cuota_actual"))
                            pago.TotalPagado = Convert.ToDecimal(reader("total_pagado"))
                            pago.SaldoActual = Convert.ToDecimal(reader("saldo_actual"))
                            pago.Detalle = If(IsDBNull(reader("detalle")), String.Empty, reader("detalle").ToString())
                            pago.Observaciones = If(IsDBNull(reader("observacion")), String.Empty, reader("observacion").ToString())
                            pago.EstadoPago = If(IsDBNull(reader("estado_pago")), "REGISTRADO", reader("estado_pago").ToString())

                            Return pago
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Log del error si es necesario
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Obtiene estadísticas de pagos para el dashboard
    ''' </summary>
    Public Shared Function ObtenerEstadisticasPagos() As Dictionary(Of String, Object)
        Dim estadisticas As New Dictionary(Of String, Object)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Estadísticas del mes actual
                Dim consultaMes As String = "
                    SELECT 
                        COUNT(*) as pagos_mes,
                        COALESCE(SUM(total_pagado), 0) as total_recaudado_mes,
                        COALESCE(SUM(vr_pagado_administracion), 0) as admin_recaudado_mes,
                        COALESCE(SUM(vr_pagado_intereses), 0) as intereses_recaudado_mes
                    FROM pagos 
                    WHERE strftime('%Y-%m', fecha_pago) = strftime('%Y-%m', 'now')"

                Using comando As New SQLiteCommand(consultaMes, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("pagos_mes_actual") = Convert.ToInt32(reader("pagos_mes"))
                            estadisticas("recaudacion_mes_actual") = Convert.ToDecimal(reader("total_recaudado_mes"))
                            estadisticas("admin_recaudado_mes") = Convert.ToDecimal(reader("admin_recaudado_mes"))
                            estadisticas("intereses_recaudado_mes") = Convert.ToDecimal(reader("intereses_recaudado_mes"))
                        End If
                    End Using
                End Using

                ' Estadísticas generales
                Dim consultaGeneral As String = "
                    SELECT 
                        COUNT(*) as total_pagos,
                        COALESCE(SUM(total_pagado), 0) as total_recaudado_historico,
                        COUNT(DISTINCT id_apartamentos) as apartamentos_con_pagos
                    FROM pagos"

                Using comando As New SQLiteCommand(consultaGeneral, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("total_pagos_historico") = Convert.ToInt32(reader("total_pagos"))
                            estadisticas("total_recaudado_historico") = Convert.ToDecimal(reader("total_recaudado_historico"))
                            estadisticas("apartamentos_con_pagos") = Convert.ToInt32(reader("apartamentos_con_pagos"))
                        End If
                    End Using
                End Using

                ' Último pago registrado
                Dim consultaUltimo As String = "
                    SELECT fecha_pago, numero_recibo, total_pagado 
                    FROM pagos 
                    ORDER BY fecha_pago DESC, id_pago DESC 
                    LIMIT 1"

                Using comando As New SQLiteCommand(consultaUltimo, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("ultimo_pago_fecha") = Convert.ToDateTime(reader("fecha_pago"))
                            estadisticas("ultimo_pago_recibo") = reader("numero_recibo").ToString()
                            estadisticas("ultimo_pago_valor") = Convert.ToDecimal(reader("total_pagado"))
                        End If
                    End Using
                End Using

            End Using

        Catch ex As Exception
            ' Valores por defecto en caso de error
            estadisticas("pagos_mes_actual") = 0
            estadisticas("recaudacion_mes_actual") = 0D
            estadisticas("admin_recaudado_mes") = 0D
            estadisticas("intereses_recaudado_mes") = 0D
            estadisticas("total_pagos_historico") = 0
            estadisticas("total_recaudado_historico") = 0D
            estadisticas("apartamentos_con_pagos") = 0
        End Try

        Return estadisticas
    End Function

    ''' <summary>
    ''' Anula un pago específico
    ''' </summary>
    Public Shared Function AnularPago(idPago As Integer, motivoAnulacion As String, usuarioAnula As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' Obtener información del pago antes de anular
                        Dim pago As PagoModel = Nothing
                        Dim consultaPago As String = "SELECT * FROM pagos WHERE id_pago = @idPago"

                        Using comando As New SQLiteCommand(consultaPago, conexion, transaccion)
                            comando.Parameters.AddWithValue("@idPago", idPago)
                            Using reader As SQLiteDataReader = comando.ExecuteReader()
                                If reader.Read() Then
                                    pago = New PagoModel()
                                    pago.IdPago = Convert.ToInt32(reader("id_pago"))
                                    pago.IdApartamento = Convert.ToInt32(reader("id_apartamentos"))
                                    pago.NumeroRecibo = reader("numero_recibo").ToString()
                                    pago.TotalPagado = Convert.ToDecimal(reader("total_pagado"))
                                    pago.PagoAdministracion = Convert.ToDecimal(reader("vr_pagado_administracion"))
                                    pago.PagoIntereses = Convert.ToDecimal(reader("vr_pagado_intereses"))
                                End If
                            End Using
                        End Using

                        If pago Is Nothing Then
                            Throw New Exception("Pago no encontrado")
                        End If

                        ' Marcar pago como anulado
                        Dim consultaAnular As String = "UPDATE pagos SET estado_pago = 'ANULADO', observacion = COALESCE(observacion, '') || ' | ANULADO: ' || @motivo || ' por ' || @usuario || ' el ' || datetime('now') WHERE id_pago = @idPago"

                        Using comando As New SQLiteCommand(consultaAnular, conexion, transaccion)
                            comando.Parameters.AddWithValue("@idPago", idPago)
                            comando.Parameters.AddWithValue("@motivo", motivoAnulacion)
                            comando.Parameters.AddWithValue("@usuario", usuarioAnula)
                            comando.ExecuteNonQuery()
                        End Using

                        ' Revertir cuotas marcadas como pagadas si las había
                        If pago.PagoAdministracion > 0 Then
                            RevertirCuotasPagadas(pago.IdApartamento, pago.PagoAdministracion, conexion, transaccion)
                        End If

                        ' Registrar en histórico
                        ConexionBD.RegistrarActividad(
                            usuarioAnula,
                            "pagos",
                            pago.IdApartamento,
                            "UPDATE",
                            "Pago anulado - Recibo: " & pago.NumeroRecibo & ", Motivo: " & motivoAnulacion
                        )

                        transaccion.Commit()
                        Return True

                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al anular pago: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' ============================================================================
    ' MÉTODOS AUXILIARES PRIVADOS
    ' ============================================================================

    ''' <summary>
    ''' Verifica si un número de recibo ya existe
    ''' </summary>
    Private Shared Function ExisteNumeroRecibo(numeroRecibo As String, conexion As SQLiteConnection, transaccion As SQLiteTransaction) As Boolean
        Dim consulta As String = "SELECT COUNT(*) FROM pagos WHERE numero_recibo = @numeroRecibo"
        Using comando As New SQLiteCommand(consulta, conexion, transaccion)
            comando.Parameters.AddWithValue("@numeroRecibo", numeroRecibo)
            Return Convert.ToInt32(comando.ExecuteScalar()) > 0
        End Using
    End Function

    ''' <summary>
    ''' Marca cuotas como pagadas según el monto pagado de administración
    ''' </summary>
    Private Shared Sub MarcarCuotasComoPagadas(idApartamento As Integer, montoPagado As Decimal, conexion As SQLiteConnection, transaccion As SQLiteTransaction)
        ' Obtener cuotas pendientes ordenadas por fecha de vencimiento
        Dim consulta As String = "
            SELECT id_cuota, valor_cuota 
            FROM cuotas_generadas_apartamento 
            WHERE id_apartamentos = @idApartamento 
              AND estado = 'pendiente' 
            ORDER BY fecha_vencimiento ASC"

        Dim montoRestante As Decimal = montoPagado

        Using comando As New SQLiteCommand(consulta, conexion, transaccion)
            comando.Parameters.AddWithValue("@idApartamento", idApartamento)
            Using reader As SQLiteDataReader = comando.ExecuteReader()
                Dim cuotasAPagar As New List(Of Integer)

                While reader.Read() AndAlso montoRestante > 0
                    Dim idCuota As Integer = Convert.ToInt32(reader("id_cuota"))
                    Dim valorCuota As Decimal = Convert.ToDecimal(reader("valor_cuota"))

                    If montoRestante >= valorCuota Then
                        cuotasAPagar.Add(idCuota)
                        montoRestante -= valorCuota
                    End If
                End While

                reader.Close()

                ' Marcar cuotas como pagadas
                For Each idCuota In cuotasAPagar
                    Dim consultaUpdate As String = "
                        UPDATE cuotas_generadas_apartamento 
                        SET estado = 'pagada', fecha_pago = datetime('now')
                        WHERE id_cuota = @idCuota"

                    Using comandoUpdate As New SQLiteCommand(consultaUpdate, conexion, transaccion)
                        comandoUpdate.Parameters.AddWithValue("@idCuota", idCuota)
                        comandoUpdate.ExecuteNonQuery()
                    End Using
                Next
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Registra el pago de intereses en la tabla correspondiente
    ''' </summary>
    Private Shared Sub RegistrarPagoIntereses(idApartamento As Integer, montoIntereses As Decimal, conexion As SQLiteConnection, transaccion As SQLiteTransaction)
        ' Insertar en tabla de cálculos de interés como pagado
        Dim consulta As String = "
            INSERT INTO calculos_interes (id_apartamentos, fecha_calculo, valor_interes, pagado, fecha_pago)
            VALUES (@idApartamento, date('now'), @valorInteres, 1, datetime('now'))"

        Using comando As New SQLiteCommand(consulta, conexion, transaccion)
            comando.Parameters.AddWithValue("@idApartamento", idApartamento)
            comando.Parameters.AddWithValue("@valorInteres", montoIntereses)
            comando.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' Revierte cuotas que habían sido marcadas como pagadas (para anulaciones)
    ''' </summary>
    Private Shared Sub RevertirCuotasPagadas(idApartamento As Integer, montoPagado As Decimal, conexion As SQLiteConnection, transaccion As SQLiteTransaction)
        ' Obtener cuotas pagadas más recientes del monto correspondiente
        Dim consulta As String = "
            SELECT id_cuota, valor_cuota 
            FROM cuotas_generadas_apartamento 
            WHERE id_apartamentos = @idApartamento 
              AND estado = 'pagada' 
            ORDER BY fecha_pago DESC"

        Dim montoRestante As Decimal = montoPagado

        Using comando As New SQLiteCommand(consulta, conexion, transaccion)
            comando.Parameters.AddWithValue("@idApartamento", idApartamento)
            Using reader As SQLiteDataReader = comando.ExecuteReader()
                Dim cuotasARevertir As New List(Of Integer)

                While reader.Read() AndAlso montoRestante > 0
                    Dim idCuota As Integer = Convert.ToInt32(reader("id_cuota"))
                    Dim valorCuota As Decimal = Convert.ToDecimal(reader("valor_cuota"))

                    If montoRestante >= valorCuota Then
                        cuotasARevertir.Add(idCuota)
                        montoRestante -= valorCuota
                    End If
                End While

                reader.Close()

                ' Revertir estado de cuotas
                For Each idCuota In cuotasARevertir
                    Dim consultaUpdate As String = "
                        UPDATE cuotas_generadas_apartamento 
                        SET estado = 'pendiente', fecha_pago = NULL
                        WHERE id_cuota = @idCuota"

                    Using comandoUpdate As New SQLiteCommand(consultaUpdate, conexion, transaccion)
                        comandoUpdate.Parameters.AddWithValue("@idCuota", idCuota)
                        comandoUpdate.ExecuteNonQuery()
                    End Using
                Next
            End Using
        End Using
    End Sub

End Class