' ============================================================================
' PAGOSEXTRA DAL - CAPA DE ACCESO A DATOS PARA PAGOS EXTRA
' ✅ Extiende la funcionalidad de PagosDAL para manejar pagos extra
' ✅ Compatible con la estructura existente de la base de datos
' ============================================================================

Imports System.Data.SQLite
Imports System.Windows.Forms

Public Class PagosExtraDAL

    ''' <summary>
    ''' Registra un pago extra en la base de datos usando la tabla existente 'pagos'
    ''' </summary>
    Public Shared Function RegistrarPagoExtra(pagoExtra As PagoModel) As Boolean
        If pagoExtra Is Nothing Then
            MessageBox.Show("El objeto pago extra no puede ser nulo", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        ' Validar campos específicos para pagos extra
        If String.IsNullOrEmpty(pagoExtra.TipoPago) Then
            MessageBox.Show("Debe especificar el tipo de pago extra", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        If pagoExtra.TotalPagado <= 0 Then
            MessageBox.Show("El valor del pago extra debe ser mayor a cero", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' 1. Verificar que el número de recibo sea único
                        If ExisteNumeroRecibo(pagoExtra.NumeroRecibo, conexion, transaccion) Then
                            Throw New Exception("El número de recibo " & pagoExtra.NumeroRecibo & " ya existe")
                        End If

                        ' 2. Para pagos extra, usar un ID de cuota especial o crear uno temporal
                        Dim idCuotaParaUsar As Integer = ObtenerOCrearCuotaPagoExtra(pagoExtra.IdApartamento, pagoExtra.TipoPago, pagoExtra.TotalPagado, conexion, transaccion)

                        ' 3. Insertar el pago extra en la tabla pagos
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
                            comandoPago.Parameters.AddWithValue("@idApartamento", pagoExtra.IdApartamento)
                            comandoPago.Parameters.AddWithValue("@idCuota", idCuotaParaUsar)
                            comandoPago.Parameters.AddWithValue("@fechaPago", pagoExtra.FechaPago.ToString("yyyy-MM-dd"))
                            comandoPago.Parameters.AddWithValue("@numeroRecibo", pagoExtra.NumeroRecibo)
                            comandoPago.Parameters.AddWithValue("@saldoAnterior", pagoExtra.SaldoAnterior)
                            comandoPago.Parameters.AddWithValue("@pagoAdmin", pagoExtra.PagoAdministracion)
                            comandoPago.Parameters.AddWithValue("@pagoIntereses", pagoExtra.PagoIntereses)
                            comandoPago.Parameters.AddWithValue("@cuotaActual", pagoExtra.CuotaActual)
                            comandoPago.Parameters.AddWithValue("@totalPagado", pagoExtra.TotalPagado)
                            comandoPago.Parameters.AddWithValue("@saldoActual", pagoExtra.SaldoActual)
                            comandoPago.Parameters.AddWithValue("@detalle", pagoExtra.Detalle)
                            comandoPago.Parameters.AddWithValue("@observaciones", If(String.IsNullOrEmpty(pagoExtra.Observaciones), "", pagoExtra.Observaciones))
                            comandoPago.Parameters.AddWithValue("@estadoPago", pagoExtra.EstadoPago)
                            comandoPago.Parameters.AddWithValue("@registradoPor", If(String.IsNullOrEmpty(pagoExtra.UsuarioRegistro), "Sistema", pagoExtra.UsuarioRegistro))

                            comandoPago.ExecuteNonQuery()
                        End Using

                        ' 4. Registrar en histórico de cambios con tipo específico
                        ConexionBD.RegistrarActividad(
                            If(String.IsNullOrEmpty(pagoExtra.UsuarioRegistro), "Sistema", pagoExtra.UsuarioRegistro),
                            "pagos",
                            pagoExtra.IdApartamento,
                            "INSERT_PAGO_EXTRA",
                            $"Pago Extra {pagoExtra.TipoPago} registrado - Recibo: {pagoExtra.NumeroRecibo}, Total: {pagoExtra.TotalPagado:C}"
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
            MessageBox.Show("Error al registrar pago extra: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Obtiene un pago extra por número de recibo
    ''' </summary>
    Public Shared Function ObtenerPagoExtraPorNumeroRecibo(numeroRecibo As String) As PagoModel
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

                            ' Extraer tipo de pago del detalle si es un pago extra
                            If pago.Detalle.StartsWith("PAGO EXTRA - ") Then
                                Dim partes As String() = pago.Detalle.Split(":"c)
                                If partes.Length > 0 Then
                                    pago.TipoPago = partes(0).Replace("PAGO EXTRA - ", "").Trim()
                                End If
                            Else
                                pago.TipoPago = "PAGO_EXTRA"
                            End If

                            ' Obtener matrícula inmobiliaria
                            pago.MatriculaInmobiliaria = PagosDAL.ObtenerMatriculaInmobiliaria(pago.IdApartamento)

                            Return pago
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al obtener pago extra: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Obtiene el historial de pagos extra de un apartamento
    ''' </summary>
    Public Shared Function ObtenerHistorialPagosExtra(idApartamento As Integer) As List(Of PagoModel)
        Dim pagosExtra As New List(Of PagoModel)
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Filtrar pagos que sean identificados como "pago extra" por el detalle
                Dim consulta As String = "
                    SELECT * FROM pagos 
                    WHERE id_apartamentos = @idApartamento 
                    AND (detalle LIKE 'PAGO EXTRA -%' OR vr_pagado_administracion = 0)
                    ORDER BY fecha_pago DESC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

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

                            ' Identificar tipo de pago extra
                            If pago.Detalle.StartsWith("PAGO EXTRA - ") Then
                                Dim partes As String() = pago.Detalle.Split(":"c)
                                If partes.Length > 0 Then
                                    pago.TipoPago = partes(0).Replace("PAGO EXTRA - ", "").Trim()
                                End If
                            End If

                            pagosExtra.Add(pago)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al obtener historial de pagos extra: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return pagosExtra
    End Function

    ''' <summary>
    ''' Obtiene estadísticas de pagos extra
    ''' </summary>
    Public Shared Function ObtenerEstadisticasPagosExtra() As Dictionary(Of String, Object)
        Dim estadisticas As New Dictionary(Of String, Object)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Estadísticas del mes actual para pagos extra
                Dim consultaMes As String = "
                    SELECT 
                        COUNT(*) as pagos_extra_mes,
                        COALESCE(SUM(total_pagado), 0) as total_recaudado_extra_mes
                    FROM pagos 
                    WHERE strftime('%Y-%m', fecha_pago) = strftime('%Y-%m', 'now')
                    AND (detalle LIKE 'PAGO EXTRA -%' OR vr_pagado_administracion = 0)"

                Using comando As New SQLiteCommand(consultaMes, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("pagos_extra_mes_actual") = Convert.ToInt32(reader("pagos_extra_mes"))
                            estadisticas("recaudacion_extra_mes_actual") = Convert.ToDecimal(reader("total_recaudado_extra_mes"))
                        End If
                    End Using
                End Using

                ' Estadísticas por tipo de pago extra
                Dim consultaTipos As String = "
                    SELECT 
                        CASE 
                            WHEN detalle LIKE '%MULTA%' THEN 'MULTA'
                            WHEN detalle LIKE '%ADICION%' THEN 'ADICION'
                            WHEN detalle LIKE '%SANCION%' THEN 'SANCION'
                            WHEN detalle LIKE '%REPARACION%' THEN 'REPARACION'
                            ELSE 'OTROS'
                        END as tipo_pago,
                        COUNT(*) as cantidad,
                        SUM(total_pagado) as total_valor
                    FROM pagos 
                    WHERE detalle LIKE 'PAGO EXTRA -%'
                    GROUP BY tipo_pago"

                Using comando As New SQLiteCommand(consultaTipos, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        Dim tiposPago As New Dictionary(Of String, Dictionary(Of String, Object))

                        While reader.Read()
                            Dim tipoPago As String = reader("tipo_pago").ToString()
                            tiposPago(tipoPago) = New Dictionary(Of String, Object) From {
                                {"cantidad", Convert.ToInt32(reader("cantidad"))},
                                {"total_valor", Convert.ToDecimal(reader("total_valor"))}
                            }
                        End While

                        estadisticas("tipos_pago_extra") = tiposPago
                    End Using
                End Using

            End Using

        Catch ex As Exception
            ' Valores por defecto en caso de error
            estadisticas("pagos_extra_mes_actual") = 0
            estadisticas("recaudacion_extra_mes_actual") = 0D
            estadisticas("tipos_pago_extra") = New Dictionary(Of String, Dictionary(Of String, Object))
        End Try

        Return estadisticas
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
    ''' Obtiene o crea una cuota específica para pagos extra
    ''' </summary>
    Private Shared Function ObtenerOCrearCuotaPagoExtra(idApartamento As Integer, tipoPago As String, valor As Decimal, conexion As SQLiteConnection, transaccion As SQLiteTransaction) As Integer
        Try
            ' Obtener matrícula inmobiliaria
            Dim matricula As String = ""
            Dim consultaMatricula As String = "SELECT COALESCE(matricula_inmobiliaria, '') FROM Apartamentos WHERE id_apartamentos = @id"
            Using cmdMatricula As New SQLiteCommand(consultaMatricula, conexion, transaccion)
                cmdMatricula.Parameters.AddWithValue("@id", idApartamento)
                Dim resultado = cmdMatricula.ExecuteScalar()
                matricula = If(resultado IsNot Nothing, resultado.ToString(), "N/A")
            End Using

            ' Crear cuota específica para pago extra
            Dim consulta As String = "
            INSERT INTO cuotas_generadas_apartamento 
            (id_apartamentos, matricula_inmobiliaria, fecha_cuota, valor_cuota, 
             fecha_vencimiento, estado, tipo_cuota, tipo_piso, id_asamblea)
            VALUES 
            (@idApartamento, @matricula, date('now'), @valor, date('now'), 
             'pagada', @tipoCuota, 'PAGO_EXTRA', 1)"

            Using comando As New SQLiteCommand(consulta, conexion, transaccion)
                comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                comando.Parameters.AddWithValue("@matricula", matricula)
                comando.Parameters.AddWithValue("@valor", valor)
                comando.Parameters.AddWithValue("@tipoCuota", $"Pago Extra - {tipoPago}")

                comando.ExecuteNonQuery()
                Return Convert.ToInt32(conexion.LastInsertRowId)
            End Using

        Catch ex As Exception
            ' Si falla la creación, usar ID genérico
            System.Diagnostics.Debug.WriteLine($"Error creando cuota para pago extra: {ex.Message}")
            Return 1
        End Try
    End Function

End Class