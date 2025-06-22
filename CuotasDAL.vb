' ============================================================================
' CUOTAS DAL COMPLETO - IMPLEMENTACIÓN DE MÉTODOS FALTANTES
' Añade todas las funcionalidades necesarias para gestión de cuotas
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
        Public TipoCuota As String
        Public TipoPiso As String
    End Structure

    Public Class CuotasDAL

        Public Class CuotaPendienteInfo
            Public Property ExisteCuotaPendiente As Boolean
            Public Property IdCuota As Integer

            Public Sub New()
                ExisteCuotaPendiente = False
                IdCuota = 1
            End Sub
        End Class

        Public Shared Function ObtenerCuotaPendienteMasAntigua(idApartamento As Integer) As CuotaPendienteInfo
            Dim resultado As New CuotaPendienteInfo()

            Try
                Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                    conexion.Open()

                    Dim consulta As String = "SELECT id_cuota FROM cuotas WHERE id_apartamentos = @id AND estado = 'pendiente' ORDER BY fecha_incio ASC LIMIT 1"

                    Using comando As New SQLiteCommand(consulta, conexion)
                        comando.Parameters.AddWithValue("@id", idApartamento)
                        Dim resultadoConsulta = comando.ExecuteScalar()

                        If resultadoConsulta IsNot Nothing Then
                            resultado.ExisteCuotaPendiente = True
                            resultado.IdCuota = Convert.ToInt32(resultadoConsulta)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                ' En caso de error, usar valores por defecto
                resultado.ExisteCuotaPendiente = False
                resultado.IdCuota = 1
            End Try

            Return resultado
        End Function

    End Class

    ''' <summary>
    ''' Obtiene la cuota pendiente más antigua de un apartamento
    ''' </summary>
    Public Shared Function ObtenerCuotaPendienteMasAntigua(idApartamento As Integer) As CuotaPendienteInfo
        Dim resultado As New CuotaPendienteInfo()
        resultado.ExisteCuotaPendiente = False

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        id_cuota, 
                        fecha_incio, 
                        valor_cuota, 
                        fecha_cuota,
                        tipo_cuota,
                        tipo_piso
                    FROM cuotas_generadas_apartamento 
                    WHERE id_apartamentos = @idApartamento 
                        AND estado = 'pendiente' 
                        AND date(fecha_incio) < date('now') 
                    ORDER BY fecha_incio ASC 
                    LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            resultado.ExisteCuotaPendiente = True
                            resultado.IdCuota = Convert.ToInt32(reader("id_cuota"))
                            resultado.FechaVencimiento = Convert.ToDateTime(reader("fecha_incio"))
                            resultado.ValorCuota = Convert.ToDecimal(reader("valor_cuota"))
                            resultado.DiasVencida = (Date.Today - resultado.FechaVencimiento.Date).Days
                            resultado.TipoCuota = If(IsDBNull(reader("tipo_cuota")), "Administración", reader("tipo_cuota").ToString())
                            resultado.TipoPiso = If(IsDBNull(reader("tipo_piso")), "N/A", reader("tipo_piso").ToString())
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
    ''' NUEVO: Obtiene todas las cuotas pendientes por apartamento (FALTANTE EN CÓDIGO ORIGINAL)
    ''' </summary>
    Public Shared Function ObtenerCuotasPendientesPorApartamento(idApartamento As Integer) As List(Of CuotaModel)
        Dim cuotas As New List(Of CuotaModel)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        id_cuota,
                        id_apartamentos,
                        matricula_inmobiliaria,
                        fecha_cuota,
                        valor_cuota,
                        fecha_incio,
                        estado,
                        tipo_cuota,
                        tipo_piso,
                        id_asamblea,
                        fecha_pago,
                        intereses_mora
                    FROM cuotas_generadas_apartamento 
                    WHERE id_apartamentos = @idApartamento 
                        AND estado = 'pendiente'
                    ORDER BY fecha_incio ASC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim cuota As New CuotaModel() With {
                                .IdCuota = Convert.ToInt32(reader("id_cuota")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .FechaCuota = Convert.ToDateTime(reader("fecha_cuota")),
                                .ValorCuota = Convert.ToDecimal(reader("valor_cuota")),
                                .FechaVencimiento = Convert.ToDateTime(reader("fecha_incio")),
                                .Estado = reader("estado").ToString(),
                                .TipoCuota = If(IsDBNull(reader("tipo_cuota")), "Administración", reader("tipo_cuota").ToString()),
                                .TipoPiso = If(IsDBNull(reader("tipo_piso")), "N/A", reader("tipo_piso").ToString()),
                                .IdAsamblea = If(IsDBNull(reader("id_asamblea")), 0, Convert.ToInt32(reader("id_asamblea"))),
                                .InteresesMora = If(IsDBNull(reader("intereses_mora")), 0D, Convert.ToDecimal(reader("intereses_mora")))
                            }

                            ' Asignar fecha de pago si existe
                            If Not IsDBNull(reader("fecha_pago")) Then
                                cuota.FechaPago = Convert.ToDateTime(reader("fecha_pago"))
                            End If

                            cuotas.Add(cuota)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al obtener cuotas pendientes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return cuotas
    End Function

    ''' <summary>
    ''' NUEVO: Obtiene todas las cuotas (pagadas y pendientes) por apartamento
    ''' </summary>
    Public Shared Function ObtenerTodasLasCuotasPorApartamento(idApartamento As Integer) As List(Of CuotaModel)
        Dim cuotas As New List(Of CuotaModel)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        id_cuota,
                        id_apartamentos,
                        matricula_inmobiliaria,
                        fecha_cuota,
                        valor_cuota,
                        fecha_incio,
                        estado,
                        tipo_cuota,
                        tipo_piso,
                        id_asamblea,
                        fecha_pago,
                        intereses_mora
                    FROM cuotas_generadas_apartamento 
                    WHERE id_apartamentos = @idApartamento
                    ORDER BY fecha_incio DESC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim cuota As New CuotaModel() With {
                                .IdCuota = Convert.ToInt32(reader("id_cuota")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .FechaCuota = Convert.ToDateTime(reader("fecha_cuota")),
                                .ValorCuota = Convert.ToDecimal(reader("valor_cuota")),
                                .FechaVencimiento = Convert.ToDateTime(reader("fecha_incio")),
                                .Estado = reader("estado").ToString(),
                                .TipoCuota = If(IsDBNull(reader("tipo_cuota")), "Administración", reader("tipo_cuota").ToString()),
                                .TipoPiso = If(IsDBNull(reader("tipo_piso")), "N/A", reader("tipo_piso").ToString()),
                                .IdAsamblea = If(IsDBNull(reader("id_asamblea")), 0, Convert.ToInt32(reader("id_asamblea"))),
                                .InteresesMora = If(IsDBNull(reader("intereses_mora")), 0D, Convert.ToDecimal(reader("intereses_mora")))
                            }

                            If Not IsDBNull(reader("fecha_pago")) Then
                                cuota.FechaPago = Convert.ToDateTime(reader("fecha_pago"))
                            End If

                            cuotas.Add(cuota)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al obtener cuotas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return cuotas
    End Function

    ''' <summary>
    ''' Marca una cuota como pagada
    ''' </summary>
    Public Shared Function MarcarCuotaComoPagada(idCuota As Integer, idPago As Integer) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    UPDATE cuotas_generadas_apartamento 
                    SET estado = 'pagada', 
                        fecha_pago = datetime('now') 
                    WHERE id_cuota = @idCuota"

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
    ''' NUEVO: Genera cuotas automáticamente para todos los apartamentos
    ''' </summary>
    Public Shared Function GenerarCuotasMasivas(valorCuota As Decimal, fechaVencimiento As DateTime, descripcion As String, tipoPiso As String) As Integer
        Dim cuotasGeneradas As Integer = 0

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' Obtener todos los apartamentos
                        Dim apartamentos = ApartamentoDAL.ObtenerTodosLosApartamentos()

                        For Each apartamento In apartamentos
                            ' Filtrar por tipo de piso si se especifica
                            If tipoPiso <> "Todos" AndAlso Not tipoPiso.Contains(apartamento.Piso.ToString()) Then
                                Continue For
                            End If

                            ' Verificar que no exista ya una cuota para este mes
                            Dim consultaExiste As String = "
                                SELECT COUNT(*) 
                                FROM cuotas_generadas_apartamento 
                                WHERE id_apartamentos = @idApartamento 
                                    AND strftime('%Y-%m', fecha_incio) = strftime('%Y-%m', @fechaVencimiento)"

                            Using comandoExiste As New SQLiteCommand(consultaExiste, conexion, transaccion)
                                comandoExiste.Parameters.AddWithValue("@idApartamento", apartamento.IdApartamento)
                                comandoExiste.Parameters.AddWithValue("@fechaVencimiento", fechaVencimiento)

                                If Convert.ToInt32(comandoExiste.ExecuteScalar()) > 0 Then
                                    Continue For ' Ya existe cuota para este mes
                                End If
                            End Using

                            ' Insertar nueva cuota
                            Dim consultaInsertar As String = "
                                INSERT INTO cuotas_generadas_apartamento 
                                (id_apartamentos, matricula_inmobiliaria, fecha_cuota, valor_cuota, 
                                 fecha_incio, estado, tipo_cuota, tipo_piso, id_asamblea)
                                VALUES 
                                (@idApartamento, @matricula, date('now'), @valor, @fechaVencimiento, 
                                 'pendiente', @descripcion, @tipoPiso, 1)"

                            Using comandoInsertar As New SQLiteCommand(consultaInsertar, conexion, transaccion)
                                comandoInsertar.Parameters.AddWithValue("@idApartamento", apartamento.IdApartamento)
                                comandoInsertar.Parameters.AddWithValue("@matricula", apartamento.MatriculaInmobiliaria)
                                comandoInsertar.Parameters.AddWithValue("@valor", valorCuota)
                                comandoInsertar.Parameters.AddWithValue("@fechaVencimiento", fechaVencimiento)
                                comandoInsertar.Parameters.AddWithValue("@descripcion", descripcion)
                                comandoInsertar.Parameters.AddWithValue("@tipoPiso", tipoPiso)

                                If comandoInsertar.ExecuteNonQuery() > 0 Then
                                    cuotasGeneradas += 1
                                End If
                            End Using
                        Next

                        transaccion.Commit()

                    Catch ex As Exception
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al generar cuotas masivas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return cuotasGeneradas
    End Function

    ''' <summary>
    ''' NUEVO: Obtiene estadísticas de cuotas
    ''' </summary>
    Public Shared Function ObtenerEstadisticasCuotas() As Dictionary(Of String, Object)
        Dim estadisticas As New Dictionary(Of String, Object)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Cuotas del mes actual
                Dim consultaMesActual As String = "
                    SELECT 
                        COUNT(*) as total_cuotas,
                        SUM(CASE WHEN estado = 'pendiente' THEN 1 ELSE 0 END) as pendientes,
                        SUM(CASE WHEN estado = 'pagada' THEN 1 ELSE 0 END) as pagadas,
                        SUM(CASE WHEN estado = 'pendiente' THEN valor_cuota ELSE 0 END) as valor_pendiente,
                        SUM(CASE WHEN estado = 'pagada' THEN valor_cuota ELSE 0 END) as valor_pagado
                    FROM cuotas_generadas_apartamento 
                    WHERE strftime('%Y-%m', fecha_incio) = strftime('%Y-%m', 'now')"

                Using comando As New SQLiteCommand(consultaMesActual, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("total_cuotas_mes") = Convert.ToInt32(reader("total_cuotas"))
                            estadisticas("cuotas_pendientes_mes") = Convert.ToInt32(reader("pendientes"))
                            estadisticas("cuotas_pagadas_mes") = Convert.ToInt32(reader("pagadas"))
                            estadisticas("valor_pendiente_mes") = Convert.ToDecimal(reader("valor_pendiente"))
                            estadisticas("valor_pagado_mes") = Convert.ToDecimal(reader("valor_pagado"))
                        End If
                    End Using
                End Using

                ' Cuotas vencidas
                Dim consultaVencidas As String = "
                    SELECT 
                        COUNT(*) as cuotas_vencidas,
                        SUM(valor_cuota) as valor_vencido
                    FROM cuotas_generadas_apartamento 
                    WHERE estado = 'pendiente' 
                        AND date(fecha_incio) < date('now')"

                Using comando As New SQLiteCommand(consultaVencidas, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            estadisticas("cuotas_vencidas") = Convert.ToInt32(reader("cuotas_vencidas"))
                            estadisticas("valor_vencido") = Convert.ToDecimal(reader("valor_vencido"))
                        End If
                    End Using
                End Using

            End Using

        Catch ex As Exception
            ' Valores por defecto en caso de error
            estadisticas("total_cuotas_mes") = 0
            estadisticas("cuotas_pendientes_mes") = 0
            estadisticas("cuotas_pagadas_mes") = 0
            estadisticas("valor_pendiente_mes") = 0D
            estadisticas("valor_pagado_mes") = 0D
            estadisticas("cuotas_vencidas") = 0
            estadisticas("valor_vencido") = 0D
        End Try

        Return estadisticas
    End Function

    ''' <summary>
    ''' NUEVO: Calcula intereses de mora para cuotas vencidas
    ''' </summary>
    Public Shared Function CalcularInteresesMoraPorApartamento(idApartamento As Integer) As Decimal
        Dim totalIntereses As Decimal = 0D

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT id_cuota, valor_cuota, fecha_incio
                    FROM cuotas_generadas_apartamento 
                    WHERE id_apartamentos = @idApartamento 
                        AND estado = 'pendiente'
                        AND date(fecha_incio) < date('now')"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        Dim tasaInteres As Decimal = ParametrosDAL.ObtenerTasaInteresMoraActual()

                        While reader.Read()
                            Dim valorCuota As Decimal = Convert.ToDecimal(reader("valor_cuota"))
                            Dim fechaVencimiento As DateTime = Convert.ToDateTime(reader("fecha_incio"))
                            Dim diasMora As Integer = (DateTime.Now.Date - fechaVencimiento.Date).Days

                            If diasMora > 0 AndAlso tasaInteres > 0 Then
                                Dim interesCuota As Decimal = valorCuota * (tasaInteres / 100D) * (diasMora / 365D)
                                totalIntereses += interesCuota
                            End If
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error al calcular intereses de mora: {ex.Message}")
        End Try

        Return Math.Round(totalIntereses, 2)
    End Function

    ''' <summary>
    ''' NUEVO: Anula una cuota específica
    ''' </summary>
    Public Shared Function AnularCuota(idCuota As Integer, motivo As String, usuarioAnula As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    UPDATE cuotas_generadas_apartamento 
                    SET estado = 'anulada',
                        tipo_cuota = tipo_cuota || ' (ANULADA: ' || @motivo || ' por ' || @usuario || ' el ' || datetime('now') || ')'
                    WHERE id_cuota = @idCuota"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idCuota", idCuota)
                    comando.Parameters.AddWithValue("@motivo", motivo)
                    comando.Parameters.AddWithValue("@usuario", usuarioAnula)

                    Return comando.ExecuteNonQuery() > 0
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al anular cuota: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

                Dim consulta As String = "
                    SELECT COALESCE(SUM(valor_interes), 0) 
                    FROM calculos_interes 
                    WHERE id_apartamentos = @idApartamento 
                        AND pagado = 0"

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

    ''' <summary>
    ''' NUEVO: Obtiene cuotas por torre
    ''' </summary>
    Public Shared Function ObtenerCuotasPorTorre(numeroTorre As Integer, soloVencidas As Boolean) As List(Of CuotaModel)
        Dim cuotas As New List(Of CuotaModel)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        c.id_cuota,
                        c.id_apartamentos,
                        c.matricula_inmobiliaria,
                        c.fecha_cuota,
                        c.valor_cuota,
                        c.fecha_incio,
                        c.estado,
                        c.tipo_cuota,
                        c.tipo_piso,
                        c.id_asamblea,
                        c.fecha_pago,
                        c.intereses_mora,
                        a.numero_apartamento,
                        a.nombre_residente
                    FROM cuotas_generadas_apartamento c
                    INNER JOIN Apartamentos a ON c.id_apartamentos = a.id_apartamentos
                    WHERE a.id_torre = @torre"

                If soloVencidas Then
                    consulta &= " AND c.estado = 'pendiente' AND date(c.fecha_incio) < date('now')"
                End If

                consulta &= " ORDER BY c.fecha_incio DESC"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@torre", numeroTorre)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim cuota As New CuotaModel() With {
                                .IdCuota = Convert.ToInt32(reader("id_cuota")),
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .FechaCuota = Convert.ToDateTime(reader("fecha_cuota")),
                                .ValorCuota = Convert.ToDecimal(reader("valor_cuota")),
                                .FechaVencimiento = Convert.ToDateTime(reader("fecha_incio")),
                                .Estado = reader("estado").ToString(),
                                .TipoCuota = If(IsDBNull(reader("tipo_cuota")), "Administración", reader("tipo_cuota").ToString()),
                                .TipoPiso = If(IsDBNull(reader("tipo_piso")), "N/A", reader("tipo_piso").ToString()),
                                .IdAsamblea = If(IsDBNull(reader("id_asamblea")), 0, Convert.ToInt32(reader("id_asamblea"))),
                                .InteresesMora = If(IsDBNull(reader("intereses_mora")), 0D, Convert.ToDecimal(reader("intereses_mora")))
                            }

                            If Not IsDBNull(reader("fecha_pago")) Then
                                cuota.FechaPago = Convert.ToDateTime(reader("fecha_pago"))
                            End If

                            cuotas.Add(cuota)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al obtener cuotas por torre: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return cuotas
    End Function

End Class