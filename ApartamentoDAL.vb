Imports System.Data.SQLite

Public Class ApartamentoDAL

    ' Método para obtener apartamentos de una torre específica
    Public Shared Function ObtenerApartamentosPorTorre(torre As Integer) As List(Of Apartamento)
        Dim apartamentos As New List(Of Apartamento)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' CORREGIDO: Campo id_piso en lugar de id_piso
                Dim consulta As String = "
                    SELECT 
                        a.id_apartamentos,
                        a.id_torre,
                        a.id_piso,
                        a.numero_apartamento,
                        a.nombre_residente,
                        a.telefono,
                        a.correo,
                        a.matricula_inmobiliaria,
                        COALESCE(p.saldo_actual, 0) as saldo_actual,
                        MAX(p.fecha_pago) as ultimo_pago
                    FROM Apartamentos a
                    LEFT JOIN pagos p ON a.id_apartamentos = p.id_apartamentos
                    WHERE a.id_torre = @torre
                    GROUP BY a.id_apartamentos
                    ORDER BY a.id_piso, a.numero_apartamento"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@torre", torre)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim apartamento As New Apartamento With {
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .Torre = Convert.ToInt32(reader("id_torre")),
                                .Piso = Convert.ToInt32(reader("id_piso")),
                                .NumeroApartamento = reader("numero_apartamento").ToString(),
                                .NombreResidente = If(IsDBNull(reader("nombre_residente")), "", reader("nombre_residente").ToString()),
                                .Telefono = If(IsDBNull(reader("telefono")), "", reader("telefono").ToString()),
                                .Correo = If(IsDBNull(reader("correo")), "", reader("correo").ToString()),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .Activo = True,
                                .SaldoActual = If(IsDBNull(reader("saldo_actual")), 0D, Convert.ToDecimal(reader("saldo_actual")))
                            }

                            ' Asignar fecha del último pago si existe
                            If Not IsDBNull(reader("ultimo_pago")) Then
                                apartamento.UltimoPago = Convert.ToDateTime(reader("ultimo_pago"))
                                apartamento.TieneUltimoPago = True
                            Else
                                apartamento.TieneUltimoPago = False
                            End If

                            apartamentos.Add(apartamento)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al obtener apartamentos de la torre {torre}: {ex.Message}")
        End Try

        Return apartamentos
    End Function

    ' Método para obtener resumen de una torre
    Public Shared Function ObtenerResumenTorre(torre As Integer) As Dictionary(Of String, Object)
        Dim resumen As New Dictionary(Of String, Object)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT 
                        COUNT(*) as total_apartamentos,
                        SUM(CASE WHEN COALESCE(p.saldo_actual, 0) = 0 THEN 1 ELSE 0 END) as apartamentos_al_dia,
                        SUM(CASE WHEN COALESCE(p.saldo_actual, 0) > 0 THEN 1 ELSE 0 END) as apartamentos_pendientes,
                        SUM(CASE WHEN COALESCE(p.saldo_actual, 0) < 0 THEN 1 ELSE 0 END) as apartamentos_a_favor,
                        SUM(CASE WHEN COALESCE(p.saldo_actual, 0) > 0 THEN p.saldo_actual ELSE 0 END) as total_pendiente,
                        ABS(SUM(CASE WHEN COALESCE(p.saldo_actual, 0) < 0 THEN p.saldo_actual ELSE 0 END)) as total_a_favor
                    FROM Apartamentos a
                    LEFT JOIN (
                        SELECT id_apartamentos, 
                               saldo_actual,
                               ROW_NUMBER() OVER (PARTITION BY id_apartamentos ORDER BY fecha_pago DESC) as rn
                        FROM pagos 
                    ) p ON a.id_apartamentos = p.id_apartamentos AND p.rn = 1
                    WHERE a.id_torre = @torre"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@torre", torre)

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            resumen("total_apartamentos") = Convert.ToInt32(reader("total_apartamentos"))
                            resumen("apartamentos_al_dia") = Convert.ToInt32(reader("apartamentos_al_dia"))
                            resumen("apartamentos_pendientes") = Convert.ToInt32(reader("apartamentos_pendientes"))
                            resumen("apartamentos_a_favor") = Convert.ToInt32(reader("apartamentos_a_favor"))
                            resumen("total_pendiente") = Convert.ToDecimal(reader("total_pendiente"))
                            resumen("total_a_favor") = Convert.ToDecimal(reader("total_a_favor"))
                        Else
                            ' Valores por defecto si no hay datos
                            resumen("total_apartamentos") = 0
                            resumen("apartamentos_al_dia") = 0
                            resumen("apartamentos_pendientes") = 0
                            resumen("apartamentos_a_favor") = 0
                            resumen("total_pendiente") = 0
                            resumen("total_a_favor") = 0
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' En caso de error, devolver valores por defecto
            resumen("total_apartamentos") = 0
            resumen("apartamentos_al_dia") = 0
            resumen("apartamentos_pendientes") = 0
            resumen("apartamentos_a_favor") = 0
            resumen("total_pendiente") = 0
            resumen("total_a_favor") = 0
        End Try

        Return resumen
    End Function

    ' Método MEJORADO para actualizar información del propietario
    Public Shared Function ActualizarPropietario(apartamento As Apartamento) As Boolean
        If apartamento Is Nothing Then
            Throw New ArgumentNullException("apartamento", "El objeto apartamento no puede ser nulo")
        End If

        If apartamento.IdApartamento <= 0 Then
            Throw New ArgumentException("El ID del apartamento debe ser válido")
        End If

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Iniciar transacción para asegurar consistencia
                Using transaccion As SQLiteTransaction = conexion.BeginTransaction()
                    Try
                        ' Verificar que el apartamento existe
                        Dim consultaExiste As String = "SELECT COUNT(*) FROM Apartamentos WHERE id_apartamentos = @id"
                        Using comandoExiste As New SQLiteCommand(consultaExiste, conexion, transaccion)
                            comandoExiste.Parameters.AddWithValue("@id", apartamento.IdApartamento)
                            Dim existe As Integer = Convert.ToInt32(comandoExiste.ExecuteScalar())

                            If existe = 0 Then
                                Throw New Exception($"No existe un apartamento con ID {apartamento.IdApartamento}")
                            End If
                        End Using

                        ' Actualizar información del propietario
                        Dim consulta As String = "
                            UPDATE Apartamentos 
                            SET nombre_residente = @nombre,
                                telefono = @telefono,
                                correo = @correo,
                                matricula_inmobiliaria = @matricula
                            WHERE id_apartamentos = @id"

                        Using comando As New SQLiteCommand(consulta, conexion, transaccion)
                            comando.Parameters.AddWithValue("@nombre", If(String.IsNullOrWhiteSpace(apartamento.NombreResidente), DBNull.Value, CObj(apartamento.NombreResidente.Trim())))
                            comando.Parameters.AddWithValue("@telefono", If(String.IsNullOrWhiteSpace(apartamento.Telefono), DBNull.Value, CObj(apartamento.Telefono.Trim())))
                            comando.Parameters.AddWithValue("@correo", If(String.IsNullOrWhiteSpace(apartamento.Correo), DBNull.Value, CObj(apartamento.Correo.Trim())))
                            comando.Parameters.AddWithValue("@matricula", If(String.IsNullOrWhiteSpace(apartamento.MatriculaInmobiliaria), DBNull.Value, CObj(apartamento.MatriculaInmobiliaria.Trim())))
                            comando.Parameters.AddWithValue("@id", apartamento.IdApartamento)

                            Dim filasAfectadas As Integer = comando.ExecuteNonQuery()

                            If filasAfectadas > 0 Then
                                ' Confirmar transacción
                                transaccion.Commit()
                                Return True
                            Else
                                ' Rollback si no se actualizó ninguna fila
                                transaccion.Rollback()
                                Return False
                            End If
                        End Using

                    Catch ex As Exception
                        ' Rollback en caso de error
                        transaccion.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al actualizar propietario: {ex.Message}")
        End Try
    End Function

    ' Método para obtener todos los apartamentos (CORREGIDO PARA TU BD)
    Public Shared Function ObtenerTodosLosApartamentos() As List(Of Apartamento)
        Dim apartamentos As New List(Of Apartamento)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' CORREGIDO: Campo id_piso en lugar de id_piso
                Dim consulta As String = "
                    SELECT 
                        a.id_apartamentos,
                        a.id_torre,
                        a.id_piso,
                        a.numero_apartamento,
                        a.nombre_residente,
                        a.telefono,
                        a.correo,
                        a.matricula_inmobiliaria,
                        COALESCE(p.saldo_actual, 0) as saldo_actual,
                        MAX(p.fecha_pago) as ultimo_pago
                    FROM Apartamentos a
                    LEFT JOIN pagos p ON a.id_apartamentos = p.id_apartamentos
                    GROUP BY a.id_apartamentos
                    ORDER BY a.id_torre, a.id_piso, a.numero_apartamento"

                Using comando As New SQLiteCommand(consulta, conexion)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim apartamento As New Apartamento With {
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .Torre = Convert.ToInt32(reader("id_torre")),
                                .Piso = Convert.ToInt32(reader("id_piso")),
                                .NumeroApartamento = reader("numero_apartamento").ToString(),
                                .NombreResidente = If(IsDBNull(reader("nombre_residente")), "", reader("nombre_residente").ToString()),
                                .Telefono = If(IsDBNull(reader("telefono")), "", reader("telefono").ToString()),
                                .Correo = If(IsDBNull(reader("correo")), "", reader("correo").ToString()),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .Activo = True,
                                .FechaRegistro = DateTime.Now,
                                .SaldoActual = If(IsDBNull(reader("saldo_actual")), 0D, Convert.ToDecimal(reader("saldo_actual")))
                            }

                            ' Asignar fecha del último pago si existe
                            If Not IsDBNull(reader("ultimo_pago")) Then
                                apartamento.UltimoPago = Convert.ToDateTime(reader("ultimo_pago"))
                                apartamento.TieneUltimoPago = True
                            Else
                                apartamento.TieneUltimoPago = False
                            End If

                            apartamentos.Add(apartamento)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al obtener todos los apartamentos: {ex.Message}", "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New List(Of Apartamento)() ' Retornar lista vacía en caso de error
        End Try

        Return apartamentos
    End Function

    ' Método para crear un nuevo apartamento (CORREGIDO PARA TU BD)
    Public Shared Function CrearApartamento(apartamento As Apartamento) As Integer
        If apartamento Is Nothing Then
            Throw New ArgumentNullException("apartamento", "El objeto apartamento no puede ser nulo")
        End If

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' Verificar que no exista ya un apartamento con el mismo número en la misma torre y piso
                Dim consultaExiste As String = "SELECT COUNT(*) FROM Apartamentos WHERE id_torre = @torre AND id_piso = @piso AND numero_apartamento = @numero"
                Using comandoExiste As New SQLiteCommand(consultaExiste, conexion)
                    comandoExiste.Parameters.AddWithValue("@torre", apartamento.Torre)
                    comandoExiste.Parameters.AddWithValue("@piso", apartamento.Piso)
                    comandoExiste.Parameters.AddWithValue("@numero", apartamento.NumeroApartamento)

                    Dim existe As Integer = Convert.ToInt32(comandoExiste.ExecuteScalar())
                    If existe > 0 Then
                        Throw New Exception($"Ya existe un apartamento {apartamento.NumeroApartamento} en la Torre {apartamento.Torre}, Piso {apartamento.Piso}")
                    End If
                End Using

                ' CORREGIDO: Campo id_piso en lugar de id_piso
                Dim consulta As String = "
                    INSERT INTO Apartamentos (id_torre, id_piso, numero_apartamento, nombre_residente, telefono, correo, matricula_inmobiliaria)
                    VALUES (@torre, @piso, @numero, @nombre, @telefono, @correo, @matricula);
                    SELECT last_insert_rowid();"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@torre", apartamento.Torre)
                    comando.Parameters.AddWithValue("@piso", apartamento.Piso)
                    comando.Parameters.AddWithValue("@numero", apartamento.NumeroApartamento)
                    comando.Parameters.AddWithValue("@nombre", If(String.IsNullOrWhiteSpace(apartamento.NombreResidente), DBNull.Value, CObj(apartamento.NombreResidente.Trim())))
                    comando.Parameters.AddWithValue("@telefono", If(String.IsNullOrWhiteSpace(apartamento.Telefono), DBNull.Value, CObj(apartamento.Telefono.Trim())))
                    comando.Parameters.AddWithValue("@correo", If(String.IsNullOrWhiteSpace(apartamento.Correo), DBNull.Value, CObj(apartamento.Correo.Trim())))
                    comando.Parameters.AddWithValue("@matricula", If(String.IsNullOrWhiteSpace(apartamento.MatriculaInmobiliaria), DBNull.Value, CObj(apartamento.MatriculaInmobiliaria.Trim())))

                    Return Convert.ToInt32(comando.ExecuteScalar())
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al crear apartamento: {ex.Message}")
        End Try
    End Function

    ' Método para buscar apartamentos por criterio
    Public Shared Function BuscarApartamentos(criterio As String) As List(Of Apartamento)
        Dim apartamentos As New List(Of Apartamento)

        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                ' CORREGIDO: Campo id_piso en lugar de id_piso
                Dim consulta As String = "
                    SELECT 
                        a.id_apartamentos,
                        a.id_torre,
                        a.id_piso,
                        a.numero_apartamento,
                        a.nombre_residente,
                        a.telefono,
                        a.correo,
                        a.matricula_inmobiliaria,
                        COALESCE(p.saldo_actual, 0) as saldo_actual,
                        MAX(p.fecha_pago) as ultimo_pago
                    FROM Apartamentos a
                    LEFT JOIN pagos p ON a.id_apartamentos = p.id_apartamentos
                    WHERE (
                        a.numero_apartamento LIKE @criterio OR
                        a.nombre_residente LIKE @criterio OR
                        a.telefono LIKE @criterio OR
                        a.correo LIKE @criterio
                    )
                    GROUP BY a.id_apartamentos
                    ORDER BY a.id_torre, a.id_piso, a.numero_apartamento"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@criterio", $"%{criterio}%")

                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        While reader.Read()
                            Dim apartamento As New Apartamento With {
                                .IdApartamento = Convert.ToInt32(reader("id_apartamentos")),
                                .Torre = Convert.ToInt32(reader("id_torre")),
                                .Piso = Convert.ToInt32(reader("id_piso")),
                                .NumeroApartamento = reader("numero_apartamento").ToString(),
                                .NombreResidente = If(IsDBNull(reader("nombre_residente")), "", reader("nombre_residente").ToString()),
                                .Telefono = If(IsDBNull(reader("telefono")), "", reader("telefono").ToString()),
                                .Correo = If(IsDBNull(reader("correo")), "", reader("correo").ToString()),
                                .MatriculaInmobiliaria = If(IsDBNull(reader("matricula_inmobiliaria")), "", reader("matricula_inmobiliaria").ToString()),
                                .Activo = True,
                                .SaldoActual = If(IsDBNull(reader("saldo_actual")), 0D, Convert.ToDecimal(reader("saldo_actual")))
                            }

                            If Not IsDBNull(reader("ultimo_pago")) Then
                                apartamento.UltimoPago = Convert.ToDateTime(reader("ultimo_pago"))
                                apartamento.TieneUltimoPago = True
                            End If

                            apartamentos.Add(apartamento)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception($"Error al buscar apartamentos: {ex.Message}")
        End Try

        Return apartamentos
    End Function

    ' Método para obtener el último saldo de un apartamento
    Public Shared Function ObtenerUltimoSaldo(idApartamento As Integer) As Decimal
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "
                    SELECT saldo_actual 
                    FROM pagos 
                    WHERE id_apartamentos = @id 
                    ORDER BY fecha_pago DESC 
                    LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)

                    Dim resultado = comando.ExecuteScalar()
                    Return If(resultado IsNot Nothing AndAlso Not IsDBNull(resultado), Convert.ToDecimal(resultado), 0)
                End Using
            End Using

        Catch ex As Exception
            Return 0
        End Try
    End Function

    ' Obtiene un objeto Apartamento completo por su Id (MEJORADO)
    Public Shared Function ObtenerApartamentoPorId(idApartamento As Integer) As Apartamento
        If idApartamento <= 0 Then
            Throw New ArgumentException("El ID del apartamento debe ser mayor a 0")
        End If

        Dim apartamento As Apartamento = Nothing
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                ' CORREGIDO: Campo id_piso en lugar de id_piso
                Dim consulta As String = "
                    SELECT 
                        a.id_apartamentos,
                        a.id_torre,
                        a.id_piso,
                        a.numero_apartamento,
                        a.nombre_residente,
                        a.telefono,
                        a.correo,
                        a.matricula_inmobiliaria,
                        COALESCE(p.saldo_actual, 0) as saldo_actual
                    FROM Apartamentos a
                    LEFT JOIN pagos p ON a.id_apartamentos = p.id_apartamentos
                    WHERE a.id_apartamentos = @idApartamento 
                    ORDER BY p.fecha_pago DESC
                    LIMIT 1"

                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Using reader As SQLiteDataReader = comando.ExecuteReader()
                        If reader.Read() Then
                            apartamento = New Apartamento()
                            apartamento.IdApartamento = Convert.ToInt32(reader("id_apartamentos"))
                            apartamento.Torre = Convert.ToInt32(reader("id_torre"))
                            apartamento.Piso = Convert.ToInt32(reader("id_piso"))
                            apartamento.NumeroApartamento = reader("numero_apartamento").ToString()
                            apartamento.MatriculaInmobiliaria = If(reader("matricula_inmobiliaria") Is DBNull.Value, "", reader("matricula_inmobiliaria").ToString())
                            apartamento.NombreResidente = If(reader("nombre_residente") Is DBNull.Value, String.Empty, reader("nombre_residente").ToString())
                            apartamento.Correo = If(reader("correo") Is DBNull.Value, String.Empty, reader("correo").ToString())
                            apartamento.Telefono = If(reader("telefono") Is DBNull.Value, String.Empty, reader("telefono").ToString())
                            apartamento.SaldoActual = If(IsDBNull(reader("saldo_actual")), 0D, Convert.ToDecimal(reader("saldo_actual")))
                            apartamento.Activo = True
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception($"Error al obtener apartamento por ID {idApartamento}: {ex.Message}")
        End Try
        Return apartamento
    End Function

    ' Obtener total de intereses calculados
    Public Shared Function ObtenerTotalInteresesCalculados(idApartamento As Integer) As Decimal
        Dim totalIntereses As Decimal = 0D
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT COALESCE(SUM(valor_interes), 0) FROM calculos_interes WHERE id_apartamentos = @idApartamento AND pagado = 0"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@idApartamento", idApartamento)
                    Dim resultado = comando.ExecuteScalar()
                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        totalIntereses = Convert.ToDecimal(resultado)
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' En caso de error, retornar 0
            totalIntereses = 0D
        End Try
        Return totalIntereses
    End Function

    ' Método para validar si un apartamento existe
    Public Shared Function ExisteApartamento(idApartamento As Integer) As Boolean
        Try
            Using conexion As SQLiteConnection = ConexionBD.ObtenerConexion()
                conexion.Open()
                Dim consulta As String = "SELECT COUNT(*) FROM Apartamentos WHERE id_apartamentos = @id"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@id", idApartamento)
                    Return Convert.ToInt32(comando.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class