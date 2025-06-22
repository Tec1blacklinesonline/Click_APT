' ============================================================================
' MODELO DE PAGO CORREGIDO Y MEJORADO
' Añade propiedades faltantes y mejora la validación
' ============================================================================

Public Class PagoModel
    ' Propiedades principales
    Public Property IdPago As Integer
    Public Property IdApartamento As Integer
    Public Property IdCuota As Integer?
    Public Property FechaPago As DateTime
    Public Property NumeroRecibo As String
    Public Property SaldoAnterior As Decimal
    Public Property PagoAdministracion As Decimal
    Public Property PagoIntereses As Decimal
    Public Property CuotaActual As Decimal
    Public Property TotalPagado As Decimal
    Public Property SaldoActual As Decimal
    Public Property Detalle As String
    Public Property Observaciones As String
    Public Property EstadoPago As String
    Public Property UsuarioRegistro As String
    Public Property FechaRegistro As DateTime

    ' CORREGIDO: Propiedades faltantes
    Public Property MatriculaInmobiliaria As String
    Public Property TipoPago As String
    Public Property MetodoPago As String
    Public Property ReferenciaPago As String
    Public Property ComprobanteAdjunto As String

    ' Propiedades calculadas
    Public ReadOnly Property TotalConIntereses As Decimal
        Get
            Return PagoAdministracion + PagoIntereses
        End Get
    End Property

    Public ReadOnly Property CodigoApartamento As String
        Get
            If Not String.IsNullOrEmpty(MatriculaInmobiliaria) Then
                Return MatriculaInmobiliaria
            End If
            Return $"APT-{IdApartamento}"
        End Get
    End Property

    ' Constructor vacío
    Public Sub New()
        EstadoPago = "REGISTRADO"
        FechaPago = DateTime.Now
        FechaRegistro = DateTime.Now
        Detalle = String.Empty
        Observaciones = String.Empty
        MatriculaInmobiliaria = String.Empty
        TipoPago = "ADMINISTRACION"
        MetodoPago = "EFECTIVO"
        ReferenciaPago = String.Empty
        ComprobanteAdjunto = String.Empty
    End Sub

    ' Constructor con parámetros básicos
    Public Sub New(idApartamento As Integer, numeroRecibo As String, totalPagado As Decimal)
        Me.New()
        Me.IdApartamento = idApartamento
        Me.NumeroRecibo = numeroRecibo
        Me.TotalPagado = totalPagado
    End Sub

    ' Constructor completo
    Public Sub New(idApartamento As Integer, numeroRecibo As String, pagoAdmin As Decimal, pagoIntereses As Decimal, saldoAnterior As Decimal)
        Me.New(idApartamento, numeroRecibo, pagoAdmin + pagoIntereses)
        Me.PagoAdministracion = pagoAdmin
        Me.PagoIntereses = pagoIntereses
        Me.SaldoAnterior = saldoAnterior
        Me.SaldoActual = saldoAnterior - Me.TotalPagado
        Me.CuotaActual = pagoAdmin
    End Sub

    ' MEJORADO: Método de validación más completo
    Public Function Validar() As ResultadoValidacion
        Dim resultado As New ResultadoValidacion()

        ' Validar ID de apartamento
        If IdApartamento <= 0 Then
            resultado.AgregarError("El ID del apartamento debe ser mayor a 0")
        End If

        ' Validar número de recibo
        If String.IsNullOrWhiteSpace(NumeroRecibo) Then
            resultado.AgregarError("El número de recibo es obligatorio")
        ElseIf NumeroRecibo.Length < 8 Then
            resultado.AgregarError("El número de recibo debe tener al menos 8 caracteres")
        End If

        ' Validar montos
        If TotalPagado <= 0 Then
            resultado.AgregarError("El total pagado debe ser mayor a 0")
        End If

        If PagoAdministracion < 0 Then
            resultado.AgregarError("El pago de administración no puede ser negativo")
        End If

        If PagoIntereses < 0 Then
            resultado.AgregarError("El pago de intereses no puede ser negativo")
        End If

        ' Validar coherencia de montos
        If Math.Abs(TotalPagado - (PagoAdministracion + PagoIntereses)) > 0.01D Then
            resultado.AgregarError("El total pagado no coincide con la suma de administración e intereses")
        End If

        ' Validar fecha
        If FechaPago > DateTime.Now Then
            resultado.AgregarError("La fecha de pago no puede ser futura")
        End If

        If FechaPago < DateTime.Now.AddYears(-2) Then
            resultado.AgregarAdvertencia("La fecha de pago es muy antigua (más de 2 años)")
        End If

        ' Validar estado
        Dim estadosValidos As String() = {"REGISTRADO", "CONFIRMADO", "ANULADO", "PENDIENTE"}
        Dim estadoValido As Boolean = False
        For Each estado In estadosValidos
            If EstadoPago.ToUpper() = estado Then
                estadoValido = True
                Exit For
            End If
        Next

        If Not estadoValido Then
            resultado.AgregarError("Estado de pago inválido: " & EstadoPago)
        End If

        Return resultado
    End Function

    ' NUEVO: Método para obtener descripción detallada del pago
    Public Function ObtenerDescripcionCompleta() As String
        Dim descripcion As New System.Text.StringBuilder()

        descripcion.AppendLine("Recibo No: " & NumeroRecibo)
        descripcion.AppendLine("Fecha: " & FechaPago.ToString("dd/MM/yyyy"))
        descripcion.AppendLine("Apartamento: " & CodigoApartamento)

        If PagoAdministracion > 0 Then
            descripcion.AppendLine("Administración: " & PagoAdministracion.ToString("C"))
        End If

        If PagoIntereses > 0 Then
            descripcion.AppendLine("Intereses: " & PagoIntereses.ToString("C"))
        End If

        descripcion.AppendLine("Total: " & TotalPagado.ToString("C"))
        descripcion.AppendLine("Estado: " & EstadoPago)

        If Not String.IsNullOrEmpty(Observaciones) Then
            descripcion.AppendLine("Observaciones: " & Observaciones)
        End If

        Return descripcion.ToString()
    End Function

    ' NUEVO: Método para generar resumen para reportes
    Public Function ObtenerResumenReporte() As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {
            {"numeroRecibo", NumeroRecibo},
            {"fechaPago", FechaPago},
            {"apartamento", CodigoApartamento},
            {"pagoAdministracion", PagoAdministracion},
            {"pagoIntereses", PagoIntereses},
            {"totalPagado", TotalPagado},
            {"saldoAnterior", SaldoAnterior},
            {"saldoActual", SaldoActual},
            {"estadoPago", EstadoPago},
            {"metodoPago", MetodoPago},
            {"usuarioRegistro", UsuarioRegistro}
        }
    End Function

    ' NUEVO: Método para clonar el pago (útil para anulaciones)
    Public Function Clonar() As PagoModel
        Return New PagoModel() With {
            .IdPago = Me.IdPago,
            .IdApartamento = Me.IdApartamento,
            .IdCuota = Me.IdCuota,
            .FechaPago = Me.FechaPago,
            .NumeroRecibo = Me.NumeroRecibo,
            .SaldoAnterior = Me.SaldoAnterior,
            .PagoAdministracion = Me.PagoAdministracion,
            .PagoIntereses = Me.PagoIntereses,
            .CuotaActual = Me.CuotaActual,
            .TotalPagado = Me.TotalPagado,
            .SaldoActual = Me.SaldoActual,
            .Detalle = Me.Detalle,
            .Observaciones = Me.Observaciones,
            .EstadoPago = Me.EstadoPago,
            .UsuarioRegistro = Me.UsuarioRegistro,
            .FechaRegistro = Me.FechaRegistro,
            .MatriculaInmobiliaria = Me.MatriculaInmobiliaria,
            .TipoPago = Me.TipoPago,
            .MetodoPago = Me.MetodoPago,
            .ReferenciaPago = Me.ReferenciaPago,
            .ComprobanteAdjunto = Me.ComprobanteAdjunto
        }
    End Function

    ' NUEVO: Método para aplicar descuento
    Public Sub AplicarDescuento(porcentajeDescuento As Decimal, motivo As String)
        If porcentajeDescuento > 0 AndAlso porcentajeDescuento <= 100 Then
            Dim descuento As Decimal = TotalPagado * (porcentajeDescuento / 100D)
            TotalPagado -= descuento
            SaldoActual += descuento

            If Not String.IsNullOrEmpty(Observaciones) Then
                Observaciones &= " | "
            End If
            Observaciones &= "Descuento " & porcentajeDescuento.ToString() & "% aplicado: -" & descuento.ToString("C") & ". Motivo: " & motivo
        End If
    End Sub

    ' NUEVO: Método para verificar si el pago está vencido (para pagos pendientes)
    Public Function EstaVencido(Optional diasGracia As Integer = 5) As Boolean
        If EstadoPago.ToUpper() = "PENDIENTE" Then
            Return DateTime.Now > FechaPago.AddDays(diasGracia)
        End If
        Return False
    End Function

    ' NUEVO: Método para calcular días desde el pago
    Public Function DiasDesdeElPago() As Integer
        Return (DateTime.Now.Date - FechaPago.Date).Days
    End Function

    ' Método ToString mejorado
    Public Overrides Function ToString() As String
        Return "Recibo " & NumeroRecibo & " - " & TotalPagado.ToString("C") & " - " & EstadoPago & " - " & FechaPago.ToString("dd/MM/yyyy")
    End Function

End Class

' ============================================================================
' CLASE DE APOYO PARA VALIDACIONES
' ============================================================================

Public Class ResultadoValidacion
    Public Property EsValido As Boolean
    Public Property Errores As New List(Of String)
    Public Property Advertencias As New List(Of String)

    Public Sub AgregarError(mensaje As String)
        Errores.Add(mensaje)
        EsValido = False
    End Sub

    Public Sub AgregarAdvertencia(mensaje As String)
        Advertencias.Add(mensaje)
    End Sub

    Public Function ObtenerMensajeCompleto() As String
        Dim mensaje As New System.Text.StringBuilder()

        If Errores.Count > 0 Then
            mensaje.AppendLine("❌ ERRORES:")
            For Each errorMsg In Errores
                mensaje.AppendLine("  • " & errorMsg)
            Next
        End If

        If Advertencias.Count > 0 Then
            mensaje.AppendLine("⚠️ ADVERTENCIAS:")
            For Each advertencia In Advertencias
                mensaje.AppendLine("  • " & advertencia)
            Next
        End If

        Return mensaje.ToString()
    End Function

    Public Sub New()
        EsValido = True
    End Sub
End Class