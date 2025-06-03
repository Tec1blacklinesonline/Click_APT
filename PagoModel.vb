Public Class PagoModel

    ' Propiedades principales
    Public Property IdPago As Integer
    Public Property IdApartamento As Integer
    Public Property MatriculaInmobiliaria As String
    Public Property IdCuota As Integer?
    Public Property FechaPago As DateTime
    Public Property NumeroRecibo As String

    ' Montos y saldos
    Public Property SaldoAnterior As Decimal
    Public Property PagoAdministracion As Decimal
    Public Property PagoIntereses As Decimal
    Public Property CuotaActual As Decimal
    Public Property TotalPagado As Decimal
    Public Property SaldoActual As Decimal

    ' Información adicional
    Public Property Detalle As String
    Public Property Observaciones As String

    ' Propiedades calculadas y de apoyo
    Public Property NumeroApartamento As String
    Public Property NombreResidente As String
    Public Property Torre As Integer

    ' Constructor vacío
    Public Sub New()
        IdPago = 0
        IdApartamento = 0
        MatriculaInmobiliaria = ""
        IdCuota = Nothing
        FechaPago = DateTime.Now
        NumeroRecibo = ""
        SaldoAnterior = 0
        PagoAdministracion = 0
        PagoIntereses = 0
        CuotaActual = 0
        TotalPagado = 0
        SaldoActual = 0
        Detalle = ""
        Observaciones = ""
        NumeroApartamento = ""
        NombreResidente = ""
        Torre = 0
    End Sub

    ' Constructor con parámetros básicos
    Public Sub New(idApartamento As Integer, fechaPago As DateTime, pagoAdmin As Decimal)
        Me.New()
        Me.IdApartamento = idApartamento
        Me.FechaPago = fechaPago
        Me.PagoAdministracion = pagoAdmin
        CalcularTotales()
    End Sub

    ' Método para calcular totales automáticamente
    Public Sub CalcularTotales()
        ' Total Pagado = Pago Administración + Pago Intereses
        TotalPagado = PagoAdministracion + PagoIntereses

        ' Cuota Actual = Pago Administración + Intereses moratorios (si los hay)
        ' Por ahora asumimos que CuotaActual = PagoAdministracion
        If CuotaActual = 0 Then
            CuotaActual = PagoAdministracion
        End If

        ' Saldo Actual = Saldo Anterior - Total Pagado
        ' (Si hay intereses moratorios se suman al saldo)
        SaldoActual = SaldoAnterior - TotalPagado
    End Sub

    ' Método para generar número de recibo según especificación
    Public Sub GenerarNumeroRecibo()
        If String.IsNullOrEmpty(NumeroRecibo) AndAlso Not String.IsNullOrEmpty(MatriculaInmobiliaria) Then
            ' Formato: matricula + AAMMDD + HHMM
            Dim fechaFormato As String = FechaPago.ToString("yyMMdd")
            Dim horaFormato As String = FechaPago.ToString("HHmm")
            NumeroRecibo = $"{MatriculaInmobiliaria}{fechaFormato}{horaFormato}"
        End If
    End Sub

    ' Método para validar el pago antes de registrar
    Public Function EsValido() As Boolean
        ' Validaciones básicas
        If IdApartamento <= 0 Then Return False
        If PagoAdministracion < 0 Then Return False
        If PagoIntereses < 0 Then Return False
        If FechaPago > DateTime.Now.AddDays(1) Then Return False ' No puede ser futuro

        Return True
    End Function

    ' Método para obtener el estado del pago
    Public Function ObtenerEstadoPago() As String
        If SaldoActual > 0 Then
            Return "Pendiente"
        ElseIf SaldoActual < 0 Then
            Return "Saldo a favor"
        Else
            Return "Al día"
        End If
    End Function

    ' Método para clonar el objeto
    Public Function Clonar() As PagoModel
        Return New PagoModel With {
            .IdPago = Me.IdPago,
            .IdApartamento = Me.IdApartamento,
            .MatriculaInmobiliaria = Me.MatriculaInmobiliaria,
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
            .NumeroApartamento = Me.NumeroApartamento,
            .NombreResidente = Me.NombreResidente,
            .Torre = Me.Torre
        }
    End Function

    ' Override ToString para depuración
    Public Overrides Function ToString() As String
        Return $"Pago {NumeroRecibo} - Apto {NumeroApartamento} - ${TotalPagado:C} - {FechaPago:dd/MM/yyyy}"
    End Function

End Class