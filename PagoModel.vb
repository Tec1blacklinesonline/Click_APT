Public Class PagoModel
    ' Propiedades principales de la tabla pagos
    Public Property IdPago As Integer
    Public Property IdApartamento As Integer
    Public Property MatriculaInmobiliaria As String
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

    ' Propiedades adicionales para mostrar información
    Public Property NumeroApartamento As String
    Public Property NombreResidente As String
    Public Property Torre As Integer

    ' Constructor vacío
    Public Sub New()
        Me.FechaPago = DateTime.Now
        Me.NumeroRecibo = ""
        Me.MatriculaInmobiliaria = ""
        Me.SaldoAnterior = 0
        Me.PagoAdministracion = 0
        Me.PagoIntereses = 0
        Me.CuotaActual = 0
        Me.TotalPagado = 0
        Me.SaldoActual = 0
        Me.Detalle = ""
        Me.Observaciones = ""
        Me.NumeroApartamento = ""
        Me.NombreResidente = ""
        Me.Torre = 0
    End Sub

    ' Constructor con parámetros básicos
    Public Sub New(idApartamento As Integer, numeroRecibo As String)
        Me.New()
        Me.IdApartamento = idApartamento
        Me.NumeroRecibo = numeroRecibo
    End Sub

    ' Método para validar datos del pago
    Public Function ValidarDatos() As Boolean
        Return IdApartamento > 0 AndAlso
               Not String.IsNullOrEmpty(NumeroRecibo) AndAlso
               FechaPago <= DateTime.Now AndAlso
               TotalPagado >= 0
    End Function

    ' Método ToString para representación en texto
    Public Overrides Function ToString() As String
        Return $"Recibo: {NumeroRecibo} - Apartamento: {NumeroApartamento} - Total: {TotalPagado:C}"
    End Function

End Class