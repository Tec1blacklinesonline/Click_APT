' ============================================================================
' MODELO DE PAGO
' Representa un pago realizado en el sistema
' ============================================================================

Public Class PagoModel
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

    ' Constructor vacío
    Public Sub New()
        EstadoPago = "REGISTRADO"
        FechaPago = DateTime.Now
        FechaRegistro = DateTime.Now
        Detalle = String.Empty
        Observaciones = String.Empty
    End Sub

    ' Constructor con parámetros básicos
    Public Sub New(idApartamento As Integer, numeroRecibo As String, totalPagado As Decimal)
        Me.New()
        Me.IdApartamento = idApartamento
        Me.NumeroRecibo = numeroRecibo
        Me.TotalPagado = totalPagado
    End Sub

    ' Método para validar el pago
    Public Function Validar() As Boolean
        If IdApartamento <= 0 Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(NumeroRecibo) Then
            Return False
        End If

        If TotalPagado <= 0 Then
            Return False
        End If

        Return True
    End Function

    ' Método para obtener descripción del pago
    Public Function ObtenerDescripcion() As String
        Dim descripcion As String = $"Recibo: {NumeroRecibo}"

        If PagoAdministracion > 0 Then
            descripcion &= $" | Administración: {PagoAdministracion:C}"
        End If

        If PagoIntereses > 0 Then
            descripcion &= $" | Intereses: {PagoIntereses:C}"
        End If

        Return descripcion
    End Function

End Class