' ============================================================================
' CUOTAMODEL.VB - VB.NET PURO - NUEVO ARCHIVO
' ============================================================================

Public Class CuotaModel
    ' Propiedades principales
    Public Property IdCuota As Integer
    Public Property IdApartamento As Integer
    Public Property MatriculaInmobiliaria As String
    Public Property FechaCuota As DateTime
    Public Property ValorCuota As Decimal
    Public Property FechaVencimiento As DateTime
    Public Property Estado As String
    Public Property TipoCuota As String
    Public Property TipoPiso As String
    Public Property IdAsamblea As Integer
    Public Property FechaPago As DateTime
    Public Property InteresesMora As Decimal

    ' Constructor vacío
    Public Sub New()
        Estado = "pendiente"
        FechaCuota = DateTime.Now
        TipoCuota = String.Empty
        TipoPiso = String.Empty
        MatriculaInmobiliaria = String.Empty
        InteresesMora = 0
        IdAsamblea = 0
    End Sub

    ' Constructor con parámetros básicos
    Public Sub New(idApartamento As Integer, valorCuota As Decimal, fechaVencimiento As DateTime)
        Me.New()
        Me.IdApartamento = idApartamento
        Me.ValorCuota = valorCuota
        Me.FechaVencimiento = fechaVencimiento
    End Sub

    ' Método para verificar si la cuota está vencida
    Public Function EstaVencida() As Boolean
        Return DateTime.Now.Date > FechaVencimiento.Date AndAlso Estado = "pendiente"
    End Function

    ' Método para calcular días de mora
    Public Function CalcularDiasMora() As Integer
        If EstaVencida() Then
            Return (DateTime.Now.Date - FechaVencimiento.Date).Days
        End If
        Return 0
    End Function

    ' Método para obtener estado formateado
    Public Function ObtenerEstadoFormatado() As String
        Select Case Estado.ToLower()
            Case "pendiente"
                If EstaVencida() Then
                    Return "⏰ Vencida"
                Else
                    Return "⏳ Pendiente"
                End If
            Case "pagada"
                Return "✅ Pagada"
            Case "anulada"
                Return "❌ Anulada"
            Case Else
                Return Estado
        End Select
    End Function

    ' Método para obtener descripción completa
    Public Function ObtenerDescripcion() As String
        Dim descripcion As String = TipoCuota
        If Not String.IsNullOrEmpty(TipoPiso) Then
            descripcion = descripcion & " - " & TipoPiso
        End If
        descripcion = descripcion & " - " & ValorCuota.ToString("C")
        Return descripcion
    End Function

    ' Método ToString sobrescrito
    Public Overrides Function ToString() As String
        Return "Cuota " & IdCuota.ToString() & " - " & ValorCuota.ToString("C") & " - " & Estado
    End Function

End Class