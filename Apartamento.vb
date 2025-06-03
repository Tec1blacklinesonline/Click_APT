Imports System.Drawing

Public Class Apartamento
    ' Propiedades principales
    Public Property IdApartamento As Integer
    Public Property Torre As Integer
    Public Property Piso As Integer
    Public Property NumeroApartamento As String
    Public Property NombreResidente As String
    Public Property Telefono As String
    Public Property Correo As String
    Public Property Activo As Boolean
    Public Property FechaRegistro As Date

    ' Propiedades calculadas
    Public Property SaldoActual As Decimal
    Public Property EstadoCuenta As String
    Public Property UltimoPago As Date
    Public Property TieneUltimoPago As Boolean
    Public Property DiasEnMora As Integer
    Public Property TotalIntereses As Decimal

    ' Constructor vacío
    Public Sub New()
        Me.Activo = True
        Me.FechaRegistro = Date.Now
        Me.NombreResidente = ""
        Me.Telefono = ""
        Me.Correo = ""
        Me.TieneUltimoPago = False
    End Sub

    ' Constructor con parámetros básicos
    Public Sub New(torre As Integer, piso As Integer, numeroApto As String)
        Me.New()
        Me.Torre = torre
        Me.Piso = piso
        Me.NumeroApartamento = numeroApto
    End Sub

    ' Método para obtener el código completo del apartamento
    Public Function ObtenerCodigoApartamento() As String
        Return $"T{Torre}-{Piso}{NumeroApartamento}"
    End Function

    ' Método para obtener el estado de cuenta
    Public Function ObtenerEstadoCuenta() As String
        If SaldoActual = 0 Then
            Return "Al día"
        ElseIf SaldoActual < 0 Then
            Return "Saldo a favor"
        Else
            Return "Pendiente"
        End If
    End Function

    ' Método para obtener el color según el estado
    Public Function ObtenerColorEstado() As Color
        Select Case ObtenerEstadoCuenta()
            Case "Al día"
                Return Color.Black
            Case "Saldo a favor"
                Return Color.Green
            Case "Pendiente"
                Return Color.Red
            Case Else
                Return Color.Gray
        End Select
    End Function

    ' Método para calcular días en mora
    Public Function CalcularDiasEnMora() As Integer
        If SaldoActual <= 0 OrElse Not TieneUltimoPago Then
            Return 0
        End If

        Dim fechaVencimiento As Date = UltimoPago.AddDays(30) ' Asumiendo vencimiento a 30 días
        If Date.Now > fechaVencimiento Then
            Return (Date.Now - fechaVencimiento).Days
        Else
            Return 0
        End If
    End Function

    ' Método para validar datos del apartamento
    Public Function ValidarDatos() As Boolean
        Return Torre > 0 AndAlso
               Piso > 0 AndAlso
               Not String.IsNullOrEmpty(NumeroApartamento)
    End Function

    ' Método ToString para representación en texto
    Public Overrides Function ToString() As String
        Return $"{ObtenerCodigoApartamento()} - {If(String.IsNullOrEmpty(NombreResidente), "Sin residente", NombreResidente)}"
    End Function
End Class