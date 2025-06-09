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
    Public Property MatriculaInmobiliaria As String

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
        Me.MatriculaInmobiliaria = ""
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

    ''' <summary>
    ''' Calcula los días en mora para el apartamento, basándose en la fecha de vencimiento
    ''' de la cuota pendiente más antigua.
    ''' </summary>
    ''' <param name="fechaActual">La fecha actual para el cálculo.</param>
    ''' <returns>El número de días en mora (entero). Retorna 0 si no hay mora o no hay cuotas pendientes.</returns>
    Public Function CalcularDiasEnMora(fechaActual As Date) As Integer
        ' Si el saldo actual es cero o a favor, no hay mora (ya no hay deuda a cobrar)
        If SaldoActual <= 0 Then
            Return 0
        End If

        ' Obtener la cuota pendiente más antigua del DAL
        Dim cuotaPendiente As CuotasDAL.CuotaPendienteInfo = CuotasDAL.ObtenerCuotaPendienteMasAntigua(Me.IdApartamento)

        If cuotaPendiente.ExisteCuotaPendiente Then
            ' Si la fecha actual es posterior a la fecha de vencimiento de la cuota pendiente
            If fechaActual > cuotaPendiente.FechaVencimiento Then
                Return (fechaActual - cuotaPendiente.FechaVencimiento).Days
            Else
                Return 0 ' La cuota no ha vencido aún
            End If
        Else
            Return 0 ' No hay cuotas pendientes y vencidas para este apartamento
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