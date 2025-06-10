' ============================================================================
' CLASE DAL PARA GESTIÓN DE TORRES
' ============================================================================

Imports System.Data.SQLite



' Clase modelo Torre
Public Class Torre
    Public Property IdTorre As Integer
    Public Property Nombre As String
    Public Property Pisos As Integer
    Public Property TotalApartamentos As Integer

    Public Sub New()
        ' Constructor vacío
    End Sub

    Public Overrides Function ToString() As String
        Return $"Torre {IdTorre} - {Nombre}"
    End Function
End Class