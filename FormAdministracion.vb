Public Class FormAdministracion
    Private panelPrincipal As Panel


    Private Sub MenuTorres_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Aquí se mostrará la gestión de torres", "Torres")
        ' Aquí iría la lógica para cargar la vista de torres
    End Sub

    Private Sub MenuPropietarios_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Aquí se gestionarán los propietarios", "Propietarios")
    End Sub

    Private Sub MenuPagos_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Aquí se registrarán y visualizarán los pagos", "Pagos")
    End Sub

    Private Sub MenuReportes_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Aquí se generarán los reportes del conjunto", "Reportes")
    End Sub
End Class
