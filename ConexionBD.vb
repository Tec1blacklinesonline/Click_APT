Imports MySql.Data.MySqlClient

Public Class ConexionBD
    ' Cadena de conexión a la base de datos remota
    Private Shared ReadOnly cadenaConexion As String = "Server=sql5.freesqldatabase.com;Port=3306;Database=sql5772060;Uid=sql5772060;Pwd=N37ZS47X5U;SslMode=Preferred;"

    ' Retorna un objeto MySqlConnection abierto
    Public Shared Function ObtenerConexion() As MySqlConnection
        Return New MySqlConnection(cadenaConexion)
    End Function

    ' Verifica si un usuario existe con las credenciales proporcionadas
    Public Shared Function ValidarUsuario(usuario As String, contraseña As String) As Boolean
        Dim resultado As Boolean = False

        Using conexion As MySqlConnection = ObtenerConexion()
            Try
                conexion.Open()

                Dim consulta As String = "SELECT COUNT(*) FROM Usuarios WHERE NombreUsuario = @usuario AND Contraseña = @contraseña"
                Using comando As New MySqlCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@usuario", usuario)
                    comando.Parameters.AddWithValue("@contraseña", contraseña)

                    Dim count As Integer = Convert.ToInt32(comando.ExecuteScalar())
                    resultado = (count > 0)
                End Using

            Catch ex As Exception
                MessageBox.Show("Error al conectar con la base de datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using

        Return resultado
    End Function
End Class
