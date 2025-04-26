Imports System.Data.SQLite
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports BCrypt.Net

Public Class ConexionBD
    ' 🔒 Cadena de conexión a la base de datos SQLite local
    Private Shared ReadOnly cadenaConexion As String = "Data Source=C:\Users\DELL\Dropbox\BD_COOPDIASAM\CONJUNTO_2025.db;Version=3;"

    ' 🌐 Retorna un objeto SQLiteConnection
    Public Shared Function ObtenerConexion() As SQLiteConnection
        Return New SQLiteConnection(cadenaConexion)
    End Function

    ' 🔄 Prueba si hay conexión a la base de datos
    Public Shared Function ProbarConexion() As Boolean
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()
                Return True
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al conectar con la base de datos: " & ex.Message, "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 🔐 Genera el hash SHA256 de un texto (mantenemos esta función igual)
    Public Shared Function GenerarHashSHA256(texto As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(texto)
            Dim hash As Byte() = sha256.ComputeHash(bytes)
            Return BitConverter.ToString(hash).Replace("-", "").ToLower()
        End Using
    End Function






    Public Shared Function ValidarUsuario(usuario As String, contrasena As String) As Boolean
        Try
            Using conexion As SQLiteConnection = ObtenerConexion()
                conexion.Open()

                Dim consulta As String = "SELECT ContrasenaHash FROM Usuarios WHERE NombreUsuario = @usuario"
                Using comando As New SQLiteCommand(consulta, conexion)
                    comando.Parameters.AddWithValue("@usuario", usuario)

                    Dim hashGuardado As Object = comando.ExecuteScalar()

                    If hashGuardado IsNot Nothing AndAlso Not Convert.IsDBNull(hashGuardado) Then
                        Dim hashBcrypt As String = hashGuardado.ToString()
                        Return BCrypt.Net.BCrypt.Verify(contrasena, hashBcrypt)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al validar usuario: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return False
    End Function





End Class