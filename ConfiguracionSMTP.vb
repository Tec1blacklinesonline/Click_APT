Imports System.Xml
Imports System.IO

Public Class ConfiguracionSMTP

    ' Propiedades de configuración
    Public Property ServidorSMTP As String
    Public Property Puerto As Integer
    Public Property UsuarioSMTP As String
    Public Property ContrasenaSMTP As String
    Public Property UsarSSL As Boolean
    Public Property CorreoRemitente As String
    Public Property NombreRemitente As String
    Public Property TimeoutSegundos As Integer

    ' Ruta del archivo de configuración
    Private Shared ReadOnly RutaConfiguracion As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "COOPDIASAM",
        "smtp_config.xml"
    )

    ' Constructor con valores por defecto
    Public Sub New()
        ' Valores por defecto para Gmail
        ServidorSMTP = "smtp.gmail.com"
        Puerto = 587
        UsuarioSMTP = ""
        ContrasenaSMTP = ""
        UsarSSL = True
        CorreoRemitente = "conjuntocoopdiasam@yahoo.es"
        NombreRemitente = "CONJUNTO RESIDENCIAL COOPDIASAM"
        TimeoutSegundos = 30
    End Sub

    ' Cargar configuración desde archivo XML
    Public Shared Function Cargar() As ConfiguracionSMTP
        Dim config As New ConfiguracionSMTP()

        Try
            If File.Exists(RutaConfiguracion) Then
                Dim doc As New XmlDocument()
                doc.Load(RutaConfiguracion)

                Dim root As XmlNode = doc.SelectSingleNode("ConfiguracionSMTP")
                If root IsNot Nothing Then
                    config.ServidorSMTP = root.SelectSingleNode("ServidorSMTP")?.InnerText ?? config.ServidorSMTP
                    config.Puerto = Convert.ToInt32(root.SelectSingleNode("Puerto")?.InnerText ?? config.Puerto.ToString())
                    config.UsuarioSMTP = root.SelectSingleNode("UsuarioSMTP")?.InnerText ?? config.UsuarioSMTP
                    config.ContrasenaSMTP = DesencriptarTexto(root.SelectSingleNode("ContrasenaSMTP")?.InnerText ?? "")
                    config.UsarSSL = Convert.ToBoolean(root.SelectSingleNode("UsarSSL")?.InnerText ?? config.UsarSSL.ToString())
                    config.CorreoRemitente = root.SelectSingleNode("CorreoRemitente")?.InnerText ?? config.CorreoRemitente
                    config.NombreRemitente = root.SelectSingleNode("NombreRemitente")?.InnerText ?? config.NombreRemitente
                    config.TimeoutSegundos = Convert.ToInt32(root.SelectSingleNode("TimeoutSegundos")?.InnerText ?? config.TimeoutSegundos.ToString())
                End If
            End If
        Catch ex As Exception
            ' Si hay error, devolver configuración por defecto
        End Try

        Return config
    End Function

    ' Guardar configuración en archivo XML
    Public Function Guardar() As Boolean
        Try
            ' Crear directorio si no existe
            Dim directorio As String = Path.GetDirectoryName(RutaConfiguracion)
            If Not Directory.Exists(directorio) Then
                Directory.CreateDirectory(directorio)
            End If

            ' Crear documento XML
            Dim doc As New XmlDocument()
            Dim root As XmlElement = doc.CreateElement("ConfiguracionSMTP")
            doc.AppendChild(root)

            ' Agregar elementos
            AgregarElemento(doc, root, "ServidorSMTP", ServidorSMTP)
            AgregarElemento(doc, root, "Puerto", Puerto.ToString())
            AgregarElemento(doc, root, "UsuarioSMTP", UsuarioSMTP)
            AgregarElemento(doc, root, "ContrasenaSMTP", EncriptarTexto(ContrasenaSMTP))
            AgregarElemento(doc, root, "UsarSSL", UsarSSL.ToString())
            AgregarElemento(doc, root, "CorreoRemitente", CorreoRemitente)
            AgregarElemento(doc, root, "NombreRemitente", NombreRemitente)
            AgregarElemento(doc, root, "TimeoutSegundos", TimeoutSegundos.ToString())

            ' Guardar archivo
            doc.Save(RutaConfiguracion)
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ' Método auxiliar para agregar elementos XML
    Private Sub AgregarElemento(doc As XmlDocument, parent As XmlElement, nombre As String, valor As String)
        Dim elemento As XmlElement = doc.CreateElement(nombre)
        elemento.InnerText = valor
        parent.AppendChild(elemento)
    End Sub

    ' Encriptación simple (en producción usar algo más robusto)
    Private Shared Function EncriptarTexto(texto As String) As String
        If String.IsNullOrEmpty(texto) Then Return ""

        Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(texto)
        Return Convert.ToBase64String(bytes)
    End Function

    ' Desencriptación simple
    Private Shared Function DesencriptarTexto(texto As String) As String
        If String.IsNullOrEmpty(texto) Then Return ""

        Try
            Dim bytes As Byte() = Convert.FromBase64String(texto)
            Return System.Text.Encoding.UTF8.GetString(bytes)
        Catch
            Return ""
        End Try
    End Function

    ' Validar configuración
    Public Function EsValida() As Boolean
        Return Not String.IsNullOrEmpty(ServidorSMTP) AndAlso
               Puerto > 0 AndAlso
               Not String.IsNullOrEmpty(UsuarioSMTP) AndAlso
               Not String.IsNullOrEmpty(ContrasenaSMTP) AndAlso
               Not String.IsNullOrEmpty(CorreoRemitente)
    End Function

    ' Obtener configuraciones predefinidas
    Public Shared Function ObtenerConfiguracionesPredefinidas() As Dictionary(Of String, ConfiguracionSMTP)
        Dim configs As New Dictionary(Of String, ConfiguracionSMTP)

        ' Gmail
        configs.Add("Gmail", New ConfiguracionSMTP With {
            .ServidorSMTP = "smtp.gmail.com",
            .Puerto = 587,
            .UsarSSL = True
        })

        ' Yahoo
        configs.Add("Yahoo", New ConfiguracionSMTP With {
            .ServidorSMTP = "smtp.mail.yahoo.com",
            .Puerto = 587,
            .UsarSSL = True
        })

        ' Outlook/Hotmail
        configs.Add("Outlook", New ConfiguracionSMTP With {
            .ServidorSMTP = "smtp-mail.outlook.com",
            .Puerto = 587,
            .UsarSSL = True
        })

        ' Office 365
        configs.Add("Office365", New ConfiguracionSMTP With {
            .ServidorSMTP = "smtp.office365.com",
            .Puerto = 587,
            .UsarSSL = True
        })

        Return configs
    End Function

End Class