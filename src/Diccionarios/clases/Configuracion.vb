Public Class Configuracion

    Private vMdb As String
    Private vDirRepositorio As String
    Private vDirInstalacion As String
    Private vConexion As String


#Region "Setters y Getters"

    Public Property dirMdb()
        Get
            Return vMdb
        End Get
        Set(ByVal value)
            vMdb = value
        End Set
    End Property

    Public Property dirRepositorio()
        Get
            Return vDirRepositorio
        End Get
        Set(ByVal value)
            vDirRepositorio = value
        End Set
    End Property

    Public Property dirInstalacion()
        Get
            Return vDirInstalacion
        End Get
        Set(ByVal value)
            vDirInstalacion = value
        End Set
    End Property

    Public Property conexion() As String
        Get
            Return vConexion
        End Get
        Set(ByVal value As String)
            vConexion = value
        End Set
    End Property

#End Region

End Class
