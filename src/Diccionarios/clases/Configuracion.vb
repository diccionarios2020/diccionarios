Option Explicit On
Option Strict On

''' <summary>
''' Configuration class
''' </summary>
Public Class Configuracion

    Private vMdb As String
    Private vDirRepositorio As String
    Private vDirInstalacion As String
    Private vConexion As String

#Region "Setters y Getters"

    Public Property dirMdb() As String
        Get
            Return vMdb
        End Get
        Set(ByVal value As String)
            vMdb = value
        End Set
    End Property

    Public Property dirRepositorio() As String
        Get
            Return vDirRepositorio
        End Get
        Set(ByVal value As String)
            vDirRepositorio = value
        End Set
    End Property

    Public Property dirInstalacion() As String
        Get
            Return vDirInstalacion
        End Get
        Set(ByVal value As String)
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
