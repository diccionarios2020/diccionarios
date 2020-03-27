Public Class Paginas
    Private vDirRepositorio As String
    Private vNombreDirDiccionario As String
    Private vNombreArchivoImagen As String



#Region "Setters y Getters"

    Public Property dirRepositorio() As String
        Get
            Return vDirRepositorio
        End Get
        Set(ByVal value As String)
            vDirRepositorio = value
        End Set
    End Property

    Public Property nombreDirDiccionario() As String
        Get
            Return vNombreDirDiccionario
        End Get
        Set(ByVal value As String)
            vNombreDirDiccionario = value
        End Set
    End Property

    Public Property nombreArchivoImagen() As String
        Get
            Return vNombreArchivoImagen
        End Get
        Set(ByVal value As String)
            vNombreArchivoImagen = value
        End Set
    End Property

#End Region


#Region "Accessing"

    Public ReadOnly Property urlImagen() As String
        Get
            Dim url = Me.dirRepositorio & Me.nombreDirDiccionario & "\" & Me.nombreArchivoImagen
            Return url
        End Get
    End Property

#End Region


End Class
