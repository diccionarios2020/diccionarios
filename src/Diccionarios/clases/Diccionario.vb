Option Explicit On
Option Strict On

Public Class Diccionario

    Private vID_Libro As Integer
    Private vTitulo As String
    Private vIdioma As String
    Private vRuta As String
    Private vReferencia As String

#Region "Setters y Getters"

    Public Property id_libro() As Integer
        Get
            Return vID_Libro
        End Get
        Set(ByVal value As Integer)
            vID_Libro = value
        End Set
    End Property

    Public Property titulo() As String
        Get
            Return vTitulo
        End Get
        Set(ByVal value As String)
            vTitulo = value
        End Set
    End Property

    Public Property idioma() As String
        Get
            Return vIdioma
        End Get
        Set(ByVal value As String)
            vIdioma = value
        End Set
    End Property

    Public Property ruta() As String
        Get
            Return vRuta
        End Get
        Set(ByVal value As String)
            vRuta = value
        End Set
    End Property

    Public Property referencia() As String
        Get
            Return vReferencia
        End Get
        Set(ByVal value As String)
            vReferencia = value
        End Set
    End Property

#End Region





End Class
