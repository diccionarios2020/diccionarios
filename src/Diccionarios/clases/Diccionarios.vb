Imports System.Data
Imports System.Data.OleDb

Public Class Diccionarios

    Private vConexion As String


#Region "Setters y Getters"

    Public Property conexion() As String
        Get
            Return vConexion
        End Get
        Set(ByVal value As String)
            vConexion = value
        End Set
    End Property

#End Region


    Friend Sub cargaDiccionarios(ByRef diccionarios As Collection)
        Dim idioma As String  ' BORRAR!
        Dim cn As New OleDbConnection(Me.conexion)
        Dim MiDataSet As New DataSet
        Dim MiAdaptador As New OleDb.OleDbDataAdapter
        Try
            cn.Open()
            'Creamos la consulta
            Dim consultaSQL As String = "SELECT * FROM Libros WHERE libros.Idioma = '" & idioma & "'"
            ' creamos un comando
            Dim comando As New OleDbCommand(consultaSQL, cn)
            MiAdaptador.SelectCommand = comando
            MiAdaptador.Fill(MiDataSet)

            'For Each dic In MiAdaptador.


        Catch ex As Exception
            'mensaje de error
            MessageBox.Show("error 1" & Err.Number.ToString & vbCrLf & ex.Message)
        Finally
            ' cerrar la conexión si está abierta
            If (cn.State And ConnectionState.Open) <> 0 Then
                cn.Close()
            End If
        End Try
    End Sub



    Friend Sub cargarDiccionarios(ByVal idioma As String, ByRef cbDic As ComboBox, ByRef dgv As DataGridView)
        Dim cn As New OleDbConnection(Me.conexion)
        Dim MiDataSet As New DataSet
        Dim MiAdaptador As New OleDb.OleDbDataAdapter
        Try
            cn.Open()
            'Creamos la consulta
            Dim consultaSQL As String = "SELECT * FROM Libros WHERE libros.Idioma = '" & idioma & "'"
            ' creamos un comando
            Dim comando As New OleDbCommand(consultaSQL, cn)
            MiAdaptador.SelectCommand = comando
            MiAdaptador.Fill(MiDataSet)

            ' Carga la lista de diccionarios en el combobox
            cbDic.DataSource = MiDataSet.Tables(0)
            cbDic.DisplayMember = "Titulo"
            cbDic.ValueMember = "ID_Libro"
            dgv.DataSource = MiDataSet.Tables(0)

        Catch ex As Exception
            'mensaje de error
            MessageBox.Show("error 1" & Err.Number.ToString & vbCrLf & ex.Message)
        Finally
            ' cerrar la conexión si está abierta
            If (cn.State And ConnectionState.Open) <> 0 Then
                cn.Close()
            End If
        End Try
    End Sub

    Friend Sub cargarDiccionario(ByVal unaIdDiccionario As Integer, ByRef dgv As DataGridView)
        Dim cn As New OleDbConnection(Me.conexion)
        Dim MiDataSet As New DataSet
        Dim MiAdaptador As New OleDb.OleDbDataAdapter
        Try
            cn.Open()
            'Creamos la consulta
            Dim sql As String = "SELECT Encabezado, numeroPagina, archivo, primeraPalabra, ultimaPalabra FROM Paginas WHERE Paginas.ID_Libro = " & unaIdDiccionario & " ORDER BY archivo"
            ' creamos un comando
            Dim comando As New OleDbCommand(sql, cn)
            MiAdaptador.SelectCommand = comando
            MiAdaptador.Fill(MiDataSet)
            dgv.DataSource = MiDataSet.Tables(0)
        Catch ex As Exception
            'mensaje de error
            MessageBox.Show("error 2" & Err.Number.ToString & vbCrLf & ex.Message)
        Finally
            ' cerrar la conexión si está abierta
            If (cn.State And ConnectionState.Open) <> 0 Then
                cn.Close()
            End If
        End Try

    End Sub ' Carga un diccionario por su ID_libro




End Class
