Option Explicit On

Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing.Printing

''' <summary>
''' Main form, entry point
''' </summary>
Public Class frmPrincipal

#Region "DECLARACION DE VARIABLES"

    ' Crea el objeto de configuración
    Dim config As New Configuracion

    ' Crea el objeto que carga los diccionarios
    Dim cargadorDeDics As New cargadorDeDiccionarios

    ' Crea una colección de diccionarios
    Dim diccionarios As New Collection

    ' Variables del movimiento del ratón
    Private m_PanStartPoint As New Point(0, 0)

    ' Variables de la imagen actual
    Dim vZoomActual As Single = 100.0
    Dim vAnchoDePagina As Boolean = True
    Dim vAutoCentrado As Boolean = False

    ' Variables del diccionario actual
    Dim vIdiomaActual As String = "Latín"

    ' Variables de impresión
    Private PrintPageSettings As New PageSettings



#End Region

    ''' <summary>
    ''' Set up the Configuracion object
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub cargarConfiguracion()

        ' Configura el objeto 'config' con los datos del archivo diccionarios.ini
        '-- Add the location of the dictionaries to the config object.
        With config
            If Directory.Exists(My.Application.Info.DirectoryPath & "\Diccionarios\") Then
                .dirRepositorio = My.Application.Info.DirectoryPath & "\Diccionarios\"
            ElseIf Directory.Exists("D:\Proyecto Libros\Diccionarios\") Then
                .dirRepositorio = "D:\Proyecto Libros\Diccionarios\"
            Else
                .dirRepositorio = "D:\Proyecto Libros\Release\Diccionarios\"
            End If

            .dirMdb = "diccionarios.mdb"
            .dirInstalacion = My.Application.Info.DirectoryPath & "\"
            .conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source =" & .dirInstalacion & .dirMdb

        End With
    End Sub

    Private Sub darFormatoDGVIndice()
        dgvIndice.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvIndice.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
        dgvIndice.Columns(2).Visible = False 'Oculta la columna Archivo
        dgvIndice.Columns(3).Visible = False 'Oculta la columna primeraPalabra
        dgvIndice.Columns(4).Visible = False 'Oculta la columna ultimaPalabra
        dgvIndice.Columns(0).HeaderText = "Encabezado"
        dgvIndice.Columns(1).HeaderText = "Pág."
        dgvIndice.Columns(0).DefaultCellStyle.Font = New Font("Times New Roman", 10, FontStyle.Regular)
        dgvIndice.Columns(1).DefaultCellStyle.Font = New Font("Times New Roman", 10, FontStyle.Regular)
        cbDiccionario.Font = New Font("Times New Roman", 10, FontStyle.Regular)
        dgvIndice.Focus()
    End Sub

    Private Sub frmPrincipal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Carga el archivo de configuración (o en su defecto una configuración predeterminada)
        cargarConfiguracion()

        ' Carga de los diccionarios
        With cargadorDeDics
            .conexion = config.conexion
            .cargarDiccionarios(vIdiomaActual, cbDiccionario, diccionarios)
            .cargarHojasDeUnDiccionario(cbDiccionario.SelectedValue, dgvIndice)
            lblDatosBibliograficos.Text = diccionarios.Item(cbDiccionario.SelectedValue).referencia
        End With

        ' Propiedades del panel contenedor del PictureBox (pb1)
        panelImagen.AutoScroll = True
        pb1.SizeMode = PictureBoxSizeMode.AutoSize
        pb1.Cursor = New Cursor(Me.GetType(), "Manito.ico")
        zoomAlAncho()

        'Cargar como un ícono desde el imageList
        'Me.Icon = Icon.FromHandle(CType(ilIconos.Images(0), Bitmap).GetHicon())

        ' Configuración del zoom
        tstbZoom.Text = Format(vZoomActual, "#.00")

        ' Da formato al DataGridView (dgvIndice) del índice de páginas del diccionario seleccionado
        darFormatoDGVIndice()

    End Sub


#Region "SELECCION Y BUSQUEDA EN DICCIONARIO"

    Private Sub cargarDiccionario()
        Dim bp As New Buscador

        '-- Load the dictionary index and search for the word written in it?
        'Se carga el índice del diccionario y se busca la palabra que haya estado escrita en el 

        cargadorDeDics.cargarHojasDeUnDiccionario(cbDiccionario.SelectedValue, dgvIndice)
        lblDatosBibliograficos.Text = diccionarios.Item(cbDiccionario.SelectedValue).referencia
        bp.buscarPalabra(tbBuscar.Text, dgvIndice)
    End Sub

    ''' <summary>
    ''' Load the page
    ''' </summary>
    Private Sub cargarPagina()

        Dim nuevaPagina As New Paginas

        panelImagen.AutoScrollPosition = New Drawing.Point(0, 0)

        Dim indice As Integer
        ' Trata de cargar la imagen de la página
        ' Try to load the page image
        Try
            indice = Me.dgvIndice.SelectedCells.Item(0).RowIndex

            With nuevaPagina
                .dirRepositorio = config.dirRepositorio
                .nombreDirDiccionario = diccionarios.Item(cbDiccionario.SelectedValue).ruta
                .nombreArchivoImagen = Me.dgvIndice.Item("Archivo", indice).Value
                pb1.Image = New Drawing.Bitmap(.urlImagen)
            End With

        Catch ex As Exception
            indice = 0
            MsgBox("Error: " & Err.Number.ToString & vbCrLf & nuevaPagina.urlImagen & " " & ex.Message)
            pb1.Image = New Drawing.Bitmap(My.Application.Info.DirectoryPath & "\deest.bmp")
        End Try

        lblIdLibro.Text = Me.cbDiccionario.SelectedValue

        ' Mueve la imagen al vértice superior izquiero
        ' Move the image to the top left corner
        panelImagen.AutoScrollPosition = New Drawing.Point(0, 0)

    End Sub

    Private Sub cbDiccionario_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDiccionario.SelectionChangeCommitted
        cargarDiccionario()
        dgvIndice.Focus()
    End Sub ' Carga el diccionario seleccionado

    Private Sub dgvIndice_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvIndice.SelectionChanged
        If Me.dgvIndice.SelectedRows.Count > 0 Then
            cargarPagina()
        End If
    End Sub ' Carga la página seleccionada

    Private Sub btnBorrarTBBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorrarTBBuscar.Click
        tbBuscar.Text = ""
    End Sub

#End Region

#Region "DESPLAZAMIENTO"

    Private Sub centrarImagenA(ByVal x As Integer, ByVal y As Integer)
        panelImagen.AutoScrollPosition = New Drawing.Point(x, y)
    End Sub

    Private Sub centrarImagen()
        Dim DeltaY As Integer = ((panelImagen.Height / 2) - (pb1.Height / 2))
        'panelImagen.AutoScrollPosition = New Drawing.Point((DeltaX - panelImagen.AutoScrollPosition.X), (DeltaY - panelImagen.AutoScrollPosition.Y))
        Dim vertice As New Drawing.Point

        Dim DeltaX As Integer = ((panelImagen.Width / 2) - 20 - (pb1.Width / 2))
        If DeltaX > 0 Then
            vertice.X = DeltaX
        Else
            vertice.X = 0
        End If

        If DeltaY > 0 Then
            vertice.Y = DeltaY
        Else
            vertice.Y = 0
        End If

        pb1.Location = vertice
        Me.Text = "DeltaX = " & DeltaX & " ; DeltaY = " & DeltaY
    End Sub

    Private Sub btnCentrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        centrarImagen()
    End Sub

    Private Sub pb1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseDown
        ' Captura el punto inicial
        m_PanStartPoint = New Point(e.X, e.Y)
        pb1.Cursor = New Cursor(Me.GetType(), "Manito2.ico")

    End Sub

    Private Sub pb1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseMove
        'Verifica que el mouse siga presionado mientras se mueve el ratón
        If e.Button = Windows.Forms.MouseButtons.Left Then
            ' Vamos actualizando las coordenadas del mouse
            Dim DeltaX As Integer = (m_PanStartPoint.X - e.X)
            Dim DeltaY As Integer = (m_PanStartPoint.Y - e.Y)

            'Entonces seteamos la nueva posición del autoscroll
            'SIEMPRE pasar enteros positivos al panel de autoScrollPosition metod
            panelImagen.AutoScrollPosition = New Drawing.Point((DeltaX - panelImagen.AutoScrollPosition.X), (DeltaY - panelImagen.AutoScrollPosition.Y))

            lblXPanel.Text = panelImagen.Location.X
            lblYPanel.Text = panelImagen.Location.Y
        End If


    End Sub

    Private Sub pb1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseUp
        pb1.Cursor = New Cursor(Me.GetType(), "Manito.ico")

        If Control.ModifierKeys = Keys.Control Then
            Dim punto As New Point
            punto.X = e.X - 25 '- (panelImagen.Size.Width / 2)
            punto.Y = e.Y - 25 '- (panelImagen.Size.Height / 2)
            centrarImagenA(punto.X, punto.Y)
        End If
    End Sub

    Private Sub pb1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseWheel

        ' Maneja el scroll vertical si no se presiona otra tecla de control
        If e.Delta < 0 Then
            zoomImagen(zoomAnterior(vZoomActual))
        End If
        If e.Delta > 0 Then
            zoomImagen(zoomPosterior(vZoomActual))
        End If

        ' Maneja el scroll horizontal si se manteiene apretada la tecla SHIFT
        If Control.ModifierKeys = Keys.Shift Then
            If e.Delta < 0 Then

            End If
            If e.Delta > 0 Then

            End If
        End If

        ' Maneja el zoom si se mantiene apretada la letra CTRL
        If Control.ModifierKeys = Keys.Control Then
            If e.Delta < 0 Then
                zoomImagen(zoomAnterior(vZoomActual))
            End If
            If e.Delta > 0 Then
                zoomImagen(zoomPosterior(vZoomActual))
            End If
        End If



    End Sub


    Private Sub pb1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pb1.DoubleClick
        ' Dim punto As New Point
        ' punto.X = pb1.Size.Width - (panelImagen.Size.Width / 2)
        ' punto.Y = pb1.Size.Height - (panelImagen.Size.Height / 2)
        '
        '        centrarImagenA(punto.X, punto.Y)
        '        Me.Text = punto.X & " , " & punto.Y
    End Sub

#End Region ' DESPLAZAMIENTO

#Region "ZOOM"


#Region "Funciones del zoom"

    Private Sub zoomImagen(Optional ByVal zoom As Single = 100.0)
        vZoomActual = zoom
        With pb1
            .SizeMode = PictureBoxSizeMode.StretchImage
            .Width = CInt(.Image.Width * (zoom / 100))
            .Height = CInt(.Image.Height * (zoom / 100))
            tstbZoom.Text = Format(vZoomActual, "#.00")
        End With
        'panelImagen.AutoScrollPosition = m_PanStartPoint

        If vAutoCentrado Then
            centrarImagen()
        End If

        vAnchoDePagina = False
    End Sub ' cambia el zoom de la imagen al valor porcentual indicado

    Private Sub tamanoReal()
        zoomImagen(100)
    End Sub ' cambia la imagen al tamaño real en pixeles

    Private Sub zoomAlAncho()
        Try
            Dim ancho As Single
            ancho = ((panelImagen.Width - 20) * 100 / pb1.Image.Width)
            zoomImagen(ancho)
            vAnchoDePagina = True
        Catch ex As Exception

        End Try

    End Sub ' cambia el zoom de la imagen para que entre en el ancho del panel contenedor

    Private Sub zoomPaginaCompleta()
        zoomImagen((panelImagen.Height - 24) * 100 / pb1.Image.Height)
    End Sub ' cambia el zoom de la imagen para que entre la imagen completa en el contenedor


    Private Function zoomPosterior(ByVal actual As Single) As Single
        Dim vZoomCorrecto As Single
        If (actual >= 1) And (actual < 6.25) Then
            vZoomCorrecto = 6.25
        ElseIf (actual >= 6.25) And (actual < 12.5) Then
            vZoomCorrecto = 12.5
        ElseIf (actual >= 12.5) And (actual < 25) Then
            vZoomCorrecto = 25
        ElseIf (actual >= 25) And (actual < 33.33) Then
            vZoomCorrecto = 33.33
        ElseIf (actual >= 33.33) And (actual < 50) Then
            vZoomCorrecto = 50
        ElseIf (actual >= 50) And (actual < 75) Then
            vZoomCorrecto = 75
        ElseIf (actual >= 75) And (actual < 100) Then
            vZoomCorrecto = 100
        ElseIf (actual >= 100) And (actual < 125) Then
            vZoomCorrecto = 125
        ElseIf (actual >= 125) And (actual < 150) Then
            vZoomCorrecto = 150
        ElseIf (actual >= 150) And (actual < 200) Then
            vZoomCorrecto = 200
        ElseIf (actual >= 200) And (actual < 300) Then
            vZoomCorrecto = 300
        ElseIf (actual >= 300) And (actual < 400) Then
            vZoomCorrecto = 400
        ElseIf (actual >= 400) And (actual < 600) Then
            vZoomCorrecto = 600
        ElseIf (actual >= 600) Then
            vZoomCorrecto = 600
        End If

        Return vZoomCorrecto

    End Function

    Private Function zoomAnterior(ByVal actual As Single) As Single
        Dim vZoomCorrecto As Single
        If (actual >= 1) And (actual <= 6.25) Then
            vZoomCorrecto = 1
        ElseIf (actual > 6.25) And (actual <= 12.5) Then
            vZoomCorrecto = 6.25
        ElseIf (actual > 12.5) And (actual <= 25) Then
            vZoomCorrecto = 12.5
        ElseIf (actual > 25) And (actual <= 33.331) Then
            vZoomCorrecto = 25
        ElseIf (actual > 33.33) And (actual <= 50) Then
            vZoomCorrecto = 33.33
        ElseIf (actual > 50) And (actual <= 75) Then
            vZoomCorrecto = 50
        ElseIf (actual > 75) And (actual <= 100) Then
            vZoomCorrecto = 75
        ElseIf (actual > 100) And (actual <= 125) Then
            vZoomCorrecto = 100
        ElseIf (actual > 125) And (actual <= 150) Then
            vZoomCorrecto = 125
        ElseIf (actual > 150) And (actual <= 200) Then
            vZoomCorrecto = 150
        ElseIf (actual > 200) And (actual <= 300) Then
            vZoomCorrecto = 200
        ElseIf (actual > 300) And (actual <= 400) Then
            vZoomCorrecto = 300
        ElseIf (actual > 400) And (actual <= 600) Then
            vZoomCorrecto = 400
        ElseIf (actual <= 1) Then
            vZoomCorrecto = 1
        End If

        Return vZoomCorrecto
    End Function

#End Region

#Region "Eventos del zoom"

    Private Sub zoomMENOS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click, tsbAlejar.Click

        zoomImagen(zoomAnterior(vZoomActual))
    End Sub

    Private Sub zoomMAS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click, tsbAcercar.Click
        zoomImagen(zoomPosterior(vZoomActual))
    End Sub

    Private Sub tbZoom_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            'SendKeys.Send("{TAB}")
            If (tstbZoom.Text > 0) And (tstbZoom.Text < 601) Then
                zoomImagen(tstbZoom.Text)
            End If
        End If
    End Sub

    Private Sub zoomAncho_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerAlAnchoDeLaPáginaToolStripMenuItem.Click, VerAlAnchoDeLaPáginaToolStripMenuItem1.Click
        zoomAlAncho()
    End Sub

    Private Sub VerPáginaCompletaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerPáginaCompletaToolStripMenuItem.Click, VerPáginaCompletaToolStripMenuItem2.Click
        zoomPaginaCompleta()
    End Sub

    Private Sub tstbZoom_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tstbZoom.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            'SendKeys.Send("{TAB}")
            If (tstbZoom.Text > 0) And (tstbZoom.Text < 601) Then
                vZoomActual = tstbZoom.Text
                zoomImagen(vZoomActual)
            End If
        End If
    End Sub ' Valida y ejecuta el zoom indicado en forma porcentual en el textbox tbZoom

    Private Sub tbZoom_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        tstbZoom.Text = Format(vZoomActual, "#.00") ' Luego de que alguien escriba mal un número de zoom, vuelve a poner el zoom actual
    End Sub ' Restaura el valor del textbox tbZoom al valor del zoom actual

    Private Sub VerTamañoRealToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerTamañoRealToolStripMenuItem.Click
        zoomImagen()
    End Sub

    Private Sub VerAlAnchoDePáginaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerAlAnchoDePáginaToolStripMenuItem.Click
        zoomAlAncho()
    End Sub

    Private Sub VerPáginaCompletaToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerPáginaCompletaToolStripMenuItem1.Click
        zoomPaginaCompleta()
    End Sub

#End Region

#Region "Eventos del menu porcentual"

    Private Sub zoom25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click, ToolStripMenuItem22.Click
        zoomImagen(25)
    End Sub

    Private Sub zoom50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click, ToolStripMenuItem24.Click
        zoomImagen(50)
    End Sub

    Private Sub zoom75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click, ToolStripMenuItem25.Click
        zoomImagen(75)
    End Sub

    Private Sub zoom100_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click, ToolStripMenuItem26.Click, TToolStripMenuItem.Click, VerTamañoRealToolStripMenuItem1.Click
        zoomImagen(100)
    End Sub

    Private Sub zoom150_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click, ToolStripMenuItem27.Click
        zoomImagen(150)
    End Sub

    Private Sub zoom200_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click, ToolStripMenuItem28.Click
        zoomImagen(200)
    End Sub

    Private Sub zoom300_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click, ToolStripMenuItem29.Click
        zoomImagen(300)
    End Sub
#End Region ' Zoom botón por porcentaje


#End Region ' ZOOM

#Region "Manejo de archivos"

    Private Sub guardarComo()
        With guardarComoDialog
            .Filter = "Imagen Jpeg(*.jpg)|*.jpg|Imagen Bmp(*.bmp)|*.bmp|Imagen Gif(*.gif)|*.gif|Imagen Tiff(*.tif)|*.tif"
            .Title = "Seleccione donde quiere guardar la imagen de la página actual"
            .InitialDirectory = "C:\Documents and Settings\" & My.User.Name & "\Escritorio"
            .RestoreDirectory = True
            .FileName = ""
        End With

        'Se utiliza el save file y se especifica el tipo de dato a guardar
        If guardarComoDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then 'Si el pulsamos aceptar en la ventanita
            If guardarComoDialog.FileName <> "" Then
                'Si la ruta del archivo del OpenFileDialog es diferente a nada, es decir, si tiene un nombre será que hemos cargado una foto, de lo contrario nos dejaría guardar una foto que realmente no tenemos.
                If guardarComoDialog.FilterIndex = 1 Then 'Si elegimos la extensión jpg
                    pb1.Image.Save(guardarComoDialog.FileName.ToString, System.Drawing.Imaging.ImageFormat.Jpeg) 'Formateamos el Bitmap a Jpeg y lo guardamos
                    'MsgBox("jpeg")
                ElseIf guardarComoDialog.FilterIndex = 2 Then
                    pb1.Image.Save(guardarComoDialog.FileName.ToString, System.Drawing.Imaging.ImageFormat.Bmp)
                    'MsgBox("bmp")
                ElseIf guardarComoDialog.FilterIndex = 3 Then
                    pb1.Image.Save(guardarComoDialog.FileName.ToString, System.Drawing.Imaging.ImageFormat.Gif)
                    'MsgBox("gif")
                ElseIf guardarComoDialog.FilterIndex = 4 Then
                    pb1.Image.Save(guardarComoDialog.FileName.ToString, System.Drawing.Imaging.ImageFormat.Tiff)
                    'MsgBox("tiff")
                    'Con los demás es exactamente igual, pero cambiando el formato.
                End If
            End If
        End If
    End Sub

    Private Sub tsbGuardarComo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbGuardarComo.Click, ExportarToolStripMenuItem.Click
        guardarComo()
    End Sub

#End Region ' MANEJO DE ARCHIVOS

#Region "Impresión"

    Private Sub PrintGraphic(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        'ev.Graphics.DrawImage(Image.FromFile("C:\prueba.jpg"), ev.Graphics.VisibleClipBounds)
        ev.Graphics.DrawImage(pb1.Image, ev.Graphics.VisibleClipBounds)
        ev.HasMorePages = False
    End Sub

    Private Sub tsbImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbImprimir.Click, ImprimirToolStripMenuItem.Click
        Try
            imprimirImagen.DefaultPageSettings = PrintPageSettings
            PrintDialog1.Document = imprimirImagen
            Dim result As DialogResult = PrintDialog1.ShowDialog
            If result = Windows.Forms.DialogResult.OK Then
                AddHandler imprimirImagen.PrintPage, AddressOf Me.PrintGraphic
                'imprimirImagen.PrinterSettings.PrinterName = "Adobe PDF"
                imprimirImagen.Print()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem23.Click
        Try
            PageSetupDialog1.PageSettings = PrintPageSettings
            PageSetupDialog1.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiVistaPrevia.Click, ToolStripButton1.Click
        Try
            imprimirImagen.DefaultPageSettings = PrintPageSettings
            AddHandler imprimirImagen.PrintPage, AddressOf Me.PrintGraphic
            PrintPreviewDialog1.Document = imprimirImagen
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region  ' MANEJO DE LA IMPRESORA

#Region "Debug"

    Private Sub TextBox1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Delta > 1 Then
            MsgBox("Mayor a 1")
        End If
        Me.Text = "estoy arriba"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'panelImagen.AutoScrollPosition = New Drawing.Point(0, 0)
        pb1.Location = New Point(0, 0)
    End Sub

    Private Sub btImpresoras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each printer In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            MessageBox.Show(printer)
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
#End Region ' PROCESOS Y FUNCIONES DE DEBUG

#Region "ELECCIÓN Y CARGA DE DICCIONARIOS POR SU IDIOMA"

    Private Sub cargarDiccionariosDeLatin()
        If DiccionariosDeGriegoToolStripMenuItem.Checked = True Then
            DiccionariosDeLatínToolStripMenuItem.Checked = True
            DiccionariosDeGriegoToolStripMenuItem.Checked = False
            vIdiomaActual = "Latín"
        End If
        cargadorDeDics.cargarDiccionarios(vIdiomaActual, cbDiccionario, diccionarios)
        cargadorDeDics.cargarHojasDeUnDiccionario(cbDiccionario.SelectedValue, dgvIndice)
        zoomAlAncho()
    End Sub

    Private Sub cargarDiccionariosDeGriego()
        If DiccionariosDeLatínToolStripMenuItem.Checked = True Then
            DiccionariosDeGriegoToolStripMenuItem.Checked = True
            DiccionariosDeLatínToolStripMenuItem.Checked = False
            vIdiomaActual = "Griego"
        End If
        cargadorDeDics.cargarDiccionarios(vIdiomaActual, cbDiccionario, diccionarios)
        cargadorDeDics.cargarHojasDeUnDiccionario(cbDiccionario.SelectedValue, dgvIndice)
        zoomAlAncho()
    End Sub

    Private Sub DiccionariosDeLatínToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiccionariosDeLatínToolStripMenuItem.Click
        cargarDiccionariosDeLatin()
        tbBuscar.Text = Nothing
    End Sub

    Private Sub DiccionariosDeGriegoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiccionariosDeGriegoToolStripMenuItem.Click
        cargarDiccionariosDeGriego()
        tbBuscar.Text = Nothing
    End Sub

#End Region

#Region "MANEJO DEL PORTAPAPELES"

    Private Sub copiarReferenciaBibliográfica()
        Clipboard.SetText(lblDatosBibliograficos.Text)
    End Sub

    Private Sub copiarImagenAlPortapapeles()
        Clipboard.SetDataObject(pb1.Image, True)
    End Sub

    Private Sub CopiarPáginaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopiarPáginaToolStripMenuItem.Click
        copiarImagenAlPortapapeles()
    End Sub

    Private Sub tsbCopiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbCopiar.Click
        copiarImagenAlPortapapeles()
    End Sub

    Private Sub CopiarReferenciaBibliográficosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopiarReferenciaBibliográficosToolStripMenuItem.Click
        copiarReferenciaBibliográfica()
    End Sub

#End Region

#Region "Buscar"

    Private Sub tbBuscar_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbBuscar.TextChanged
        Dim bp As New Buscador
        bp.buscarPalabra(tbBuscar.Text, dgvIndice)
    End Sub ' busca la palabra al cambiar el texto del textbox (tbBuscar)

    Private Sub tbBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbBuscar.Click
        tbBuscar.SelectAll()
    End Sub

    Private Sub btnSideBar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSideBar.Click
        With scDiccionario
            If .Panel1Collapsed = True Then
                .Panel1Collapsed = False
            Else
                .Panel1Collapsed = True
            End If
        End With
        If vAnchoDePagina Then
            zoomAlAncho()
        End If

    End Sub ' Oculta o muestra la sección de búsqueda y navegación del diccionario

#End Region ' BUSQUEDA DE PALABRAS

#Region "MODIFICACION DE CARACTERES"

    Private Function castGriego(ByVal unaLetra As Char) As Char
        Dim letraDeSalida As Char
        Select Case unaLetra
            Case "a", "A"
                letraDeSalida = ChrW("&H03B1")
            Case "b", "B"
                letraDeSalida = ChrW("&H03B2")
            Case "c", "C"
                letraDeSalida = ChrW("&H03C7")
            Case "d", "D"
                letraDeSalida = ChrW("&H03B4")
            Case "e", "E"
                letraDeSalida = ChrW("&H03B5")
            Case "f", "F"
                letraDeSalida = ChrW("&H03C6")
            Case "g", "G"
                letraDeSalida = ChrW("&H03B3")
            Case "h", "H"
                letraDeSalida = ChrW("&H03B7")
            Case "i", "I"
                letraDeSalida = ChrW("&H03B9")
            Case "j", "J"
                letraDeSalida = ChrW("&H03C3")
            Case "k", "K"
                letraDeSalida = ChrW("&H03BA")
            Case "l", "L"
                letraDeSalida = ChrW("&H03BB")
            Case "m", "M"
                letraDeSalida = ChrW("&H03BC")
            Case "n", "N"
                letraDeSalida = ChrW("&H03BD")
            Case "o", "O"
                letraDeSalida = ChrW("&H03BF")
            Case "p", "P"
                letraDeSalida = ChrW("&H03C0")
            Case "q", "Q"
                letraDeSalida = ChrW("&H03B8")
            Case "r", "R"
                letraDeSalida = ChrW("&H03C1")
            Case "s", "S"
                letraDeSalida = ChrW("&H03C3")
            Case "t", "T"
                letraDeSalida = ChrW("&H03C4")
            Case "u", "U"
                letraDeSalida = ChrW("&H03C5")
            Case "w", "W"
                letraDeSalida = ChrW("&H03C9")
            Case "x", "X"
                letraDeSalida = ChrW("&H03BE")
            Case "y", "Y"
                letraDeSalida = ChrW("&H03C8")
            Case "z", "Z"
                letraDeSalida = ChrW("&H03B6")
            Case Else
                letraDeSalida = Nothing
        End Select
        Return letraDeSalida
    End Function

    Private Function castLatin(ByVal unaLetra As Char) As Char
        Dim letraDeSalida As Char
        Select Case unaLetra
            Case ""
                letraDeSalida = ""
            Case Else
                letraDeSalida = unaLetra
        End Select
        Return letraDeSalida
    End Function

    Private Sub tbBuscar_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbBuscar.KeyPress

        If Char.IsLetter(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

        If e.KeyChar = Chr(8) Then              ' controla la tecla BACKSPACE
            e.KeyChar = Chr(8)
        ElseIf e.KeyChar = Chr(27) Then         ' hace que la tecla ESC borre el texto
            tbBuscar.Clear()
        ElseIf e.KeyChar = vbCrLf Then
            tbBuscar.AutoCompleteCustomSource.Add(tbBuscar.Text)
        Else
            If vIdiomaActual = "Latín" Then
                e.KeyChar = castLatin(e.KeyChar)
            Else
                e.KeyChar = castGriego(e.KeyChar)
            End If
        End If


    End Sub

#End Region 'CAST DE CARACTERES LATINOS Y GRIEGOS



    Private Sub dgvIndice_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvIndice.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        tbBuscar.Focus()
        tbBuscar.Clear()

        If e.KeyChar = Chr(8) Then              ' controla la tecla BACKSPACE
            e.KeyChar = Chr(8)
        ElseIf e.KeyChar = Chr(27) Then         ' hace que la tecla ESC borre el texto
            tbBuscar.Clear()
        ElseIf e.KeyChar = vbCrLf Then
            tbBuscar.AutoCompleteCustomSource.Add(tbBuscar.Text)
        Else
            If vIdiomaActual = "Latín" Then
                e.KeyChar = castLatin(e.KeyChar)
            Else
                e.KeyChar = castGriego(e.KeyChar)
            End If
        End If

        tbBuscar.AppendText(e.KeyChar)
    End Sub

    Private Sub btnRecordar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        tbBuscar.AutoCompleteCustomSource.Add(tbBuscar.Text)
    End Sub

    Private Sub btnRecordar_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Delta > 0 Then
            zoomImagen(15)
        End If
    End Sub

    Private Sub frmPrincipal_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If vAnchoDePagina Then
            zoomAlAncho()
        End If
    End Sub

    Private Sub scDiccionario_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles scDiccionario.SplitterMoved
        dgvIndice.Focus()

        If vAnchoDePagina Then
            zoomAlAncho()
        End If



    End Sub


    '    Private Sub AutocentradoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        If AutocentradoToolStripMenuItem.Checked = True Then
    '            AutocentradoToolStripMenuItem.Checked = False
    '            vAutoCentrado = False
    '        Else
    '            AutocentradoToolStripMenuItem.Checked = True
    '            vAutoCentrado = True
    '        End If
    '    End Sub
    '



    Private Sub TecladoGriegoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TecladoGriegoToolStripMenuItem.Click
        Dim formulario As New Form
        With formulario
            .Name = "frmTeclado"
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
            .Size = New System.Drawing.Size(606, 224)
            .StartPosition = FormStartPosition.CenterScreen
            .BackgroundImage = New Drawing.Bitmap(Me.GetType(), "keyboard.png")
        End With
        formulario.ShowDialog()
    End Sub

    Private Sub pb1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pb1.MouseClick
        pb1.Focus()
    End Sub

    Private Sub ÍndiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ÍndiceToolStripMenuItem.Click
        System.Windows.Forms.Help.ShowHelp(Me, "diccionarios.chm")
    End Sub

    Private Sub frmPrincipal_HelpButtonClicked(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.HelpButtonClicked
        System.Windows.Forms.Help.ShowHelp(Me, "diccionarios.chm")

    End Sub

    Private Sub tsbAyuda_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAyuda.Click
        System.Windows.Forms.Help.ShowHelp(Me, "diccionarios.chm")
    End Sub



    Private Sub AcercaDeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AcercaDeToolStripMenuItem.Click
        Dim formulario As New Form
        With formulario
            .Name = "frmTeclado"
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
            .Size = New System.Drawing.Size(300, 200)
            .StartPosition = FormStartPosition.CenterScreen
            .BackgroundImage = New Drawing.Bitmap(Me.GetType(), "keyboard.png")
        End With
        formulario.ShowDialog()

    End Sub
End Class

