<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class imagen
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.imagenSeleccionada = New System.Windows.Forms.PictureBox
        CType(Me.imagenSeleccionada, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'imagenSeleccionada
        '
        Me.imagenSeleccionada.Location = New System.Drawing.Point(3, 3)
        Me.imagenSeleccionada.Margin = New System.Windows.Forms.Padding(0)
        Me.imagenSeleccionada.Name = "imagenSeleccionada"
        Me.imagenSeleccionada.Size = New System.Drawing.Size(225, 225)
        Me.imagenSeleccionada.TabIndex = 0
        Me.imagenSeleccionada.TabStop = False
        '
        'imagen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.AutoScrollMinSize = New System.Drawing.Size(2500, 0)
        Me.Controls.Add(Me.imagenSeleccionada)
        Me.Margin = New System.Windows.Forms.Padding(0)
        Me.Name = "imagen"
        Me.Size = New System.Drawing.Size(183, 183)
        CType(Me.imagenSeleccionada, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents imagenSeleccionada As System.Windows.Forms.PictureBox

End Class
