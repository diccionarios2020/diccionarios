﻿Public Class Buscador

    Friend Sub buscarPalabra(ByVal unaPalabra As String, ByRef unDGV As DataGridView)
        unaPalabra.ToLower() ' Pasa el string de la palabra a buscar a minúscula
        For Each row As DataGridViewRow In unDGV.Rows
            Dim primeraPalabra As String = row.Cells("primeraPalabra").Value.ToString().ToLower()
            Dim ultimaPalabra As String = row.Cells("ultimaPalabra").Value.ToString().ToLower()
            'If primeraPalabra <= palabra And ultimaPalabra >= palabra Then
            If ultimaPalabra >= unaPalabra Then
                row.Selected = True
                unDGV.CurrentCell = unDGV.Rows(row.Index).Cells(0)
                Exit For
            End If
        Next
    End Sub

End Class
