Private Sub Worksheet_SelectionChange(ByVal Target As Range)

On Error Resume Next
If Not Intersect(Target, Range("G:G")) Is Nothing Then

Range("D2").Value = Target
ruta_screenshot = ThisWorkbook.Path & "/img/" & Target

If Target <> "" And Target <> "No-Img.jpg" Then
Ampliacion.Show
End If

End If

End Sub
