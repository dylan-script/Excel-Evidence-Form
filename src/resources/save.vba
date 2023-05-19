Private Sub CommandButton2_Click()

Range("A5").EntireRow.Insert
Range("C5").Value = Me.txt_nombre
Range("D5").Value = Me.txt_apellido
Range("E5").Value = Me.txt_sexo
Range("F5").Value = Me.txt_edad
codigo = Range("C2").Value
Range("B5").Value = codigo

If Me.txt_ruta <> "" Then

Set tech = VBA.CreateObject("Scripting.FileSystemObject")
    origen = Me.txt_ruta.Value
    RutArchivo = ThisWorkbook.Path & "/"
    destino = RutArchivo & "img/" & codigo & ".jpg"
    tech.CopyFile origen, destino
    Range("G5").Value = codigo & ".jpg"
    
Else
    Range("G5").Value = "No-Img.jpg"

End If

Unload formulario_registro

End Sub