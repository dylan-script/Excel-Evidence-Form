Private Sub UserForm_Activate()

On Error Resume Next
ruta_screenshot = ThisWorkbook.Path & "/img/" & Range("D2").Value
screenshot.Picture = LoadPicture(ruta_screenshot)
End Sub
