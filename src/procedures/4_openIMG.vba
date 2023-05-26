Private Sub btnSelect_Click()
  Dim file_explorer, img_path
  Set file_explorer = Application.FileDialog(msoFileDialogFilePicker)
  file_explorer.Title = "Busca la imagen de la evidencia"
  file_explorer.AllowMultiSelect = False
  file_explorer.Filters.Add "Archivos Tipo Imagen", "*.jpeg, *.jpg, *.gif", 1
  
  file_explorer.Show
  
  img_path = file_explorer.SelectedItems(1)
  Me.txt_path.Value = img_path
  
  Photo.Picture = LoadPicture(img_path)
  End Sub