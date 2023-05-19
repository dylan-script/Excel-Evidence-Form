Private Sub CommandButton1_Click()

Set explorar_archivo = Application.FileDialog(msoFileDialogFilePicker)
explorar_archivo.Title = "Busca la imagen de Tío Nacho"
explorar_archivo.AllowMultiSelect = False
explorar_archivo.Filters.Add "Archivos Tipo Imagen", "*.jpeg, *.jpg, *.png, *.gif", 1

explorar_archivo.Show

ruta_imagen = explorar_archivo.SelectedItems(1)
Me.txt_ruta.Value = ruta_imagen

Image1.Picture = LoadPicture(explorar_archivo.SelectedItems(1))


End Sub