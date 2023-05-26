Private Sub btn_save_Click()
  Range("A9").EntireRow.Insert
  Range("A9").Value = Me.txt_code
  
  Range("B9").Value = Me.txt_process
  Range("C9").Value = Me.regularity
  Range("D9").Value = Me.txt_date
  Range("E9").Value = Me.executed
  Range("F9").Value = Me.txt_note
  Dim code As String
  code = Range("B6").Value

  If Me.txt_path <> "" Then
    Set evidence = CreateObject("Scripting.FileSystemObject")
    origin = Me.txt_path.Value
    filePath = ThisWorkbook.Path & "/"
    destination = filePath & "img/" & code & ".jpg"
    evidence.CopyFile origin, destination
    Range("G9").Value = code & ".jpg"
  Else
    Range("G9").Value = "No-Img.jpg"
  End If
  Unload Evidence_Form
  End Sub