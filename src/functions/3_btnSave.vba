Private Sub btn_save_Click()
  Range("A9").EntireRow.Insert
  
  Range("B6").Value = Me.txt_code
  Range("A9").Value = Me.txt_code
  Range("B9").Value = Me.txt_process
  Range("C9").Value = Me.regularity
  Range("D9").Value = Me.txt_date
  Range("E9").Value = Me.executed
  Range("F9").Value = Me.txt_note
  
  If Me.txt_path <> "" Then

    Set tech = VBA.CreateObject("Scripting.FileSystemObject")
        origin = Me.txt_path.Value
        img_path = ThisWorkbook.Path & "/"
        target_path = img_path & "img/" & codigo & ".jpg"
        tech.CopyFile origen, target_path
        Range("G5").Value = codigo & ".jpg"
        
    Else
        Range("G5").Value = "No-Img.jpg"
    End If

  Unload Evidence_Form
  End Sub