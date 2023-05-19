Private Sub btn_save_Click()
  Range("A9").EntireRow.Insert
  Range("A9").Value = Me.txt_code
  
  Range("B9").Value = Me.txt_process
  Range("C9").Value = Me.regularity
  Range("D9").Value = Me.txt_date
  Range("E9").Value = Me.executed
  Range("F9").Value = Me.txt_note
  
  Unload Evidence_Form
  
  Range("B6").Value = Right(Range("A9").Value, 1)
  End Sub