Private Sub UserForm_Activate()
  Me.regularity.AddItem ("Diario")
  Me.regularity.AddItem ("Semanal")
  Me.regularity.AddItem ("Mensual")
  Me.regularity.AddItem ("Puntual")
  d = Format(Date, "YYYY/MM/DD")
  Me.txt_date.Value = d
  
  Me.txt_code.Value = Replace(d, "/", "") & Range("B6").Value
  
  End Sub