Private Sub UserForm_Activate()
  Me.regularity.AddItem ("Diario")
  Me.regularity.AddItem ("Semanal")
  Me.regularity.AddItem ("Mensual")
  Me.regularity.AddItem ("Puntual")
  d = Format(Date, "YYYY/MM/DD")
  Me.txt_date.Value = d
  d = Replace(d, "/", "")
  Me.txt_code.Value = d
End Sub