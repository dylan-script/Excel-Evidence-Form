Public taskN As Integer
Public imgTaskN As Integer


Private Sub CommandButton1_Click()
Range("A5").EntireRow.Insert
' Range("D3).Value =
d = Format(Date, "YYYY/MM/DD")
d = Replace(d, "/", "")
MsgBox "Today's Date is " & d
End Sub
