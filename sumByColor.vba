Public Function SumByColor(cellColor As Range, sumRange As Range) As Double
  Dim count As Range
  For Each count In sumRange
    If count.Interior.ColorIndex = cellColor.Cells(1, 1).Interior.ColorIndex Then SumByColor = SumByColor + count
  Next count
  Set count = Nothing
End Function

Public Sub Test()
SumByColor(Range("B8").Select, Range("B4:B53").Select)
End Sub


