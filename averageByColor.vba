Function AverageByColor(cellColor As Range, averageRange As Range) As Double
  Dim count As Range
  Dim amount As Integer
  For Each count In averageRange
    If count.Interior.ColorIndex = cellColor.Cells(1, 1).Interior.ColorIndex Then AverageByColor = AverageByColor + count
    amount = amount + 1
  Next count
  Set count = Nothing
  AverageByColor = AverageByColor / amount
End Function