Function CountCellsByColorAndValue(data_range As Range, cell_ref As Range) As Long
  Dim indRefColor As Long
  Dim indRefValue As Long
  Dim cellCurrent As Range
  Dim cntRes As Long

Application.Volatile
  cntRes = 0
  indRefColor = cell_ref.Cells(1, 1).Interior.Color
  indRefValue = cell_ref.Cells(1, 1).Value
  For Each cellCurrent In data_range
    If indRefColor = cellCurrent.Interior.Color And indRefValue = cellCurrent.Value Then
      cntRes = cntRes + 1
    End If
  Next cellCurrent

CountCellsByColorAndValue = cntRes
End Function
