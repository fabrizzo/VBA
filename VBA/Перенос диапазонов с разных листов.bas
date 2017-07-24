Attribute VB_Name = "Module1"
Sub tt()
Dim iLastRow As Integer
 With Worksheets("Лист3")
  .Cells.Clear
  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("Шапка").Copy .Cells(iLastRow, 1)
  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("Табличная_часть").Copy .Cells(iLastRow, 1)
  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("Подвал").Copy .Cells(iLastRow, 1)
 End With
End Sub

  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
  .Cells(iLastRow, 1).Resize(Range("Шапка").Rows.Count, Range("Шапка").Columns.Count).Formula = _
  "=Лист1!" & Replace(Range("Шапка").Item(1).Address, "$", "")