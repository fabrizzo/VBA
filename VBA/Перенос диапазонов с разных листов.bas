Attribute VB_Name = "Module1"
Sub tt()
Dim iLastRow As Integer
 With Worksheets("����3")
  .Cells.Clear
  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("�����").Copy .Cells(iLastRow, 1)
  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("���������_�����").Copy .Cells(iLastRow, 1)
  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
    Range("������").Copy .Cells(iLastRow, 1)
 End With
End Sub

  iLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
  .Cells(iLastRow, 1).Resize(Range("�����").Rows.Count, Range("�����").Columns.Count).Formula = _
  "=����1!" & Replace(Range("�����").Item(1).Address, "$", "")