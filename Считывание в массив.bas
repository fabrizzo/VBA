Attribute VB_Name = "Module3"
Sub test3()
Dim Sheet1_WS, Sheet2_WS As Worksheet
Dim i As Long
Dim R_data As Variant
Dim M_data As Variant

Dim FinalRow, FinalCol As Long
Set Sheet1_WS = Application.ThisWorkbook.Sheets(1)
Set Sheet2_WS = Application.ThisWorkbook.Sheets(2)
FinalRow = Sheet2_WS.Cells(Rows.Count, 1).End(xlUp).Row
FinalCol = Sheet2_WS.Cells(1, Columns.Count).End(xlToLeft).Column
R_data = Sheet2_WS.Range(Sheet2_WS.Cells(1, 1), Sheet2_WS.Cells(FinalRow, FinalCol))
For i = 1 To FinalRow
ReDim M_data(1, 1 To FinalRow)
M_data(1, i) = R_data(i, 1)
Next i
For i = 1 To FinalRow
Debug.Print M_data(i, 1)
Next i
'Sheet1_WS.Range(Sheet1_WS.Cells(1, 1), Sheet1_WS.Cells(FinalRow, FinalCol)) = M_data
End Sub
