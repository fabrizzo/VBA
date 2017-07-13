Attribute VB_Name = "Module1"
Sub text()
Dim X As Variant
Dim nRange As Range
Dim LastRows As Integer
Dim StatTime As Date, EndTime As Date
StartTime = Timer
For j = 1 To 3000
Sheets("Лист2").Select
Range("B1:B40").Value = Range("A1:A40").Value
Set nRange = Range("B1:B40")
ReDim X(1 To nRange.Rows.Count, 1)
X = nRange
Sheets("Лист1").Select
If Range("A1").Value = "" Then
Range(Cells(1, 1), Cells(1, nRange.Rows.Count)) = X
Else
LastRows = Range("A65536").End(xlUp).Row
LastRows = LastRows + 1
For i = 1 To nRange.Rows.Count
Cells(LastRows, i) = X(i, 1)
Next i
End If
Next j
EndTime = Timer
MsgBox Format(EndTime - StartTime, "0.000")
End Sub


