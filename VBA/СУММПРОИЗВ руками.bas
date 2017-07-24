Attribute VB_Name = "Module1"
Option Explicit
Sub RangetoVariant()
Dim UserRange As Range
Dim x As Variant
Dim r As Long, c As Integer
Dim mcs%, comvd%, odmvd%, sk%, e%
Dim LastRow As Integer
LastRow = Range("A65536").End(xlUp).Row
Set UserRange = Range("A1:B" & LastRow)
x = UserRange
mcs = 1
comvd = 1
odmvd = 1
sk = 1
e = 1
For c = 1 To 39
If c > 1 Then
mcs = 1
comvd = 1
odmvd = 1
sk = 1
e = 1
End If
For r = 1 To UBound(x, 1)
If x(r, 1) = c And x(r, 2) = "МЧС России" Then
Cells(c, 4).Value = mcs
mcs = mcs + 1
End If
If x(r, 1) = c And x(r, 2) = "СО - МВД России" Then
Cells(c, 5).Value = comvd
comvd = comvd + 1
End If
If x(r, 1) = c And x(r, 2) = "ОД - МВД России" Then
Cells(c, 6).Value = odmvd
odmvd = odmvd + 1
End If
If x(r, 1) = c And x(r, 2) = "СК России" Then
Cells(c, 7).Value = sk
sk = sk + 1
End If
If x(r, 1) = c And x(r, 2) = "ФССП России" Then
Cells(c, 8).Value = e
e = e + 1
End If
Next r
Range("C" & c).Value = c
Next c
Rows("1:1").Select
Selection.Insert Shift:=xlDown
Cells(1, 4).Value = "МЧС России"
Cells(1, 5).Value = "СО - МВД России"
Cells(1, 6).Value = "ОД - МВД России"
Cells(1, 7).Value = "СК России"
Cells(1, 8).Value = "ФССП России"
Columns("D:D").EntireColumn.AutoFit
Columns("E:E").EntireColumn.AutoFit
Columns("F:F").EntireColumn.AutoFit
Columns("G:G").EntireColumn.AutoFit
Columns("H:H").EntireColumn.AutoFit

End Sub
