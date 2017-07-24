Attribute VB_Name = "Module1"
Option Explicit
Sub RangetoVariant()
Dim UserRange As Range
Dim x As Variant
Dim r As Long, c As Integer
Dim a As Integer
Set UserRange = Range("A1:B15")
x = UserRange
a = 1
For c = 1 To 39
If c > 1 Then
a = 1
End If
For r = 1 To UBound(x, 1)
If x(r, 1) = c And x(r, 2) = "укрыто" Then
Cells(c, 4).Value = a
a = a + 1
End If
Next r
Next c
End Sub

