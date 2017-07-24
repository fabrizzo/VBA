Attribute VB_Name = "Module2"
Option Explicit
Sub FastKill()
Dim LastRow As Long
Dim i As Integer
LastRow = Range("C" & Rows.Count).End(xlUp).Row
For i = LastRow To 1 Step -1
If Rows(i).Hidden Then
Rows(i).Delete
End If
Next i
End Sub
