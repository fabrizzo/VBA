Attribute VB_Name = "Module2"
Option Explicit
Sub DeleteEmpty()
Dim LastRow As Long, r As Long
LastRow = Range("B" & Rows.Count).End(xlUp).Row
Application.ScreenUpdating = False
For r = LastRow To 2 Step -1
If Range("P" & r) = "" Then
Rows(r).Delete
End If
Next r
End Sub
Sub StayOnlyEmpty()
Dim LastRow As Long, r As Long
LastRow = Range("B" & Rows.Count).End(xlUp).Row
Application.ScreenUpdating = False
For r = LastRow To 2 Step -1
If Range("P" & r) <> "" Then
Rows(r).Delete
End If
Next r
End Sub



