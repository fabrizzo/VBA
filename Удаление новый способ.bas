Attribute VB_Name = "Module1"
Option Explicit
Sub t()
Dim lCol As Long
Dim LastRow As Long
Dim li As Long
Dim arr As Variant
lCol = 1
LastRow = Range("A" & Rows.Count).End(xlUp).Row
arr = Cells(1, lCol).Resize(LastRow).Value
Application.ScreenUpdating = False
Dim rr As Range
For li = 1 To LastRow
    If arr(li, 1) = "" Then
    Set rr = Cells(li, 1)
    If Not rr Is Nothing Then
    rr.EntireRow.Delete
    li = li - 1
    End If
    End If

Next
 Application.ScreenUpdating = True

End Sub
