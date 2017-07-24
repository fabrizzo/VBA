Attribute VB_Name = "Module1"
Option Explicit

Sub t()
Dim LastRow As Long
Dim Stops As Integer
Dim iRow As Long
iRow = 1
While Range("A" & iRow) <> ""
    If CInt(Left(Range("A" & iRow), 1)) < 3 Then
    Rows(iRow).Delete
    iRow = iRow - 1
    End If
    iRow = iRow + 1
Wend

End Sub
