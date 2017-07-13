Attribute VB_Name = "Module1"
Option Explicit
Sub t()
Dim LastRow As Long
Dim iRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
iRow = 1
While iRow <= LastRow
    If Range("A" & iRow).Value = "" Then
    Rows(iRow).Delete
        If iRow = 1 Then
        Else
            'iRow = iRow - 1
        End If
    Else
    iRow = iRow + 1
    End If
Wend
End Sub
