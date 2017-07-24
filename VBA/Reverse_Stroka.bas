Attribute VB_Name = "Reverse_Stroka"
Option Explicit

Sub str_reverse()
Dim Cell As Object
Dim rRange As Range
Dim LastRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
'Set rRange = Range("A2:A" & LastRow)
Set rRange = Range("C2:C" & LastRow)
For Each Cell In rRange

'If IsNumeric(Right(Cell.Value, 3)) Then
'Range("B" & Cell.Row).Value = Right(Cell.Value, 3)
'ElseIf Mid(Right(Cell.Value, 3), 2, 1) = "." Then
'    Range("B" & Cell.Row).Value = CStr(Right(Cell.Value, 5))
'End If

'If IsNumeric(Left(Cell.Value, 1)) Then
'Range("D" & Cell.Row).Value = Left(Cell.Value, 1)
'End If

'If Left(Cell.Value, 1) = "÷" Then
'Range("D" & Cell.Row).Value = Mid(Cell.Value, 2, 1)
'Rows(Cell.Row + 1).Select
'Selection.Insert Shift:=xlDown
'Range("B" & Cell.Row + 1).Value = Range("B" & Cell.Row).Value
'Range("D" & Cell.Row + 1).Value = Mid(Cell.Value, 4, 1)
'End If

'If Mid(Cell.Value, 4, 1) = "." Then
'Range("B" & Cell.Row).Value = Right(Cell.Value, 1)
'Else
'Range("B" & Cell.Row).Value = 0
'End If

If Cell.Value = "" Then
Cell.Value = 1
End If



Next Cell


End Sub
