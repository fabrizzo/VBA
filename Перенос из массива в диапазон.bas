Attribute VB_Name = "Module1"
Option Explicit
Sub RangeMass()
Dim CellsDown As Long, CellsAcross As Integer
Dim i As Long, j As Integer
Dim StartTime As Date
Dim TempArray() As Long
Dim TheRange As Range
Dim CurrVal As Long
CellsDown = 39
CellsAcross = 12
StartTime = Timer
ReDim TempArray(1 To CellsDown, 1 To CellsAcross)
Set TheRange = ActiveCell.Range(Cells(1, 1), Cells(CellsDown, CellsAcross))
CurrVal = 0
Application.ScreenUpdating = False
For i = 1 To CellsDown
For j = 1 To CellsAcross
TempArray(i, j) = CurrVal + 1
CurrVal = CurrVal + 1
Next j
Next i
TheRange.Value = TempArray
Application.ScreenUpdating = True
MsgBox Format(Timer - StartTime, "00.00") & "секунд"
End Sub


