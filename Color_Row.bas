Attribute VB_Name = "Color_Row"
Option Explicit
Sub Color()
Dim rRange As Range
Dim LastRow As Long
Dim Cell As Object
Dim Colornum As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
Set rRange = Range("AI1:AI" & LastRow)
For Each Cell In rRange
Select Case Cell.Value
Case "Æ"
    Colornum = 0
    Colornum = Cell.Row
    Range("A" & Colornum & ":AH" & Colornum).Interior.Color = RGB(255, 255, 0)
Case "Ç"
    Colornum = 0
    Colornum = Cell.Row
    Range("A" & Colornum & ":AH" & Colornum).Interior.Color = RGB(0, 255, 0)
Case "Î"
    Colornum = 0
    Colornum = Cell.Row
    Range("A" & Colornum & ":AH" & Colornum).Interior.Color = RGB(255, 128, 0)
Case "Ð"
    Colornum = 0
    Colornum = Cell.Row
    Range("A" & Colornum & ":AH" & Colornum).Interior.Color = RGB(255, 203, 219)
Case "Ô"
    Colornum = 0
    Colornum = Cell.Row
    Range("A" & Colornum & ":AH" & Colornum).Interior.Color = RGB(90, 0, 157)
End Select
Next Cell
End Sub
