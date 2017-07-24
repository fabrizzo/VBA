Attribute VB_Name = "Module1"
Option Explicit

Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.AutoFilter
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],""_"",RC[-1])"
    Range("D3").Select
    Columns("D:D").ColumnWidth = 9.57
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D125717")
    Range("D2:D125717").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Fabula").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Fabula").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("D1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Fabula").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Dim LastRow As Long
Dim i As Integer
LastRow = Range("A" & Rows.Count).End(xlUp).Row
For i = 2 To LastRow
If Range("D" & i).Value = Range("D" & i + 1).Value Then
Rows(Range("D" & i).Row).Delete
End If
End Sub
