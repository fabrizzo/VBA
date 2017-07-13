Attribute VB_Name = "Change_Security"
Function GetFilenamesCollection(Optional ByVal Title As String = "Выберите файлы для обработки", _
                                Optional ByVal InitialPath As String = "c:\") As FileDialogSelectedItems
With Application.FileDialog(3)
    .ButtonName = "Загрузить"
    .Title = Title:
    .InitialFileName = InitialPath
If .Show <> -1 Then Exit Function
Set GetFilenamesCollection = .SelectedItems
End With
End Function
Sub RWFiles()
Application.ScreenUpdating = False
Dim Списокфайлов As FileDialogSelectedItems
Dim wb As Workbook
Set Списокфайлов = GetFilenamesCollection("Загрузка журналов", ThisWorkbook.Path)
If Списокфайлов Is Nothing Then
    Exit Sub
End If
For Each File In Списокфайлов
    Workbooks.Open Filename:=File
    Call Change_Rights_And_Security
Next File
End Sub

Function Change_Rights_And_Security()
    ActiveSheet.Unprotect Password:="123321"
    Windows("Образец заполнения ДЛЯ ПЕЧАТИ.xls").Activate
    Selection.Copy
    Workbooks(3).Activate
    Rows("6:6").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    ActiveWindow.FreezePanes = True
    Dim LastRow As Long
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    If LastRow = 11 Then
    Range("A11").FormulaR1C1 = "1"
    ElseIf LastRow = 12 Then
    Range("A11").FormulaR1C1 = "1"
    Range("A12").FormulaR1C1 = "2"
    ElseIf LastRow = 13 Then
    Range("A11").FormulaR1C1 = "1"
    Range("A12").FormulaR1C1 = "2"
    Range("A13").FormulaR1C1 = "3"
    Range("A11:A13").Select
    Selection.AutoFill Destination:=Range("A11:A" & LastRow)
    End If
    Range("A6").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFiltering:=True, Password:="123321"
    ActiveWorkbook.Close SaveChanges:=True

End Function

