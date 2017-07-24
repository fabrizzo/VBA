Attribute VB_Name = "Очистить_строку"
Option Explicit
Sub CleanRow()
Application.ScreenUpdating = False
Dim numrow As String
Dim row As Integer
Dim LastRow As Long
row = 0
numrow = ""
row = 8
numrow = InputBox("Введите номер строки(число) которую хотите удалить, если ввести * будут удалены все данные", "Удаление строки из всех журналов")
If numrow = vbNullString Then
Exit Sub
End If
If IsNumeric(numrow) Then
    numrow = CInt(numrow)
    If numrow <= 7 Then
    MsgBox "Cтроки начинаются с 8 значения", vbInformation
    Exit Sub
    End If
    Worksheets("Журнал 1").Activate
    Range("A" & numrow & ":" & "AQ" & numrow).ClearContents
    Worksheets("Журнал 2").Activate
    Range("A" & numrow & ":" & "BR" & numrow).ClearContents
    Worksheets("Журнал 3").Activate
    Range("A" & numrow & ":" & "AJ" & numrow).ClearContents
    Worksheets("Журнал 1").Activate
ElseIf numrow = "*" Then
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    Worksheets("Журнал 1").Activate
    Range("A" & row & ":" & "AQ" & LastRow).ClearContents
    Worksheets("Журнал 2").Activate
    Range("A" & row & ":" & "BR" & LastRow).ClearContents
    Worksheets("Журнал 3").Activate
    Range("A" & row & ":" & "AJ" & LastRow).ClearContents
    Worksheets("Журнал 1").Activate
Else
    Exit Sub
End If
End Sub
