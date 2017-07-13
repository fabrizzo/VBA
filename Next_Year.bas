Attribute VB_Name = "Следующий_период"
Option Explicit
Sub New_Year()
Attribute New_Year.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Dim LastRow, FirstRow, NextRow As Long
    Dim Cell As Object
    Dim Cell2 As Object
    Dim Cell3 As Object
    Dim rRange As Range
    Worksheets(1).Activate
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    If LastRow = 7 Then
        Exit Sub
    End If
    'ActiveWorkbook.SaveCopyAs ActiveWorkbook.Path & "\" & "Резервная копия_" & ActiveWorkbook.Name
'***********************************************
    Worksheets(1).Activate
    ActiveSheet.Unprotect Password:="njvrjpghbjle"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=2, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=3, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=4, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=24, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=25, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=26, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AQ$" & LastRow).AutoFilter Field:=28, Criteria1:="0"
    Sheets("Журнал 1").Select
    Sheets("Журнал 1").Copy Before:=Sheets(2)
    Sheets("Журнал 1").Select
    Selection.AutoFilter
    Range("A8:AQ" & LastRow).ClearContents
    Sheets("Журнал 1 (2)").Select
    Range("A8:AQ" & LastRow).Copy
    Sheets("Журнал 1").Select
    Range("A8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Журнал 1 (2)").Delete
    Range("A7:AQ7").AutoFilter
    Range("A8").Select
    ActiveSheet.Protect Password:="njvrjpghbjle"
'***********************************************
    Worksheets(2).Activate
    ActiveSheet.Unprotect Password:="njvrjpghbjle"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=61, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=62, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=63, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=64, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=65, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=66, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=67, Criteria1:="0"
    ActiveSheet.Range("$A$7:$BR$" & LastRow).AutoFilter Field:=70, Criteria1:="0"
    Sheets("Журнал 2").Select
    Sheets("Журнал 2").Copy Before:=Sheets(3)
    Sheets("Журнал 2").Select
    Selection.AutoFilter
    Range("A8:BR" & LastRow).ClearContents
    Sheets("Журнал 2 (2)").Select
    Range("A8:BR" & LastRow).Copy
    Sheets("Журнал 2").Select
    Range("A8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Журнал 2 (2)").Delete
    Range("A7:BR7").AutoFilter
    Range("A8").Activate
    ActiveSheet.Protect Password:="njvrjpghbjle"
'***********************************************
    Worksheets(3).Activate
    ActiveSheet.Unprotect Password:="njvrjpghbjle"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=27, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=28, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=29, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=30, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=31, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=32, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=33, Criteria1:="0"
    ActiveSheet.Range("$A$7:$AJ$2885").AutoFilter Field:=36, Criteria1:="0"
       Sheets("Журнал 3").Select
    Sheets("Журнал 3").Copy Before:=Sheets(4)
    Sheets("Журнал 3").Select
    Selection.AutoFilter
    Range("A8:AJ" & LastRow).ClearContents
    Sheets("Журнал 3 (2)").Select
    Range("A8:AJ" & LastRow).Copy
    Sheets("Журнал 3").Select
    Range("A8").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Журнал 3 (2)").Delete
    Range("A7:AJ7").AutoFilter
    Range("A8").Activate
    ActiveSheet.Protect Password:="njvrjpghbjle"
    
    
    
    
    
    
    
    








    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub
Sub Version()
Dim Ver As String
Ver = Application.Version
Ver = Left(Ver, 2)
If CInt(Ver) = 12 Then
MsgBox "Вы используете MS Office 2007 версия программного обеспечения 2.0 по вопросам программного обеспечения обращаться по тел. 2-44-14-47, 2-44-14-96", vbInformation
ElseIf CInt(Ver) = 11 Then
MsgBox "Вы используете MS Office 2003 версия программного обеспечения 2.0 по вопросам программного обеспечения обращаться по тел. 2-44-14-47, 2-44-14-96", vbInformation
Else
MsgBox "Вы используете неверную версию MS Office для полной работоспособоности программноо обеспечения перейдите на версию MS Office 2003-2007", vbInformation
End If
End Sub
