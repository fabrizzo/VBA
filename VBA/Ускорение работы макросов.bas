Attribute VB_Name = "Module2"
Sub Prepare()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
ActiveSheet.DisplayPageBreaks = False
End Sub
Sub Ended()
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ActiveSheet.DisplayPageBreaks = True
End Sub
