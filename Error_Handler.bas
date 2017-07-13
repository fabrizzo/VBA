Attribute VB_Name = "Error_Handler"
Option Explicit

Dim msg As String
On Error GoTo Error_Handler
Exit Sub
Error_Handler:
msg = "Ошибка в [" & Application.VBE.ActiveCodePane.CodeModule & "/Выгрузка]" & " Ошибка #" & Str(Err.Number) & " в проекте [" & Err.Source & "] описание: " & Err.Description & Chr(13) & "В случае возникновения ошибки пожалуйстa обратитесь к разработчику"
MsgBox msg, vbInformation, "Ой!!!"
