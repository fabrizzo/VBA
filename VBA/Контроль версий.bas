Attribute VB_Name = "Module1"
Sub Version()
If Val(Application.Version) <= 12 Then
MsgBox "Вы используете правильный Excel"
Else
MsgBox "Установите другой Microsoft Office для работы с программой"
End If
End Sub
