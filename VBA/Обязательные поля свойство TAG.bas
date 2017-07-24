Attribute VB_Name = "Module1"
Function EmptyCount2()
Dim ctl As Control
Dim i  As Integer
Dim EmptyCount() As String
i = 0
For Each ctl In UserForm1.Controls
If TypeName(ctl) = "TextBox" Then
If ctl.Tag = "Required" Then
If ctl.Text = "" Then
Count = Count + 1
If Count = 1 Then
ReDim EmptyCount(Count)
End If
EmptyCount(i) = ctl.Name
i = i + 1
End If
End If
End If
Next ctl

For j = 0 To UBound(EmptyCount)
msg = msg & EmptyCount(j) & vbCrLf
Next j
MsgBox "Вы незаполнили след. поля:" & vbCrLf & msg
End Function
