Attribute VB_Name = "Module1"
Sub Version()
If Val(Application.Version) <= 12 Then
MsgBox "�� ����������� ���������� Excel"
Else
MsgBox "���������� ������ Microsoft Office ��� ������ � ����������"
End If
End Sub
