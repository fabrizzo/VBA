Attribute VB_Name = "Module1"
Sub GetImportFileName()
Dim Filt As String
Dim FilterIndex As Integer
Dim Filename As Variant
Dim Title As String
Filt = "��������� ����� (*.txt) , *.txt, " & _
       "���� Excel 2003 (*.xls) , *.xls, " & _
       "���� Excel 2007 (*.xlsx), *.xlsx," & _
       "��� �����(*.*), *.*"
FilterIndex = 4
Title = "�������� ����"
Filename = Application.GetOpenFilename _
(FileFilter:=Filt, _
 FilterIndex:=FilterIndex, _
 Title:=Title)
If Filename = False Then
MsgBox "�� �� ������� ����"
Exit Sub
End If
MsgBox "�� ������� " & Filename
End Sub
