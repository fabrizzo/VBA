Attribute VB_Name = "Module1"
Sub GetImportFileName()
Dim Filt As String
Dim FilterIndex As Integer
Dim Filename As Variant
Dim Title As String
Dim i As Integer
Dim msg As String
Filt = "��������� ����� (*.txt) , *.txt, " & _
       "���� Excel 2003 (*.xls) , *.xls, " & _
       "���� Excel 2007 (*.xlsx), *.xlsx," & _
       "��� �����(*.*), *.*"
FilterIndex = 4
Title = "�������� �����"
Filename = Application.GetOpenFilename _
(FileFilter:=Filt, _
 FilterIndex:=FilterIndex, _
 Title:=Title, _
 MultiSelect:=True)
If IsArray(Filename) = False Then
MsgBox "�� ������� ����"
Else
For i = LBound(Filename) To UBound(Filename)
msg = msg & Filename(i) & vbCrLf
Next i
MsgBox "�� �������:" & vbCrLf & msg
End If
End Sub
