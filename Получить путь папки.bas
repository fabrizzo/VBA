Attribute VB_Name = "Module1"
Sub getfolder()
With Application.FileDialog(msoFileDialogFolderPicker)
.InitialFileName = Application.DefaultFilePath & "\"
.Title = "������� ������������ ��������� �����"
.Show
If .SelectedItems.Count = 0 Then
MsgBox "��������"
Else
MsgBox .SelectedItems(1)
End If
End With
End Sub
