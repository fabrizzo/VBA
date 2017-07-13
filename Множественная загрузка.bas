Attribute VB_Name = "��������"
Function GetFilenamesCollection(Optional ByVal Title As String = "�������� ����� ��� ���������", _
                                Optional ByVal InitialPath As String = "c:\") As FileDialogSelectedItems
With Application.FileDialog(3)
    .ButtonName = "���������"
    .Title = Title:
    .InitialFileName = InitialPath
If .Show <> -1 Then Exit Function
Set GetFilenamesCollection = .SelectedItems
End With
End Function

Sub UploadMultyFiles()
Application.ScreenUpdating = False
Dim ������������ As FileDialogSelectedItems
Set ������������ = GetFilenamesCollection("�������� ��������", ThisWorkbook.Path)
If ������������ Is Nothing Then
    Exit Sub
End If
For Each File In ������������
    Workbooks.Open Filename:=File
    
Next File
End Sub


