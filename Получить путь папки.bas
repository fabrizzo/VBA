Attribute VB_Name = "Module1"
Sub getfolder()
With Application.FileDialog(msoFileDialogFolderPicker)
.InitialFileName = Application.DefaultFilePath & "\"
.Title = "Укажите расположение резервной копии"
.Show
If .SelectedItems.Count = 0 Then
MsgBox "отменено"
Else
MsgBox .SelectedItems(1)
End If
End With
End Sub
