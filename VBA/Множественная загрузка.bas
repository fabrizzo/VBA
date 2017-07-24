Attribute VB_Name = "Загрузка"
Function GetFilenamesCollection(Optional ByVal Title As String = "Выберите файлы для обработки", _
                                Optional ByVal InitialPath As String = "c:\") As FileDialogSelectedItems
With Application.FileDialog(3)
    .ButtonName = "Загрузить"
    .Title = Title:
    .InitialFileName = InitialPath
If .Show <> -1 Then Exit Function
Set GetFilenamesCollection = .SelectedItems
End With
End Function

Sub UploadMultyFiles()
Application.ScreenUpdating = False
Dim Списокфайлов As FileDialogSelectedItems
Set Списокфайлов = GetFilenamesCollection("Загрузка журналов", ThisWorkbook.Path)
If Списокфайлов Is Nothing Then
    Exit Sub
End If
For Each File In Списокфайлов
    Workbooks.Open Filename:=File
    
Next File
End Sub


