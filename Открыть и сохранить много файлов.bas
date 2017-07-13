Attribute VB_Name = "Module1"
Sub GetImportFileName()
Dim Filt As String
Dim FilterIndex As Integer
Dim Filename As Variant
Dim Title As String
Dim Title2 As String
Dim i As Integer
Dim msg As String
Filt = "Текстовые файлы (*.txt) , *.txt, " & _
       "Файл Excel 2003 (*.xls) , *.xls, " & _
       "Файл Excel 2007 (*.xlsx), *.xlsx," & _
       "Все файлы(*.*), *.*"
FilterIndex = 4
Title = "Выберите файлы"
Filename = Application.GetOpenFilename _
(FileFilter:=Filt, _
 FilterIndex:=FilterIndex, _
 Title:=Title, _
 MultiSelect:=True)
If IsArray(Filename) = False Then
MsgBox "не выбрали файл"
Else
For i = LBound(Filename) To UBound(Filename)
msg = msg & Filename(i) & vbCrLf
Next i
MsgBox "Вы выбрали:" & vbCrLf & msg
End If
Title2 = "Сохранить как"
Filename = Application.GetSaveAsFilename _
(FileFilter:=Filt, _
 InitialFileName:="Книга1", _
 FilterIndex:=FilterIndex, _
 Title:=Title2)

End Sub
