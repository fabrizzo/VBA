Attribute VB_Name = "Module1"
Sub GetImportFileName()
Dim Filt As String
Dim FilterIndex As Integer
Dim Filename As Variant
Dim Title As String
Filt = "Текстовые файлы (*.txt) , *.txt, " & _
       "Файл Excel 2003 (*.xls) , *.xls, " & _
       "Файл Excel 2007 (*.xlsx), *.xlsx," & _
       "Все файлы(*.*), *.*"
FilterIndex = 4
Title = "Выберите файл"
Filename = Application.GetOpenFilename _
(FileFilter:=Filt, _
 FilterIndex:=FilterIndex, _
 Title:=Title)
If Filename = False Then
MsgBox "Вы не выбрали файл"
Exit Sub
End If
MsgBox "вы выбрали " & Filename
End Sub
