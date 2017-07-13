Attribute VB_Name = "Analize"
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
Sub RWFiles()
Dim Списокфайлов As FileDialogSelectedItems
Dim wb As Workbook
Application.ScreenUpdating = False
Set Списокфайлов = GetFilenamesCollection("Загрузка журналов", ThisWorkbook.Path)
If Списокфайлов Is Nothing Then
    Exit Sub
End If
For Each file In Списокфайлов
    Call Copy_Function(file)
    For Each wb In Workbooks
        If Not wb Is Workbooks("Анализатор выгрузки.xlsm") Then
        wb.Close SaveChanges:=False
    End If
    Next wb
Next file
UserForm1.TextBox1.Text = "Выгрузка идентична!"
Call Main
End Sub
Function Copy_Function(ByVal file As String)
  Dim LastRow As Long
  Workbooks.OpenText Filename:=file, _
  Origin:=866, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 2), Array(7, 1)), TrailingMinusNumbers:=True
  Range("A1").Select
  Range(Selection, Range("A1").End(xlDown)).Copy
  If UCase(Left(ActiveWorkbook.Name, 3)) = "PRE" Then
  Workbooks("Анализатор выгрузки.xlsm").Activate
      If Range("A1").Value = "" Then
          Range("A1").PasteSpecial
          Application.CutCopyMode = False
      Else
          LastRow = 0
          LastRow = Range("A" & Rows.Count).End(xlUp).Row
          Range("A" & LastRow + 1).PasteSpecial
          Application.CutCopyMode = False
      End If
  ElseIf UCase(Left(ActiveWorkbook.Name, 3)) = "CUR" Then
  Workbooks("Анализатор выгрузки.xlsm").Activate
      If Range("B1").Value = "" Then
      Range("B1").PasteSpecial
      Application.CutCopyMode = False
      Else
      LastRow = 0
      LastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("B" & LastRow + 1).PasteSpecial
      Application.CutCopyMode = False
      End If
  End If
End Function
Function Main()
Dim LastRow As Integer
LastRow = Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To LastRow
If Range("A" & i).Value = Range("B" & i).Value Then
Range("C" & i).Value = 1
Range("C" & i).Interior.Color = QBColor(10)
Else
Range("C" & i).Value = 0
Range("C" & i).Interior.Color = QBColor(12)
End If
Next i

For j = 1 To LastRow
If Range("C" & j).Value = 0 Then
UserForm1.TextBox1.Text = "Выгрузка не идентична!"
End If
Next j
End Function

