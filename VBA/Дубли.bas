Attribute VB_Name = "Module1"
Option Explicit
Public vararray() As Variant, i, col_i, str_begin, str_num As Integer, col_letter As String

'========================================================
'============вызываем все остальные процедуры отсюда======
'========================================================
Public Sub list_duplicates()

'------------вводим условия--------------------
i = ActiveSheet.Index
col_i = ActiveCell.Column
col_letter = ActiveCell.Columns.Address(, ColumnAbsolute:=False)
col_letter = Left(col_letter, (InStr(1, col_letter, "$") - 1))

str_begin = val(InputBox("Введите номер строки, с которой начинать поиск...", "Введите значение"))
If str_begin < 1 Then Exit Sub
str_num = val(InputBox("Введите количество строк для обработки...", "Введите значение"))
If str_num < 1 Then Exit Sub

'------------вызываем основные процедуры-------
Call list_unique 'создаем список уникальных значений
Call list_add 'создаем новый лист для вывода отчета
Call duplicates_count 'подсчитываем совпадения и выводим их на лист
End Sub

'========================================================
'============создаем список уникальных значений==========
'========================================================
Private Sub list_unique()
On Error GoTo list_unique_error
Dim k, r As Integer
Dim val_equal As Boolean
Dim pr As Integer

'------------основная часть--------------------
ReDim Preserve vararray(0) 'создаем массив
ReDim Preserve vararray(UBound(vararray) + 1) 'увеличиваем массив
vararray(UBound(vararray) - 1) = Sheets(i).Cells(str_begin, col_i).Value
For k = str_begin + 1 To str_begin + str_num

For r = 1 To k - str_begin 'сравниваем значение ячейки с предыдущим
If Sheets(i).Cells(k, col_i).Value = Sheets(i).Cells(k - r, col_i).Value Then
val_equal = True 'если есть совпадения,
End If
Next r

If Not val_equal = True Then 'не вносим значение в массив
ReDim Preserve vararray(UBound(vararray) + 1) 'меняем размерность массива
vararray(UBound(vararray) - 1) = Sheets(i).Cells(k, col_i).Value
End If

val_equal = False 'обнуляем признак совпадения
Next k
ReDim Preserve vararray(UBound(vararray) - 1) 'уменьшаем размерность массива

Debug.Print "Размерность - " + Trim(Str(UBound(vararray)))
For pr = 0 To UBound(vararray)
Debug.Print vararray(pr)
Next pr
Exit Sub

'------------перехватываем ошибки--------------
list_unique_error:
MsgBox Err.Description
End Sub

'========================================================
'============добавляем лист отчета=======================
'========================================================
Private Sub list_add()
Dim f_i As Integer
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Результат").Delete
Application.DisplayAlerts = True

On Error GoTo list_add_error
Sheets.Add After:=Sheets(Sheets.Count)
Application.ScreenUpdating = True

ActiveSheet.Name = "Результат"

ActiveSheet.Range("A2").Value = " Результат поиска дубликатов значений на листе " + _
Chr(34) + Sheets(i).Name + Chr(34) + " в ячейках " + _
col_letter + Trim(Str(str_begin)) + " - " + col_letter + Trim(Str(str_begin + str_num))
ActiveSheet.Range("C4").Value = "Повторяющиеся значения:"
Range("C4").Font.Bold = True
Range("C4").HorizontalAlignment = xlCenter
ActiveSheet.Range("F4").Value = "Количество дубликатов:"
Range("F4").Font.Bold = True
Range("F4").HorizontalAlignment = xlCenter

For f_i = 6 To str_num - str_begin
Range("C" + Trim(Str(f_i))).HorizontalAlignment = xlCenter
Range("F" + Trim(Str(f_i))).HorizontalAlignment = xlCenter
Next f_i

ActiveSheet.Range("A1").Select

Exit Sub

'------------перехватываем ошибки--------------
list_add_error:
Application.DisplayAlerts = False
Sheets("Результат").Delete 'при ошибке удаляем лист без подтверждений
Application.DisplayAlerts = True
MsgBox "Произошла ошибка при создании нового листа.", vbOKOnly + vbCritical, "Ошибка"
End Sub

'========================================================
'============выводим список на лист отчета===============
'========================================================
Private Sub duplicates_count()
On Error GoTo d_c_error
Dim val As Variant
Dim a_i, res_i, k, counter As Integer

res_i = 6
For a_i = 0 To UBound(vararray)
val = vararray(a_i) 'запоминаем, с чем сравнивать
counter = -1 'т.к. одно совпадение ячеек будет найдено (ячейка сама с собой)
For k = str_begin To (str_begin + str_num)
If val = Sheets(i).Cells(k, col_i).Value Then counter = counter + 1
Next k

If counter > 0 And val <> "" Then
Sheets("Результат").Range("C" + Trim(Str(res_i))).HorizontalAlignment = xlCenter
Sheets("Результат").Range("F" + Trim(Str(res_i))).HorizontalAlignment = xlCenter
Sheets("Результат").Range("C" + Trim(Str(res_i))).Value = val
Sheets("Результат").Range("F" + Trim(Str(res_i))).Value = counter
res_i = res_i + 1
End If

Next a_i

Sheets("Результат").Range("D" + Trim(Str(res_i + 2))).Font.Italic = True
Sheets("Результат").Range("D" + Trim(Str(res_i + 2))).HorizontalAlignment = xlLeft
Sheets("Результат").Range("D" + Trim(Str(res_i + 2))).Value = "Всего дубликатов:"
Sheets("Результат").Range("F" + Trim(Str(res_i + 2))).HorizontalAlignment = xlCenter
Sheets("Результат").Range("F" + Trim(Str(res_i + 2))).Select
ActiveCell.Formula = "=SUM(F6:F" + Trim(Str(res_i)) + ")"
Sheets("Результат").Range("A1").Select
Exit Sub

d_c_error:
MsgBox Err.Description
End Sub '
    Application.Run "Книга1!list_duplicates"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Home\Documents\Определение дублей.xls", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Application.Goto Reference:="list_duplicates"
    ActiveWorkbook.Save

