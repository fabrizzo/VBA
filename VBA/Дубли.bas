Attribute VB_Name = "Module1"
Option Explicit
Public vararray() As Variant, i, col_i, str_begin, str_num As Integer, col_letter As String

'========================================================
'============�������� ��� ��������� ��������� ������======
'========================================================
Public Sub list_duplicates()

'------------������ �������--------------------
i = ActiveSheet.Index
col_i = ActiveCell.Column
col_letter = ActiveCell.Columns.Address(, ColumnAbsolute:=False)
col_letter = Left(col_letter, (InStr(1, col_letter, "$") - 1))

str_begin = val(InputBox("������� ����� ������, � ������� �������� �����...", "������� ��������"))
If str_begin < 1 Then Exit Sub
str_num = val(InputBox("������� ���������� ����� ��� ���������...", "������� ��������"))
If str_num < 1 Then Exit Sub

'------------�������� �������� ���������-------
Call list_unique '������� ������ ���������� ��������
Call list_add '������� ����� ���� ��� ������ ������
Call duplicates_count '������������ ���������� � ������� �� �� ����
End Sub

'========================================================
'============������� ������ ���������� ��������==========
'========================================================
Private Sub list_unique()
On Error GoTo list_unique_error
Dim k, r As Integer
Dim val_equal As Boolean
Dim pr As Integer

'------------�������� �����--------------------
ReDim Preserve vararray(0) '������� ������
ReDim Preserve vararray(UBound(vararray) + 1) '����������� ������
vararray(UBound(vararray) - 1) = Sheets(i).Cells(str_begin, col_i).Value
For k = str_begin + 1 To str_begin + str_num

For r = 1 To k - str_begin '���������� �������� ������ � ����������
If Sheets(i).Cells(k, col_i).Value = Sheets(i).Cells(k - r, col_i).Value Then
val_equal = True '���� ���� ����������,
End If
Next r

If Not val_equal = True Then '�� ������ �������� � ������
ReDim Preserve vararray(UBound(vararray) + 1) '������ ����������� �������
vararray(UBound(vararray) - 1) = Sheets(i).Cells(k, col_i).Value
End If

val_equal = False '�������� ������� ����������
Next k
ReDim Preserve vararray(UBound(vararray) - 1) '��������� ����������� �������

Debug.Print "����������� - " + Trim(Str(UBound(vararray)))
For pr = 0 To UBound(vararray)
Debug.Print vararray(pr)
Next pr
Exit Sub

'------------������������� ������--------------
list_unique_error:
MsgBox Err.Description
End Sub

'========================================================
'============��������� ���� ������=======================
'========================================================
Private Sub list_add()
Dim f_i As Integer
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("���������").Delete
Application.DisplayAlerts = True

On Error GoTo list_add_error
Sheets.Add After:=Sheets(Sheets.Count)
Application.ScreenUpdating = True

ActiveSheet.Name = "���������"

ActiveSheet.Range("A2").Value = " ��������� ������ ���������� �������� �� ����� " + _
Chr(34) + Sheets(i).Name + Chr(34) + " � ������� " + _
col_letter + Trim(Str(str_begin)) + " - " + col_letter + Trim(Str(str_begin + str_num))
ActiveSheet.Range("C4").Value = "������������� ��������:"
Range("C4").Font.Bold = True
Range("C4").HorizontalAlignment = xlCenter
ActiveSheet.Range("F4").Value = "���������� ����������:"
Range("F4").Font.Bold = True
Range("F4").HorizontalAlignment = xlCenter

For f_i = 6 To str_num - str_begin
Range("C" + Trim(Str(f_i))).HorizontalAlignment = xlCenter
Range("F" + Trim(Str(f_i))).HorizontalAlignment = xlCenter
Next f_i

ActiveSheet.Range("A1").Select

Exit Sub

'------------������������� ������--------------
list_add_error:
Application.DisplayAlerts = False
Sheets("���������").Delete '��� ������ ������� ���� ��� �������������
Application.DisplayAlerts = True
MsgBox "��������� ������ ��� �������� ������ �����.", vbOKOnly + vbCritical, "������"
End Sub

'========================================================
'============������� ������ �� ���� ������===============
'========================================================
Private Sub duplicates_count()
On Error GoTo d_c_error
Dim val As Variant
Dim a_i, res_i, k, counter As Integer

res_i = 6
For a_i = 0 To UBound(vararray)
val = vararray(a_i) '����������, � ��� ����������
counter = -1 '�.�. ���� ���������� ����� ����� ������� (������ ���� � �����)
For k = str_begin To (str_begin + str_num)
If val = Sheets(i).Cells(k, col_i).Value Then counter = counter + 1
Next k

If counter > 0 And val <> "" Then
Sheets("���������").Range("C" + Trim(Str(res_i))).HorizontalAlignment = xlCenter
Sheets("���������").Range("F" + Trim(Str(res_i))).HorizontalAlignment = xlCenter
Sheets("���������").Range("C" + Trim(Str(res_i))).Value = val
Sheets("���������").Range("F" + Trim(Str(res_i))).Value = counter
res_i = res_i + 1
End If

Next a_i

Sheets("���������").Range("D" + Trim(Str(res_i + 2))).Font.Italic = True
Sheets("���������").Range("D" + Trim(Str(res_i + 2))).HorizontalAlignment = xlLeft
Sheets("���������").Range("D" + Trim(Str(res_i + 2))).Value = "����� ����������:"
Sheets("���������").Range("F" + Trim(Str(res_i + 2))).HorizontalAlignment = xlCenter
Sheets("���������").Range("F" + Trim(Str(res_i + 2))).Select
ActiveCell.Formula = "=SUM(F6:F" + Trim(Str(res_i)) + ")"
Sheets("���������").Range("A1").Select
Exit Sub

d_c_error:
MsgBox Err.Description
End Sub '
    Application.Run "�����1!list_duplicates"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Home\Documents\����������� ������.xls", FileFormat:=xlNormal, _
        Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Application.Goto Reference:="list_duplicates"
    ActiveWorkbook.Save

