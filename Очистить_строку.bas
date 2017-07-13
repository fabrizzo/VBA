Attribute VB_Name = "��������_������"
Option Explicit
Sub CleanRow()
Dim numrow As String: Dim row As Integer: Dim LastRow As Long
Application.ScreenUpdating = False
row = 0: numrow = "": row = 8
numrow = InputBox("������� ����� ������(�����) ������� ������ �������, ����� ������� ��� ������ �������:  *", "�������� ������ �� ���� ��������") '���������� ����� ������ ��� ��� ������
If numrow = vbNullString Then
Exit Sub
End If
If IsNumeric(numrow) Then
    numrow = CInt(numrow) '������� ����� ������
    If numrow <= 7 Then
        MsgBox "C����� ���������� � 8 ��������", vbInformation
        Exit Sub
    End If
    Worksheets("������ 1").Activate: Range("A" & numrow & ":" & "AQ" & numrow).ClearContents '��������� �������
    Worksheets("������ 2").Activate: Range("A" & numrow & ":" & "BR" & numrow).ClearContents
    Worksheets("������ 3").Activate: Range("A" & numrow & ":" & "AJ" & numrow).ClearContents
    Worksheets("������ 1").Activate
ElseIf numrow = "*" Then
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    Worksheets("������ 1").Activate: Range("A" & row & ":" & "AQ" & LastRow).ClearContents '������� ��� ������
    Worksheets("������ 2").Activate: Range("A" & row & ":" & "BR" & LastRow).ClearContents
    Worksheets("������ 3").Activate: Range("A" & row & ":" & "AJ" & LastRow).ClearContents
    Worksheets("������ 1").Activate
Else
    Exit Sub
End If
Application.Calculate '������������� ����������
End Sub
