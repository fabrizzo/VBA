Attribute VB_Name = "Module1"
Option Explicit

Private Sub Worksheet_Activate()
' ������ ������� 13.12.2010 (�����������)
'�������� ������� ���� "��������", "����������", _
"��������", "����������� �������"
'� ���������� ���� ������
With CommandBars("Cell")
.Controls(1).Enabled = False    '��������
.Controls(2).Enabled = False    '����������
.Controls(3).Enabled = False    '��������
.Controls(4).Enabled = False    '����������� �������
End With
'� ���������� ���� ������
With CommandBars("Column")
.Controls(1).Enabled = False    '��������
End With
'� ���������� ���� ������
With CommandBars("Row")
.Controls(2).Enabled = False    '����������
End With
'"������" �������������� ����
With CommandBars("Worksheet Menu Bar")
.Controls(2).Enabled = False
End With

Exit Sub

'������� �� �������
With CommandBars("Cell")
.Controls(1).Enabled = True    '��������
.Controls(2).Enabled = True    '����������
.Controls(3).Enabled = True    '��������
.Controls(4).Enabled = True    '����������� �������
End With
'� ���������� ���� ������
With CommandBars("Column")
.Controls(1).Enabled = True    '��������
End With
'� ���������� ���� ������
With CommandBars("Row")
.Controls(2).Enabled = True    '����������
End With
'"������" �������������� ����
With CommandBars("Worksheet Menu Bar")
.Controls(2).Enabled = True
End With

End Sub




