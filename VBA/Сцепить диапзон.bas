Attribute VB_Name = "Concatenate_Range"
Option Explicit
Public Function ���������������(ByRef �������� As Excel.Range, Optional ByVal ����������� As String = "") As String
Dim rCell As Range
Dim MergeText As String
For Each rCell In ��������
    If rCell.Text <> "" Then
    MergeText = MergeText & ����������� & rCell.Text
    End If
Next
MergeText = Mid(MergeText, Len(�����������) + 1)
��������������� = MergeText
End Function

