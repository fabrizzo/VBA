Attribute VB_Name = "Error_Handler"
Option Explicit

Dim msg As String
On Error GoTo Error_Handler
Exit Sub
Error_Handler:
msg = "������ � [" & Application.VBE.ActiveCodePane.CodeModule & "/��������]" & " ������ #" & Str(Err.Number) & " � ������� [" & Err.Source & "] ��������: " & Err.Description & Chr(13) & "� ������ ������������� ������ ���������a ���������� � ������������"
MsgBox msg, vbInformation, "��!!!"
