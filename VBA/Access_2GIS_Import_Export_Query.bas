Attribute VB_Name = "Import/Export_Query"
Option Compare Database
Option Explicit
Sub export_ALL_Query_to_DBF()
'������������ ��� ������� ������� DBF � ����� DBF
Dim i As Integer
Dim j As Integer
Dim rst As Integer
i = 1
j = 15
For i = 1 To Application.CurrentDb.QueryDefs.Count - 3
If Dir(Application.CurrentProject.Path & "\DBF\�����" & " " & j & ".dbf") = "" Then
rst = DCount("*", "�����" & " " & j)
If rst <> 0 Then
DoCmd.TransferDatabase acExport, "dBase IV", Application.CurrentProject.Path & "\DBF", acQuery, "�����" & " " & j, "�����" & j & ".dbf"
End If
j = j + 1
End If
Next i

End Sub
'**************************************************************************************************************************
'******************************************������ ���� DBF ��� ������******************************************************
'**************************************************************************************************************************
Sub import_ALL()
Dim i As Integer
Dim j As Integer
Dim rst As Integer
Dim strsql As String
i = 1
j = 15
For i = 1 To Application.CurrentDb.QueryDefs.Count - 3
If Dir(Application.CurrentProject.Path & "\DBF\�����" & j & ".dbf") <> "" Then
If Application.CurrentDb.TableDefs("�����" & j).DateCreated <> "" Then
strsql = "DROP TABLE �����" & j
Application.CurrentDb.Execute (strsql)
End If
DoCmd.TransferDatabase acImport, "dBase IV", Application.CurrentProject.Path & "\DBF", acTable, "�����" & j & ".dbf", "�����" & j
End If
j = j + 1
Next i
End Sub
'******************************************************************************************************
'*******************************�������� ���� ������� ����*********************************************
'******************************************************************************************************
Sub export_ALL_rtf()
Dim i As Integer
Dim j As Integer
Dim rst As Integer
Dim strsql As String
i = 1
For j = 15 To 34
For i = 1 To Application.CurrentDb.TableDefs.Count - 1
If Application.CurrentDb.TableDefs(i).Name = "�����" & j Then
DoCmd.OutputTo acOutputReport, "�����" & " " & j, acFormatRTF, Application.CurrentProject.Path & "\�����" & " " & j & ".rtf", False
End If
Next i
Next j
End Sub
