Attribute VB_Name = "Import/Export_Query"
Option Compare Database
Option Explicit
Sub export_ALL_Query_to_DBF()
'Экспортируем все запросы формата DBF в папку DBF
Dim i As Integer
Dim j As Integer
Dim rst As Integer
i = 1
j = 15
For i = 1 To Application.CurrentDb.QueryDefs.Count - 3
If Dir(Application.CurrentProject.Path & "\DBF\Глава" & " " & j & ".dbf") = "" Then
rst = DCount("*", "Глава" & " " & j)
If rst <> 0 Then
DoCmd.TransferDatabase acExport, "dBase IV", Application.CurrentProject.Path & "\DBF", acQuery, "Глава" & " " & j, "Глава" & j & ".dbf"
End If
j = j + 1
End If
Next i

End Sub
'**************************************************************************************************************************
'******************************************ИМПОРТ ВСЕХ DBF КАК ТАБЛИЦ******************************************************
'**************************************************************************************************************************
Sub import_ALL()
Dim i As Integer
Dim j As Integer
Dim rst As Integer
Dim strsql As String
i = 1
j = 15
For i = 1 To Application.CurrentDb.QueryDefs.Count - 3
If Dir(Application.CurrentProject.Path & "\DBF\Глава" & j & ".dbf") <> "" Then
If Application.CurrentDb.TableDefs("Глава" & j).DateCreated <> "" Then
strsql = "DROP TABLE Глава" & j
Application.CurrentDb.Execute (strsql)
End If
DoCmd.TransferDatabase acImport, "dBase IV", Application.CurrentProject.Path & "\DBF", acTable, "Глава" & j & ".dbf", "Глава" & j
End If
j = j + 1
Next i
End Sub
'******************************************************************************************************
'*******************************Выгрузка всех отчетов глав*********************************************
'******************************************************************************************************
Sub export_ALL_rtf()
Dim i As Integer
Dim j As Integer
Dim rst As Integer
Dim strsql As String
i = 1
For j = 15 To 34
For i = 1 To Application.CurrentDb.TableDefs.Count - 1
If Application.CurrentDb.TableDefs(i).Name = "Глава" & j Then
DoCmd.OutputTo acOutputReport, "Глава" & " " & j, acFormatRTF, Application.CurrentProject.Path & "\Глава" & " " & j & ".rtf", False
End If
Next i
Next j
End Sub
