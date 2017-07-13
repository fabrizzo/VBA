Attribute VB_Name = "Import_Del"
Option Compare Database
Sub Clean_ALL_Tables()
    Dim i As Integer
    Dim strSQL As String
    For i = 0 To Application.CurrentDb.TableDefs.Count - 1
        If Left(Application.CurrentDb.TableDefs(i).Name, 4) = "Form" Then
            If Application.CurrentDb.TableDefs(i).RecordCount > 0 Then
                strSQL = "Delete * FROM " & Application.CurrentDb.TableDefs(i).Name
                CurrentDb.Execute (strSQL)
            End If
        End If
    If Left(Application.CurrentDb.TableDefs(i).Name, 6) = "Fabula" Then
        If Application.CurrentDb.TableDefs(i).RecordCount > 0 Then
            strSQL = ""
            strSQL = "Delete * FROM Fabula "
            CurrentDb.Execute (strSQL)
        End If
    End If
    Next i
    MsgBox "Готово", vbInformation
End Sub
Sub Import_ALL_Tables()
    Dim i As Integer
    For i = 0 To Application.CurrentDb.TableDefs.Count - 1
        If Left(Application.CurrentDb.TableDefs(i).Name, 4) = "Form" Then
            If Application.CurrentDb.TableDefs(i).RecordCount = 0 Then
                If Dir(Application.CurrentProject.Path & "\" & Application.CurrentDb.TableDefs(i).Name & ".xlsx") <> "" Then
                    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, Application.CurrentDb.TableDefs(i).Name, Application.CurrentProject.Path & "\" & Application.CurrentDb.TableDefs(i).Name & ".xlsx", True
                End If
            End If
        End If
    If Left(Application.CurrentDb.TableDefs(i).Name, 6) = "Fabula" Then
        If Application.CurrentDb.TableDefs(i).RecordCount = 0 Then
            If Dir(Application.CurrentProject.Path & "\" & Application.CurrentDb.TableDefs(i).Name & ".xlsx") <> "" Then
                DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Fabula", Application.CurrentProject.Path & "\Fabula" & ".xlsx", True
            End If
        End If
    End If
    Next i
    MsgBox "Готово", vbInformation
    End Sub
