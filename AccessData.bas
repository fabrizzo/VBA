Attribute VB_Name = "AccessData"
Option Explicit
Sub AccessDublicateData()
    Dim Acc As Object
    Dim qdf As Object
    Const acImport = 0
    Const acExport = 1
    Const acSpreadsheetTypeExcel12 = 9
    Const acSpreadsheetTypeExcel10 = 10
    Dim WorkName As String
   WorkName = ActiveWorkbook.Path & "\" & "FabulaDataBase.accdb"
    Set Acc = CreateObject("Access.Application")
    Acc.Visible = True
    If Dir(WorkName) <> "" Then
   Acc.OpenCurrentDataBase (WorkName)
    End If
    If Dir(Acc.CurrentProject.Path & "\Fabula.xlsx") <> "" Then
    Acc.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Fabula", Acc.CurrentProject.Path & "\Fabula.xlsx", False
    Dim strSQL As String
    strSQL = "SELECT Last(Fabula.F1) AS [ÎÂÄ], First(Fabula.[F2]) AS [ÍÎÌÅÐ ÏÐÅÑÒ], First(Fabula.[F3]) AS [ÎÑÍ], Last(Fabula.F4) AS [ÄÀÒÀ ÂÎÇÁÓÆÄÅÍÈß], Last(Fabula.F5) AS [ÏÎÑÒ ÈÖ], Last(Fabula.F6) AS [ÑÒ], Last(Fabula.F7) AS [Ç], Last(Fabula.F8) AS [×], Last(Fabula.F9) AS [Ï1], Last(Fabula.F10) AS [Ï2], Last(Fabula.F11) AS [Ï3], Last(Fabula.F12) AS [ÄÀÒÀ_ÐÅØÅÍÈß], Last(Fabula.F13) AS [ÔÀÁÓËÀ], Last(Fabula.F14) AS [ÔÀÉË] FROM Fabula GROUP BY Fabula.[F2], Fabula.[F3] HAVING (((Count(Fabula.[F2]))>=1) AND ((Count(Fabula.[F3]))>=1))"
    Set qdf = Acc.CurrentDb.CreateQueryDef("tempQry", strSQL)
    Acc.DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel10, "tempQry", Acc.CurrentProject.Path & "\Unique Fabula.xlsx"
    End If
    Acc.Quit
    Set Acc = Nothing
End Sub
