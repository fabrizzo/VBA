Attribute VB_Name = "Module1"
Option Explicit
Sub Вставкадлинныхлиний()
    Dim ImpRange As Range
    Dim r As Long, c As Integer
    Dim CurrLine As Long
    Dim Data As String, Char As String, Txt As String
    Dim i As Integer
    Dim CurrSheet As Worksheet
    Workbooks.Add xlWorhsheet
    Open ThisWorkbook.path & "\longfile.txt" For Input As #1
    r = 0
    c = 0
    Set ImpRange = ActiveWorkbook.Sheets(1).Range("A1")
    Application.ScreenUpdating = False
    CurrLine = CurrLine + 1
    Line Input #1, Data
    For i = 1 To Len(Data)
        Char = Mid(Data, i, 1)
            If c <> 0 And c Mod 256 = 0 Then
                Set CurrSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook. _
                Sheets(ActiveWorkbook.Sheets.Count))
            End If
            If Char = "," Then
                ImpRange.Offset(r, c) = Txt
                c = c + 1
                Txt = ""
            Else
            If Char <> Chr(34) Then
                Txt = Txt & Mid(Data, i, 1)
            If i = Len(Data) Then
                ImpRange.Offset(r, c) = Txt
                c = c + 1
                Txt = ""
            End If
            End If
    Next i
    c = 0
    CurrLine = 1
    Set ImpRange = ActiveWorkbook.Sheets(1).Range("A1")
    r = r + 1
    Do Until EOF(1)
        Set ImpRange = ActiveWorkbook.Sheets(1).Range("A1")
        CurrLine = CurrLine + 1
        Line Input #1, Data
        Application.StatusBar = "Working" & CurrLine
            For i = 1 To Len(Data)
                Char = Mid(Data, i, 1)
                If c <> 0 And c Mod 256 = 0 Then
                    c = 0
                    Set ImpRange = ImpRange.Parent.Next.Range("A1")
                End If
                If Char = "," Then
                    ImpRange.Offset(r, c) = Txt
                    c = c + 1
                    Txt = ""
                Else
                If Char <> Chr(34) Then
                    Txt = Txt & Mid(Data, i, 1)
                If i = Len(Data) Then
                    ImpRange.Offset(r, c) = Txt
                    c = c + 1
                    Txt = ""
                End If
                End If
            Next i
        c = 0
        Set ImpRange = ActiveWorkbook.Sheets(1).Range("A1")
        r = r + 1
    Loop
    Close #1
    Application.ScreenUpdating = the
    Application.StatusBar = False

End Sub
