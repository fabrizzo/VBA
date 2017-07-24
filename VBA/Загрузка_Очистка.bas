Attribute VB_Name = "Загрузка_Очистка"
Option Explicit
'***********************************************************************
'*****************Сбор имен таблиц**************************************
'***********************************************************************
Sub FieldNames()
    Application.ScreenUpdating = False
    Dim LastCol, LastRow, i, DelRows As Integer
    Dim Cell As Object: Dim rRange As Range
    Dim iWS As Worksheet
    '******Затираем старые записи**********************
        Worksheets("my").Activate
        DelRows = Range("A" & Rows.Count).End(xlUp).row
        If DelRows = 1 Then
        DelRows = DelRows + 1
        End If
        Range("A2:C" & DelRows).ClearContents
     '******Создаем новые**********************
    For Each iWS In ThisWorkbook.Worksheets
    If iWS.Name = "my" Then
        Worksheets("общее количество исков").Activate
        MsgBox "Названия столбцов обновлены", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    Else
    iWS.Activate
    End If
    LastCol = 0
    LastRow = 0
    LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    Set rRange = Range("C1", Cells(1, LastCol))
    For Each Cell In rRange
        If Cell <> "" Then
            With Worksheets("my")
                .Activate
                If .Range("B2").Value = "" Then
                    LastRow = 2
                Else
                    LastRow = .Range("B65536").End(xlUp).row
                    LastRow = LastRow + 1
                End If
                .Range("C" & LastRow).Value = Cell.Value
                .Range("B" & LastRow).Value = Cell.Value
                .Range("A" & LastRow).Value = iWS.Name
            End With
        End If
    Next Cell
    Next iWS
End Sub
'************************************************************************
'*****************Загрузка карточек в таблицу*****************************
'************************************************************************
Sub InsertNew()
    Call Start
    Dim fto
    Dim wbThis As Workbook, wb As Workbook, wsMy As Worksheet, ws As Worksheet
    Dim i&, j%, sWbName$, sShtName$, sColName$, sRegName$
    Dim sText$, c As Range, cMy As Range, cRegion As Range, lr&
    Set wbThis = ThisWorkbook
    Set wsMy = wbThis.Worksheets("my")
    Application.ScreenUpdating = False
    ChDir ActiveWorkbook.Path
    fto = Application.GetOpenFilename _
    (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
    MultiSelect:=True, Title:="выберите файлы с городами")
    If TypeName(fto) = "Boolean" Then
        MsgBox "Не выбрано ни одного файла!"
        Call Final
        Exit Sub
    End If
'************************************************************************
    For i = 1 To UBound(fto)
        Workbooks.Open Filename:=fto(i)
        Set wb = ActiveWorkbook: sRegName = Left(Range("E3").Value, 4)
        Set ws = wb.Worksheets(1)
'************************************************************************
        For j = 2 To wsMy.[myTable].Rows.Count
            If wsMy.[myTable].Cells(j, 1) = "" Then Exit For
                sShtName = wsMy.[myTable].Cells(j, 1).Value
                sColName = wsMy.[myTable].Cells(j, 2).Value
                sText = wsMy.[myTable].Cells(j, 3).Value
                Set c = ws.UsedRange.Find(what:=sText, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
            If c Is Nothing Then Exit For
'************************************************************************
            With wbThis.Worksheets(sShtName)
                Set cMy = .Rows(1).Find(what:=sColName)
                If cMy Is Nothing Then Exit For
                    lr = .Cells(.Rows.Count, 1).End(xlUp).row
                        If lr < 5 Then
                            lr = 5
                        Else
                            Set cRegion = .Range(.[a4], .Cells(lr, 1)).Find(what:=sRegName)
                            If cRegion Is Nothing Then lr = lr + 1 Else lr = cRegion.row
                            End If
                            .Cells(lr, cMy.Column).Resize(1, 9).Value = ws.Cells(c.row, 7).Resize(1, 9).Value
            End With
'************************************************************************
            Next
            wb.Close False
        Next
            Call Final
End Sub
'************************************************************************
'*****************Очистка таблицы***************************************
'************************************************************************
Sub Clean1Row()
Dim numrow As String: Dim rRange As Range: Dim LastRow As Long
Dim Cell As Object
Application.ScreenUpdating = False
numrow = ""
numrow = InputBox("Введите код подразделения который хотите очистить, чтобы удалить все данные введите знак :  * ", "Удаление строки из всех отчетов") 'спрашиваем какую строку или все данные
If numrow = vbNullString Then
Exit Sub
End If
If IsNumeric(numrow) Then
    numrow = CInt(numrow) 'смотрим номер строки
       Worksheets("общее количество исков").Activate
    Set rRange = Range("A4:A80").Find(what:=numrow)
            Range("C" & rRange.row & ":K" & rRange.row).ClearContents
            Worksheets("гражданское производство").Activate
            Range("C" & rRange.row & ":K" & rRange.row).ClearContents
            Worksheets("в интересах граждан").Activate
            Range("C" & rRange.row & ":IK" & rRange.row).ClearContents
            Worksheets("в защиту несовершеннолетних").Activate
            Range("C" & rRange.row & ":CE" & rRange.row).ClearContents
            Worksheets("В интересах РФ").Activate
            Range("C" & rRange.row & ":CE" & rRange.row).ClearContents
            Worksheets("КАС РФ").Activate
            Range("C" & rRange.row & ":T" & rRange.row).ClearContents
            Worksheets("в порядке УПК РФ").Activate
            Range("C" & rRange.row & ":BD" & rRange.row).ClearContents
            Worksheets("общее количество исков").Activate
            Range("C4").Activate
'************************************************************************
ElseIf numrow = "*" Then
        Worksheets("общее количество исков").Activate
        Range("C4:K11,C13:K33,C35:K43,C45:K49,C51:K64,C66:K72,C75:K80").ClearContents
        Worksheets("гражданское производство").Activate
        Range("C4:K11,C13:K33,C35:K43,C45:K49,C51:K64,C66:K72,C75:K80").ClearContents
        Worksheets("в интересах граждан").Activate
        Range("C4:IK11,C13:IK33,C35:IK43,C45:IK49,C51:IK64,C66:IK72,C75:IK80").ClearContents
        Worksheets("в защиту несовершеннолетних").Activate
        Range("C4:CE11,C13:CE33,C35:CE43,C45:CE49,C51:CE64,C66:CE72,C75:CE80").ClearContents
        Worksheets("В интересах РФ").Activate
        Range("C4:CE11,C13:CE33,C35:CE43,C45:CE49,C51:CE64,C66:CE72,C75:CE80").ClearContents
        Worksheets("КАС РФ").Activate
        Range("C4:T11,C13:T33,C35:T43,C45:T49,C51:T64,C66:T72,C75:T80").ClearContents
        Worksheets("в порядке УПК РФ").Activate
        Range("C4:BD11,C13:BD33,C35:BD43,C45:BD49,C51:BD64,C66:BD72,C75:BD80").ClearContents
        Worksheets("общее количество исков").Activate
        Range("C4").Activate
Else
    Exit Sub
End If
End Sub
'************************************************************************
'*****************Выгрузка памяти***************************************
'************************************************************************
Sub Start()
    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual: Application.EnableEvents = False
    Application.DisplayStatusBar = False: Application.DisplayAlerts = False: ActiveSheet.DisplayPageBreaks = False
End Sub
Sub Final()
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True
    Application.DisplayStatusBar = True: Application.DisplayAlerts = True: ActiveSheet.DisplayPageBreaks = True
End Sub

