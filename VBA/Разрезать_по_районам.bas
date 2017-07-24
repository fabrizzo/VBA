Attribute VB_Name = "Разрезать_по_районам"
Sub Нарезка_по_районам1()
    Dim kod, a As String: Dim Fname(70, 2) As String
    Dim kol As Integer: Dim b As Integer: Dim j As Integer
    Dim i As Integer
    Dim region() As String
    Dim maxnum As Integer
    Dim num As Variant
    Dim bf As Document
    For Each bf In Documents
    bf.Close SaveChanges:=False
    Next bf
    num = InputBox("Сколько подразделений будем резать ? (Длинная)")
    Select Case StrPtr(num)
    Case 0
        Exit Sub
    Case Else
        maxnum = CInt(num)
    End Select
    ReDim region(1 To maxnum)
    For i = 1 To maxnum
    region(i) = InputBox("Введите подразделение: " & i)
    Next i
    For j = 1 To maxnum
    kod = region(j)
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    Windows(1).Activate
    kol = 0
    '12 // коррупция
    kol = kol + 1: Fname(kol, 1) = "12": Fname(kol, 2) = "фабула"
    '13 // экономика
    kol = kol + 1: Fname(kol, 1) = "13": Fname(kol, 2) = "фабула"
    '14_1_14 // многоэпизодные дела
    kol = kol + 1: Fname(kol, 1) = "14_1_14": Fname(kol, 2) = "фабула"
    '15 // ЖКХ
    kol = kol + 1: Fname(kol, 1) = "15": Fname(kol, 2) = "фабула"
    '16 // ТЭК
    kol = kol + 1: Fname(kol, 1) = "16": Fname(kol, 2) = "фабула"
    '17 // НАЦ.ПРОЕКТ
    kol = kol + 1: Fname(kol, 1) = "17": Fname(kol, 2) = "фабула"
    '18 //  ОСВОЕНИЕ БЮДЖЕТНЫХ СРЕДСТВ
    kol = kol + 1: Fname(kol, 1) = "18": Fname(kol, 2) = "фабула"
    '19 //  В СФЕРЕ ЛЕСОИСПОЛЬЗОВАНИЯ
    kol = kol + 1: Fname(kol, 1) = "19": Fname(kol, 2) = "фабула"
    '20 //  С ИСПОЛЬЗОВАНИЕМ, В ТОМ ЧИСЛЕ С ПРИМЕНЕНИЕМ ОРУЖИЯ
    kol = kol + 1: Fname(kol, 1) = "20": Fname(kol, 2) = "фабула"
    '21 // совершено в общественных местах (без улиц)
    kol = kol + 1: Fname(kol, 1) = "21": Fname(kol, 2) = "фабула"
    '22 // совершено НЕ в общественных местах и улицах
    kol = kol + 1: Fname(kol, 1) = "22": Fname(kol, 2) = "фабула"
    '23 // совершено на улицах
    kol = kol + 1: Fname(kol, 1) = "23": Fname(kol, 2) = "фабула"
    '25 //  СОВЕРШЕНО В ГРУППЕ
    kol = kol + 1: Fname(kol, 1) = "25": Fname(kol, 2) = ""
    '26 // СОВЕРШЕНО В ГРУППЕ ПО СГОВОРУ
    kol = kol + 1: Fname(kol, 1) = "26": Fname(kol, 2) = ""
    '27 // СОВЕРШЕННЫХ В СОСТАВЕ ОПГ И ОПС
    kol = kol + 1: Fname(kol, 1) = "27": Fname(kol, 2) = ""
    '28 // лица , совершившие преступления, по которым отказано в возбуждении по п.3 ч.1 ст.24 и п.4 ч.1 ст.24 УПК РФ
    kol = kol + 1: Fname(kol, 1) = "28": Fname(kol, 2) = ""
    '29 // СОВЕРШЕННЫЕ НЕСОВЕРШЕННОЛЕТНИМИ
    kol = kol + 1: Fname(kol, 1) = "29": Fname(kol, 2) = "фабула"
    '29 // СОВЕРШЕННЫЕ НЕСОВЕРШЕННОЛЕТНИМИ
    kol = kol + 1: Fname(kol, 1) = "29_1": Fname(kol, 2) = ""
    ' 30 // в алкогольном опьянении
    kol = kol + 1: Fname(kol, 1) = "30": Fname(kol, 2) = ""
    ' 30_1 // в наркотическом опьянении
    kol = kol + 1: Fname(kol, 1) = "30_1": Fname(kol, 2) = ""
    ' 30_2 // в токсическом опьянении
    kol = kol + 1: Fname(kol, 1) = "30_2": Fname(kol, 2) = ""
    ' 30_3 // алкоголик
    kol = kol + 1: Fname(kol, 1) = "30_3": Fname(kol, 2) = ""
    ' 30_4 // таксикоманам
    kol = kol + 1: Fname(kol, 1) = "30_4": Fname(kol, 2) = ""
    ' 30_5 // потребитель наркотическиз средств и психотропных веществ
    kol = kol + 1: Fname(kol, 1) = "30_5": Fname(kol, 2) = ""
    ' 31 // хронический алкоголик
    kol = kol + 1: Fname(kol, 1) = "31": Fname(kol, 2) = ""
    ' 31_1 // хронический токсикоманом
    kol = kol + 1: Fname(kol, 1) = "31_1": Fname(kol, 2) = ""
    ' 31_2 // не потребитель наркотическиз средств и психотропных веществ
    kol = kol + 1: Fname(kol, 1) = "31_2": Fname(kol, 2) = ""
    ' 31_3 // в алкогольном опьянении
    kol = kol + 1: Fname(kol, 1) = "31_3": Fname(kol, 2) = ""
    ' 31_4 // в алкогольном опьянении
    kol = kol + 1: Fname(kol, 1) = "31_4": Fname(kol, 2) = ""
    ' 32 // НЕ в алкогольном опьянении
    kol = kol + 1: Fname(kol, 1) = "32": Fname(kol, 2) = ""
    ' 33_2 // НЕ потребителем наркотических средств  и психотропных веществ
    kol = kol + 1: Fname(kol, 1) = "33_2": Fname(kol, 2) = ""
    '34 // ранее совершавшие преступления
    kol = kol + 1: Fname(kol, 1) = "34": Fname(kol, 2) = ""
    '34_1 // ранее совершавшие преступления
    kol = kol + 1: Fname(kol, 1) = "34_1":  Fname(kol, 2) = ""
    '35 // РАНЕЕ СУДИМЫЕ ЛИЦА
    kol = kol + 1: Fname(kol, 1) = "35":  Fname(kol, 2) = ""
    '35_1 // РАНЕЕ СУДИМЫЕ ЛИЦА
    kol = kol + 1: Fname(kol, 1) = "35_1": Fname(kol, 2) = ""
    '36 //  БЕЗРАБОТНЫЕ
    kol = kol + 1: Fname(kol, 1) = "36": Fname(kol, 2) = ""
    '37 //  БЕЗРАБОТНЫЕ
    kol = kol + 1: Fname(kol, 1) = "37":  Fname(kol, 2) = ""
    '38 //  БЕЗ ПОСТОЯННОГО ИСТОЧНИКА ДОХОДА
    kol = kol + 1: Fname(kol, 1) = "38":  Fname(kol, 2) = ""
    '39 // совершено иностранными гражданами
    kol = kol + 1: Fname(kol, 1) = "39":  Fname(kol, 2) = ""
    '39_1 // совершено лицами без гражданства
    kol = kol + 1: Fname(kol, 1) = "39_1": Fname(kol, 2) = ""
    '39_2 // совершено мигрантами
    kol = kol + 1: Fname(kol, 1) = "39_2": Fname(kol, 2) = ""
    '40 // совершено в отношении иностранных граждан
    kol = kol + 1: Fname(kol, 1) = "40":   Fname(kol, 2) = ""
    '40_1 // совершено в отношении мигрантов
    kol = kol + 1: Fname(kol, 1) = "40_1": Fname(kol, 2) = ""
    '43 //   общеуголовной направленности  причиненный, возмещенный ущерб
    kol = kol + 1: Fname(kol, 1) = "43":  Fname(kol, 2) = ""
    '43_1 // преступления коррупционной направленности, по которым выставлена карточка ф-4 с указанием суммы ущерба
    kol = kol + 1: Fname(kol, 1) = "43_1": Fname(kol, 2) = ""
    '43_2 //   экономика  причиненный, возмещенный ущерб
    kol = kol + 1: Fname(kol, 1) = "43_2": Fname(kol, 2) = ""
    '43_3 //  прекращенные причиненный, возмещенный ущерб
    kol = kol + 1: Fname(kol, 1) = "43_3": Fname(kol, 2) = ""
    '43_4 //  об отказе  причиненный, возмещенный ущерб
    kol = kol + 1: Fname(kol, 1) = "43_4": Fname(kol, 2) = ""
    '43_5 //  лесопользование  причиненный, возмещенный ущерб
    kol = kol + 1: Fname(kol, 1) = "43_5": Fname(kol, 2) = ""
    For b = 1 To kol Step 1
        If Fname(b, 2) = "фабула" Then
            ' C ФАБУЛАМИ
            Application.ScreenUpdating = False
            Application.DisplayScrollBars = False
            
            ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_с_фабулами(kod)
            ActiveDocument.Close
            
            ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_с_фабулами(kod)
            ActiveDocument.Close
            
            ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_с_фабулами(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_с_фабулами(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_с_фабулами(kod)
            ActiveDocument.Close
        Else ' БЕЗ ФАБУЛ
    
            ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_без_фабул(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_без_фабул(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_без_фабул(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_без_фабул(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = Вырезка_без_фабул(kod)
            ActiveDocument.Close
        End If
    Next b
    Call Обработка_готового_файла
    Call MyFileSave(kod)
    ActiveDocument.Close SaveChanges:=False
Next j
 Application.ScreenUpdating = True
 Application.DisplayScrollBars = True
End Sub
Sub Обработка_готового_файла()
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(3)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(21)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    Selection.WholeStory
    Selection.Font.Size = 6
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "fssp"
        .Replacement.Text = "ФССП"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    
  Selection.HomeKey wdStory
End Sub
Sub MyFileSave(ByVal id As String)
Dim sPath As String
sPath = "C:\"
With Dialogs(wdDialogFileSaveAs)
    .name = sPath & "ОВД - " & id
    .Show
    End With
End Sub
Function Вырезка_с_фабулами(q) As String
ChangeWinErr:
    Dim p As String
    p = ":" & q & "*-----------------------"
    Dim r, f, flag As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
    flag = False
    Do
        With r.Find
            .ClearFormatting
            .Text = p
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            If .Execute Then
                If flag = False Then
                    flag = True
                Selection.HomeKey wdStory                   ' перемещаем курсор в начало документа
                With Selection.Find
                    .Text = "//*=Q"    ' искомый текст
                    .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
                    .Wrap = wdFindStop                  ' остановиться на найденном
                    Do
                        .Execute
                    Loop Until Not .Found               ' искать пока поиск (.Find) не завершил поиск
                End With 'теперь на 1-м фрагменте закладка a, на 2-м (если есть) - b; и далее (см. по Ctrl-Shift-F5)
                On Error GoTo ChangeWinErr
                Selection.Copy
                Windows(1).Activate
                Windows(2).Activate
                Selection.PasteAndFormat (wdPasteDefault)
                Selection.TypeParagraph
                Windows(1).Activate
                End If
                If f Then
                    If r.Start = firstOccurence Then ' MsgBox "Файл сформирован", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select 'MsgBox "ДАЛЕЕ...", vbInformation
    Selection.Copy
    Windows(1).Activate
    Windows(2).Activate
    Selection.PasteAndFormat (wdPasteDefault)
    Windows(1).Activate
                Set r = ActiveDocument.Range(r.End, r.End)
            Else 'MsgBox "Текст не найден", vbExclamation
                Exit Do
            End If
        End With
    Loop
End Function
Function Вырезка_без_фабул(q) As String
ChangeWinErr:
    Dim p As String
    p = ":" & q & " "
    Dim r, f, flag As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
    flag = False
    Do
        With r.Find
            .ClearFormatting
            .Text = p
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            If .Execute Then
                    If flag = False Then
                    flag = True
                    Selection.HomeKey wdStory                   ' перемещаем курсор в начало документа
                    With Selection.Find
                        .Text = "//*==Q"    ' искомый текст
                        .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
                        .Wrap = wdFindStop                  ' остановиться на найденном
                        Do
                            .Execute
                     
                        Loop Until Not .Found               ' искать пока поиск (.Find) не завершил поиск
                    End With 'теперь на 1-м фрагменте закладка a, на 2-м (если есть) - b; и далее (см. по Ctrl-Shift-F5)
                    On Error GoTo ChangeWinErr
                    Selection.Copy
                    Windows(1).Activate
                    Windows(2).Activate
                    Selection.PasteAndFormat (wdPasteDefault)
                    Selection.TypeParagraph
                    Windows(1).Activate
                    End If
                If f Then
                    If r.Start = firstOccurence Then 'MsgBox "Файл сформирован", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select
                Selection.MoveEndUntil Cset:=Chr(13)
    Selection.Copy
    Windows(1).Activate
    Windows(2).Activate
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.TypeParagraph
    Windows(1).Activate
                Set r = ActiveDocument.Range(r.End, r.End)
            Else 'MsgBox "Текст не найден", vbExclamation
                Exit Do
            End If
        End With
    Loop
End Function
