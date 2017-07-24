Attribute VB_Name = "NewMacros3"

Function Вырезка_с_фабулами(q) As String

'
 

'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
 '   Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."


' макрос2 Макрос
'ловим шапку

 Selection.HomeKey wdStory                   ' перемещаем курсор в начало документа
    With Selection.Find
        .Text = "//*=Q"    ' искомый текст
        .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
        .Wrap = wdFindStop                  ' остановиться на найденном
        Do
            .Execute
     
        Loop Until Not .Found               ' искать пока поиск (.Find) не завершил поиск
    End With
    'теперь на 1-м фрагменте закладка a, на 2-м (если есть) - b; и далее (см. по Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate


' MsgBox "ДАЛЕЕ.1.."




'  поиск нужного фрагмента и выделение
Dim p As String

p = ":" & q & "*-----------------------"

 ' MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                       ' MsgBox "Файл сформирован", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select
                 'MsgBox "ДАЛЕЕ...", vbInformation
                
                ' нашли
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)


'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "Текст не найден", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function

Function Вырезка_без_фабул(q) As String
'
'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
'    Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."


' макрос2 Макрос
'ловим шапку

 Selection.HomeKey wdStory                   ' перемещаем курсор в начало документа
    With Selection.Find
        .Text = "//*==Q"    ' искомый текст
        .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
        .Wrap = wdFindStop                  ' остановиться на найденном
        Do
            .Execute
     
        Loop Until Not .Found               ' искать пока поиск (.Find) не завершил поиск
    End With
    'теперь на 1-м фрагменте закладка a, на 2-м (если есть) - b; и далее (см. по Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."



'  поиск нужного фрагмента и выделение
Dim p As String

p = ":" & q & " "

' MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                        'MsgBox "Файл сформирован", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select
                Selection.MoveEndUntil Cset:=Chr(13)

               ' MsgBox "ДАЛЕЕ...", vbInformation
                
                ' нашли
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)
Selection.TypeParagraph

'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "Текст не найден", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function


Function ФСКН_Вырезка_с_фабулами(q) As String

'
'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
 '   Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."


' макрос2 Макрос
'ловим шапку

 Selection.HomeKey wdStory                   ' перемещаем курсор в начало документа
    With Selection.Find
        .Text = "//*==Q"    ' искомый текст
        .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
        .Wrap = wdFindStop                  ' остановиться на найденном
        Do
            .Execute
     
        Loop Until Not .Found               ' искать пока поиск (.Find) не завершил поиск
    End With
    'теперь на 1-м фрагменте закладка a, на 2-м (если есть) - b; и далее (см. по Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."



'  поиск нужного фрагмента и выделение
Dim p As String

p = " " & q & " 2*-----------------------"

 'MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                        'MsgBox "Файл сформирован", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 22, r.End + 1).Select
         
                ' MsgBox "ДАЛЕЕ...", vbInformation
                
                ' нашли
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)


'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "Текст не найден", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function


Function ФСКН_Вырезка_без_фабул(q) As String
'
'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
'    Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."


' макрос2 Макрос
'ловим шапку

 Selection.HomeKey wdStory                   ' перемещаем курсор в начало документа
    With Selection.Find
        .Text = "//*==Q"    ' искомый текст
        .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
        .Wrap = wdFindStop                  ' остановиться на найденном
        Do
            .Execute
     
        Loop Until Not .Found               ' искать пока поиск (.Find) не завершил поиск
    End With
    'теперь на 1-м фрагменте закладка a, на 2-м (если есть) - b; и далее (см. по Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate
' MsgBox "ДАЛЕЕ.1.."



'  поиск нужного фрагмента и выделение
Dim p As String

p = " " & q & " 2"


' MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                        'MsgBox "Файл сформирован", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 22, r.End + 1).Select
                Selection.MoveEndUntil Cset:=Chr(13)

               ' MsgBox "ДАЛЕЕ...", vbInformation
                
                ' нашли
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)
Selection.TypeParagraph

'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "Текст не найден", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function
Sub Нарезка_по_районам()
Attribute Нарезка_по_районам.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос2"
    
 Dim kod, a As String
' создаем новый документ
Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
Windows(1).Activate
      
kod = InputBox("введите код ОВД")
        
 ' 09 // выявленные и поставленные на учет по инициативе прокурора // (форма 1 реквизит 9 код 125)
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

' 10 // при отмене постановлений об отказе в возбуждении уг.дела
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

'10_1 // при отмене постановлений об отказе в возбуждении уг.дела по материалам следствия // реквизит 103 код 31
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'10_2 // при отмене постановлений об отказе в возбуждении уг.дела по материалам дознания // реквизит 103 код 32
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'10_3 // при отмене постановлений об отказе в возбуждении уг.дела по материалам дознания // реквизит 103 код 33
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

'11 // о выделенных преступлениях без постановки на учет
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

'12 // коррупция
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'13 // экономика
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'14_1_14 // многоэпизодные дела
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'15 // ЖКХ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


'16 // ТЭК
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'17 // НАЦ.ПРОЕКТ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'18 //  ОСВОЕНИЕ БЮДЖЕТНЫХ СРЕДСТВ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'19 //  В СФЕРЕ ЛЕСОИСПОЛЬЗОВАНИЯ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'20 //  С ИСПОЛЬЗОВАНИЕМ, В ТОМ ЧИСЛЕ С ПРИМЕНЕНИЕМ ОРУЖИЯ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
'21 // совершено в общественных местах (без улиц)
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

'22 // совершено НЕ в общественных местах и улицах
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

'23 // совершено на улицах
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

'24 // совершено НЕ на улицах
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

'25 //  СОВЕРШЕНО В ГРУППЕ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close


'26 // СОВЕРШЕНО В ГРУППЕ ПО СГОВОРУ

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

'27 // СОВЕРШЕННЫХ В СОСТАВЕ ОПГ И ОПС
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

'28 // лица , совершившие преступления, по которым отказано в возбуждении по п.3 ч.1 ст.24 и п.4 ч.1 ст.24 УПК РФ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close


'29 // СОВЕРШЕННЫЕ НЕСОВЕРШЕННОЛЕТНИМИ
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_с_фабулами(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close
 
' 30 // в алкогольном опьянении
    ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
 ' 30_1 // в наркотическом опьянении
     ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
 ' 30_2 // в токсическом опьянении
      ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
 ' 31 // хронический алкоголик
      ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
 ' 31_1 // хронический токсикоманом
      ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
  
  ' 31_2 // не потребитель наркотическиз средств и психотропных веществ
       ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
   
 ' 32 // НЕ в алкогольном опьянении
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
 
 ' 33_2 // НЕ потребителем наркотических средств  и психотропных веществ

       ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
   
'34 // ранее совершавшие преступления
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'35 // РАНЕЕ СУДИМЫЕ ЛИЦА
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'36 //  БЕЗРАБОТНЫЕ
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

'38 //  БЕЗ ПОСТОЯННОГО ИСТОЧНИКА ДОХОДА
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
 '39 // совершено иностранными гражданами
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
'39_1 // совершено лицами без гражданства
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
'39_2 // совершено мигрантами
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
'40 // совершено в отношении иностранных граждан
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
 
'40_1 // совершено в отношении мигрантов
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

'43 //   общеуголовной направленности  причиненный, возмещенный ущерб
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'43_1 // преступления коррупционной направленности, по которым выставлена карточка ф-4 с указанием суммы ущерба
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'43_2 //   экономика  причиненный, возмещенный ущерб
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'43_3 //  прекращенные причиненный, возмещенный ущерб
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'43_4 //  об отказе  причиненный, возмещенный ущерб
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close
'43_5 //  лесопользование  причиненный, возмещенный ущерб
         ChangeFileOpenDirectory "C:\ОБРАБОТКА\CK\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\ОБРАБОТКА\SLEDSTV\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\DOZNAN\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ФСКН_Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = Вырезка_без_фабул(kod)
 ActiveDocument.Close



' КОНЭЦ
 Application.Run MacroName:="Обработка_готового_файла"
MsgBox "Файл сформирован", vbInformation

End Sub
Sub Обработка_готового_файла()
Attribute Обработка_готового_файла.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос4"
'
' Макрос4 Макрос
'
'
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



Sub Нарезка_по_районам1()

    
 Dim kod, a As String
 
 Dim Fname(70, 2) As String
 Dim kol As Integer
 Dim b As Integer
 
' создаем новый документ
Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
Windows(1).Activate
      
kod = InputBox("введите код ОВД")


kol = 0



' 09 // выявленные и поставленные на учет по инициативе прокурора // (форма 1 реквизит 9 код 125)
kol = kol + 1
Fname(kol, 1) = "09"
Fname(kol, 2) = "фабула"

' 10 // при отмене постановлений об отказе в возбуждении уг.дела
kol = kol + 1
Fname(kol, 1) = "10"
Fname(kol, 2) = ""

'10_1 // при отмене постановлений об отказе в возбуждении уг.дела по материалам следствия // реквизит 103 код 31
kol = kol + 1
Fname(kol, 1) = "10_1"
Fname(kol, 2) = ""

'10_2 // при отмене постановлений об отказе в возбуждении уг.дела по материалам дознания // реквизит 103 код 32
kol = kol + 1
Fname(kol, 1) = "10_3"
Fname(kol, 2) = ""

'10_3 // при отмене постановлений об отказе в возбуждении уг.дела по материалам дознания // реквизит 103 код 33
kol = kol + 1
Fname(kol, 1) = "10_3"
Fname(kol, 2) = ""

'11 // о выделенных преступлениях без постановки на учет
kol = kol + 1
Fname(kol, 1) = "11"
Fname(kol, 2) = "фабула"

'12 // коррупция
kol = kol + 1
Fname(kol, 1) = "12"
Fname(kol, 2) = "фабула"
 
'13 // экономика
kol = kol + 1
Fname(kol, 1) = "13"
Fname(kol, 2) = "фабула"
 
'14_1_14 // многоэпизодные дела
kol = kol + 1
Fname(kol, 1) = "14_1_14"
Fname(kol, 2) = "фабула"
 
'15 // ЖКХ
kol = kol + 1
Fname(kol, 1) = "15"
Fname(kol, 2) = "фабула"

'16 // ТЭК
kol = kol + 1
Fname(kol, 1) = "16"
Fname(kol, 2) = "фабула"
 
'17 // НАЦ.ПРОЕКТ
kol = kol + 1
Fname(kol, 1) = "17"
Fname(kol, 2) = "фабула"
 
'18 //  ОСВОЕНИЕ БЮДЖЕТНЫХ СРЕДСТВ
kol = kol + 1
Fname(kol, 1) = "18"
Fname(kol, 2) = "фабула"
 
'19 //  В СФЕРЕ ЛЕСОИСПОЛЬЗОВАНИЯ
kol = kol + 1
Fname(kol, 1) = "19"
Fname(kol, 2) = "фабула"
 
'20 //  С ИСПОЛЬЗОВАНИЕМ, В ТОМ ЧИСЛЕ С ПРИМЕНЕНИЕМ ОРУЖИЯ
 kol = kol + 1
Fname(kol, 1) = "20"
Fname(kol, 2) = "фабула"
 
'21 // совершено в общественных местах (без улиц)
kol = kol + 1
Fname(kol, 1) = "21"
Fname(kol, 2) = "фабула"

'22 // совершено НЕ в общественных местах и улицах
kol = kol + 1
Fname(kol, 1) = "22"
Fname(kol, 2) = "фабула"

'23 // совершено на улицах
kol = kol + 1
Fname(kol, 1) = "23"
Fname(kol, 2) = "фабула"

'24 // совершено НЕ на улицах
kol = kol + 1
Fname(kol, 1) = "24"
Fname(kol, 2) = "фабула"

'40 // совершено в отношении иностранных граждан
kol = kol + 1
Fname(kol, 1) = "40"
Fname(kol, 2) = ""
 
'40_1 // совершено в отношении мигрантов
kol = kol + 1
Fname(kol, 1) = "40_1"
Fname(kol, 2) = ""

'43 //   общеуголовной направленности  причиненный, возмещенный ущерб
kol = kol + 1
Fname(kol, 1) = "43"
Fname(kol, 2) = ""

'43_1 // преступления коррупционной направленности, по которым выставлена карточка ф-4 с указанием суммы ущерба
kol = kol + 1
Fname(kol, 1) = "43_1"
Fname(kol, 2) = ""

'43_2 //   экономика  причиненный, возмещенный ущерб
kol = kol + 1
Fname(kol, 1) = "43_2"
Fname(kol, 2) = ""

'43_3 //  прекращенные причиненный, возмещенный ущерб
kol = kol + 1
Fname(kol, 1) = "43_3"
Fname(kol, 2) = ""

'43_4 //  об отказе  причиненный, возмещенный ущерб
kol = kol + 1
Fname(kol, 1) = "43_4"
Fname(kol, 2) = ""

'43_5 //  лесопользование  причиненный, возмещенный ущерб
kol = kol + 1
Fname(kol, 1) = "43_5"
Fname(kol, 2) = ""



For b = 1 To kol Step 1
        
        
If Fname(b, 2) = "фабула" Then

        
        
 ' C ФАБУЛАМИ
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

'ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
'    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
'a = ФСКН_Вырезка_с_фабулами(kod)
' ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSSP\MEC\"
    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\ОБРАБОТКА\GPN\MEC\"
    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
a = Вырезка_с_фабулами(kod)
 ActiveDocument.Close




Else

' БЕЗ ФАБУЛ
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

'ChangeFileOpenDirectory "C:\ОБРАБОТКА\FSKN\MEC\"
'    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
'a = ФСКН_Вырезка_без_фабул(kod)
' ActiveDocument.Close

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





' КОНЭЦ
 Application.Run MacroName:="Обработка_готового_файла"
MsgBox "Файл сформирован", vbInformation


End Sub

