Attribute VB_Name = "ToXlsx"
Option Explicit
Dim BookXL As Object
Sub ConvertToExcel()
    Dim V As String: Dim BrowseFolder As String
    Dim ProName, InitialPath As String
    Set BookXL = CreateObject("Excel.Application")
    BookXL.Visible = True
    BookXL.Workbooks.Add (1)
    BookXL.Workbooks(1).Worksheets(1).Name = "Data"
    ProName = ActiveDocument.Path
    InitialPath = ProName
    '*********************************************************
    On Error Resume Next
    With Application.FileDialog(4)
        .Title = "Выберите папку"
        .InitialFileName = InitialPath
        .Show
    Err.Clear
    V = .SelectedItems(1)
        If Err.Number <> 0 Then
            MsgBox "Не выбрали ничего"
            Exit Sub
        End If
    End With
    BrowseFolder = CStr(V): ListFilesinFolder BrowseFolder, True
    '*********************************************************
    MsgBox "Готово", vbInformation
    'BookXL.Visible = True
    Unload UserForm1
End Sub
Sub ListFilesinFolder(ByVal SourceFolderName As String, ByVal IncludeSubFolders As Boolean)
    Dim FSO As Object: Dim SourceFolder As Object: Dim SubFolder As Object: Dim OldFolder As String: Dim FileItem As Object
    Dim maxlen, Pos As Integer: Dim Fname, strFileName, NewFile As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.getfolder(SourceFolderName)
    NewFile = UserForm1.ComboBox1.Text
    '*********************************************************
        If UserForm1.CheckBox1.Value = True And UserForm1.CheckBox2.Value = False Then
            If Right(SourceFolder, 3) = "GOD" Then
                For Each FileItem In SourceFolder.Files
                    maxlen = Len(FileItem)
                    Pos = InStrRev(FileItem, "\")
                    Fname = Right(FileItem, maxlen - Pos)
                    If Fname = NewFile Then
                    Call Name_Selector(Fname, FileItem)
                    End If
                Next FileItem
            End If
    '*********************************************************
    ElseIf UserForm1.CheckBox2.Value = True And UserForm1.CheckBox1.Value = True Then
          For Each FileItem In SourceFolder.Files
            maxlen = Len(FileItem)
            Pos = InStrRev(FileItem, "\")
            Fname = Right(FileItem, maxlen - Pos)
            If Fname = NewFile Then
                Call Name_Selector(Fname, FileItem)
                End If
          Next FileItem
    '*********************************************************
    ElseIf UserForm1.CheckBox2.Value = True And UserForm1.CheckBox1.Value = False Then
        If Right(SourceFolder, 3) = "MEC" Then
            For Each FileItem In SourceFolder.Files
                maxlen = Len(FileItem)
                Pos = InStrRev(FileItem, "\")
                Fname = Right(FileItem, maxlen - Pos)
                If Fname = NewFile Then
                Call Name_Selector(Fname, FileItem)
               End If
            Next FileItem
        End If
    End If
        
'*********************************************************
        If IncludeSubFolders Then
            For Each SubFolder In SourceFolder.SubFolders
                ListFilesinFolder SubFolder.Path, True
            Next SubFolder
        End If
'*********************************************************
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub
'***********************************************************************************************************************
'********************************DATA_COPY WITH FABULA***********************************************************
'***********************************************************************************************************************
Function Data_Fab_Copy()
    Dim LastRow, LngLastRow As Long
    Dim LastCol As Long
    Dim Pos As Integer
    Dim strFileName As String
    Dim strFindWord As String
    Dim c As Range
    Const xlToLeft = -4159
    Const xlWhole = 1
    Const xlValues = -4163
    Const xlUp = -4162
    Const xlCenter = -4108
    Const xlLeft = -4131
    Const xlByColumns = 2
    Const xlNext = 1
    strFindWord = "====="
    LastCol = 0
    LngLastRow = 0
    
    Set c = BookXL.Cells.Find(What:=strFindWord, LookIn:=xlValues, LookAt:= _
    xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=True _
    , SearchFormat:=False)
    If c Is Nothing Then Exit Function
    
    LastCol = BookXL.Cells(BookXL.c.Row, BookXL.Columns.Count).End(xlToLeft).Column
    LastCol = LastCol + 1
    
    LngLastRow = BookXL.Range("A" & BookXL.Rows.Count).End(xlUp).Row
    BookXL.Columns(LastCol).Select
    BookXL.Selection.ColumnWidth = 125
    With BookXL.Selection
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = True: .Orientation = 0
        .AddIndent = False: .IndentLevel = 0: .ShrinkToFit = False: .ReadingOrder = xlContext: .MergeCells = False
    End With
    BookXL.Range("A1", BookXL.Cells(LngLastRow, LastCol)).Copy
    BookXL.Workbooks(1).Worksheets("Data").Activate
    BookXL.Columns(LastCol).Select
    BookXL.Selection.ColumnWidth = 125
    If BookXL.Range("A1").Value = "" Then
        BookXL.Range("A1").Select
        BookXL.ActiveSheet.Paste
        BookXL.CutCopyMode = False
        BookXL.Workbooks(2).Close SaveChanges:=False
    Else
        BookXL.Range("A1").Select
        LastRow = 0
        LastRow = BookXL.Range("A" & BookXL.Rows.Count).End(xlUp).Row
        LastRow = LastRow + 1
        BookXL.Range("A" & LastRow).PasteSpecial
        BookXL.CutCopyMode = False
        BookXL.Workbooks(2).Close SaveChanges:=False
    End If
End Function
'***********************************************************************************************************************
'***************************DATA_COPY WITHOUT FABULA***********************************************************
'***********************************************************************************************************************
Function Data_Copy()
    Dim LastRow, Row As Long
    Dim LastCol As Long
    Dim Pos As Integer
    Dim strFileName As String
    Const xlToLeft = -4159
    Const xlUp = -4162
    
    LastCol = 0
    Row = 0
    LastCol = BookXL.Cells(14, BookXL.Columns.Count).End(xlToLeft).Column
    Row = BookXL.Range("A" & BookXL.Rows.Count).End(xlUp).Row
    BookXL.Range("A1", BookXL.Cells(Row, LastCol)).Copy
    BookXL.Workbooks(1).Worksheets("Data").Activate
    If BookXL.Range("A1").Value = "" Then
        BookXL.Range("A1").Select
        BookXL.ActiveSheet.Paste
        BookXL.CutCopyMode = False
        BookXL.Workbooks(2).Close SaveChanges:=False
    Else
        BookXL.Range("A1").Select
        LastRow = 0
        LastRow = BookXL.Range("A" & BookXL.Rows.Count).End(xlUp).Row
        LastRow = LastRow + 1
        BookXL.Range("A" & LastRow).PasteSpecial
        BookXL.CutCopyMode = False
        BookXL.Workbooks(2).Close SaveChanges:=False
    End If
End Function
'***********************************************************************************************************************
'**********************************************CUT_HEADER***********************************************************
'***********************************************************************************************************************
Function Cut_Header(ByVal filename As String) 'Вырезаем лишние символы и шапку
    Documents.Open filename:=filename, Encoding:=1251
    Const wdFindContinue = 1: Dim maxsimbol, Pos, Max As Integer: Dim Fname, Lef, l_Left, DocPath As String
        Fname = ActiveDocument.Name:  maxsimbol = Len(Fname): l_Left = Left(Fname, maxsimbol - 4): Fname = l_Left
        Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "===Q": .Replacement.Text = "": .Forward = True: .Wrap = wdFindContinue: .Format = False
        .MatchCase = False: .MatchWholeWord = False: .MatchWildcards = False: .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute
    If Selection.Start = 0 Then
        ActiveDocument.Close
        Exit Function
    Else
            Pos = Selection.End
            ActiveDocument.Range(0, Pos).Select
            Selection.Cut
    Call Cut_Symbols
            Selection.Paste: Selection.TypeParagraph
            DocPath = ActiveDocument.Path
            On Error Resume Next
            ActiveDocument.SaveAs DocPath & "\" & Fname & ".txt"
            Debug.Print "Error: " & Err.Number & Err.Description
            ActiveDocument.Close
    End If
End Function
'***********************************************************************************************************************
'**********************************************CUT_SYMBOLS**********************************************************
'***********************************************************************************************************************
Function Cut_Symbols() ' Убераем лишние символы
    Const wdFindAsk = 2: Const wdReplaceAll = 2
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p": .Replacement.Text = "": .Forward = True: .Wrap = wdFindAsk: .Format = False: .MatchCase = False
        .MatchWholeWord = False: .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "-----------------------": .Replacement.Text = "^p": .Forward = True: .Wrap = wdFindAsk: .Format = False: .MatchCase = False
        .MatchWholeWord = False: .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
End Function
'***********************************************************************************************************************
'***********************************************************************************************************************
Function Name_Selector(ByVal strFname As String, ByVal Path As String)
         Select Case strFname
                     Case "02.lst", "02_1_02_2.lst", "09.lst", "14_1_14.lst", "15.lst", "16.lst", "17.lst", "18.lst", "19.lst", "20.lst", "21.lst", "22.lst", "23.lst", "24.lst", "29.lst", "27.lst"
                            Call Cut_Header(Path)
                            Call Разрезка1(Path)
                        Case "11.lst", "11_1.lst"
                            Call Cut_Header(Path)
                            Call Разрезка2(Path)
                        Case "47.lst"
                            Call Cut_Header(Path)
                            Call Разрезка3(Path)
                        Case "12.lst", "13.lst"
                            Call Cut_Header(Path)
                            Call Разрезка4(Path)
                        Case "52.lst", "53.lst", "56.lst"
                            Call Cut_Header(Path)
                            Call Разрезка5(Path)
                        Case "01.lst", "03.lst", "03_1.lst", "03_2.lst", "04.lst", "04_1.lst", "04_2.lst", "04_3.lst", "05.lst", "06.lst", "08.lst", "07.lst", "25.lst", "26.lst", "30.lst", "30_1.lst", "30_3.lst", "33_2.lst", "33_1.lst", "33.lst", "32_2.lst", "32_1.lst", "32.lst", "30_5.lst", "30_4.lst", "39.lst", "39_1.lst", "39_2.lst", "40.lst", "40_1.lst", "41.lst"
                            Call Разрезка_1(Path)
                        Case "10.lst", "10_1.lst", "10_2.lst", "10_3.lst"
                            Call Разрезка_2(Path)
                        Case "34_1.lst", "36.lst", "35_1.lst", "34_3.lst"
                            Call Разрезка_3(Path)
                        Case "38.lst"
                            Call Разрезка_4(Path)
                        Case "42.lst"
                            Call Разрезка_5(Path)
                        Case "43.lst", "43_6.lst", "43_1.lst", "43_2.lst", "43_3.lst", "43_4.lst", "43_5.lst"
                            Call Разрезка_6(Path)
                        Case "45.lst"
                            Call Разрезка_7(Path)
                        Case "46.lst"
                            Call Разрезка_8(Path)
                        Case "46_1.lst"
                            Call Разрезка_9(Path)
                        Case "48_1.lst"
                            Call Разрезка_10(Path)
                        Case "54.lst", "55.lst", "37.lst", "35.lst", "34_2.lst", "31.lst", "31_1.lst", "33_3.lst", "33_4.lst", "33_5.lst", "34.lst", "32_3.lst", "32_4.lst", "32_5.lst", "31_2.lst", "31_3.lst", "31_4.lst", "29_1.lst", "28.lst"
                            Call Разрезка_11(Path)
                        Case Else
                            MsgBox "САМИ РАЗРЕЗАЙТЕ Я НЕ ЗНАЮ ТАКОГО"
                    End Select
End Function
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_1 (С ФАБУЛАМИ)*************************************************
'***********************************************************************************************************************
Sub Разрезка1(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    maxlen = Len(file)
    Fname = Left(file, maxlen - 4)
    Fname = Fname & ".txt"
    If Dir(Fname) <> "" Then
        BookXL.Workbooks.OpenText filename:=Fname, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(35, 1) _
        , Array(44, 1), Array(48, 1), Array(50, 1), Array(52, 1), Array(54, 1), Array(56, 1), Array( _
        58, 1), Array(60, 1), Array(62, 1), Array(64, 1), Array(81, 1), Array(84, 1), Array(88, 1), _
        Array(92, 1), Array(96, 1), Array(100, 1), Array(104, 1), Array(108, 1), Array(112, 1), Array( _
        116, 1), Array(120, 1), Array(124, 1), Array(128, 1), Array(132, 1), Array(144, 1), Array( _
        147, 1), Array(151, 1), Array(162, 1), Array(175, 1), Array(179, 1), Array(183, 1), Array( _
        187, 1)), TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete:
        LastRow = BookXL.Range("A65536").End(xlUp).Row
        BookXL.Rows(LastRow).Delete
        Call Data_Fab_Copy
    End If
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_2 (С ФАБУЛАМИ)*************************************************
'***********************************************************************************************************************
Sub Разрезка2(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    maxlen = Len(file)
    Fname = Left(file, maxlen - 4)
    Fname = Fname & ".txt"
    If Dir(Fname) <> "" Then
    BookXL.Workbooks.OpenText filename:=Fname, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(49, 1), Array(51, 1), Array(54, 1), Array(58, 1), Array(61, 1), Array(63, 1), Array( _
    67, 1), Array(69, 1), Array(71, 1), Array(73, 1), Array(75, 1), Array(77, 1), Array(79, 1), _
    Array(81, 1), Array(85, 1), Array(87, 1), Array(89, 1), Array(91, 1), Array(93, 1), Array( _
    95, 1), Array(112, 1), Array(115, 1), Array(124, 1), Array(127, 1), Array(130, 1), Array(133 _
    , 1), Array(136, 1), Array(139, 1), Array(148, 1), Array(165, 1)), TrailingMinusNumbers:=True
    BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows(LastRow).Delete
    Call Data_Fab_Copy
    End If
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_3 (С ФАБУЛАМИ)*************************************************
'***********************************************************************************************************************
Sub Разрезка3(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    maxlen = Len(file)
    Fname = Left(file, maxlen - 4)
    Fname = Fname & ".txt"
    If Dir(Fname) <> "" Then
    BookXL.Workbooks.OpenText filename:=Fname, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(42, 2), Array(59, 1), Array(62, 1), Array(71, 1), Array(74, 1), Array(83, 1), Array( _
    85, 1), Array(87, 1), Array(96, 1), Array(99, 1), Array(108, 1)), TrailingMinusNumbers:=True
    BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows(LastRow).Delete
    Call Data_Fab_Copy
    End If
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_4 (С ФАБУЛАМИ)*************************************************
'***********************************************************************************************************************
Sub Разрезка4(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    maxlen = Len(file)
    Fname = Left(file, maxlen - 4)
    Fname = Fname & ".txt"
    If Dir(Fname) <> "" Then
    BookXL.Workbooks.OpenText filename:=Fname, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(49, 1), Array(51, 1), Array(54, 1), Array(57, 1), Array(59, 1), Array(63, 1), Array( _
    65, 1), Array(67, 1), Array(69, 1), Array(71, 1), Array(73, 1), Array(75, 1), Array(79, 1), _
    Array(81, 1), Array(83, 1), Array(85, 1), Array(87, 1), Array(89, 1), Array(106, 1), Array( _
    109, 1), Array(118, 1), Array(121, 1), Array(124, 1), Array(127, 1), Array(130, 1), Array( _
    133, 1), Array(142, 1), Array(150, 1), Array(161, 1)), TrailingMinusNumbers:=True
    BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows(LastRow).Delete
    Call Data_Fab_Copy
    End If
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_5 (С ФАБУЛАМИ)*************************************************
'***********************************************************************************************************************
Sub Разрезка5(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    maxlen = Len(file)
    Fname = Left(file, maxlen - 4)
    Fname = Fname & ".txt"
    If Dir(Fname) <> "" Then
    BookXL.Workbooks.OpenText filename:=Fname, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(49, 1), Array(53, 1), Array(55, 1), Array(57, 1), Array(59, 1), Array(61, 1), Array( _
    63, 1), Array(65, 1), Array(67, 1), Array(70, 1), Array(79, 1), Array(81, 1), Array(98, 1), _
    Array(101, 1), Array(104, 1)), TrailingMinusNumbers:=True
    BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows(LastRow).Delete
    Call Data_Fab_Copy
    End If
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_1 (БЕЗ ФАБУЛ)***************************************************
'***********************************************************************************************************************
Sub Разрезка_1(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(35, 1 _
        ), Array(44, 1), Array(48, 1), Array(50, 1), Array(52, 1), Array(54, 1), Array(56, 1), Array _
        (58, 1), Array(60, 1), Array(62, 1), Array(64, 2), Array(81, 1), Array(84, 1), Array(88, 1), _
        Array(92, 1), Array(96, 1), Array(100, 1), Array(104, 1), Array(108, 1), Array(112, 1), _
        Array(116, 1), Array(120, 1), Array(124, 1), Array(128, 1), Array(132, 1), Array(144, 1), _
        Array(147, 1), Array(151, 1), Array(162, 1), Array(175, 1), Array(179, 1), Array(183, 1), _
        Array(187, 1)), TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
        LastRow = BookXL.Range("A65536").End(xlUp).Row
        BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
        Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_2****************************************************************
'***********************************************************************************************************************
Sub Разрезка_2(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=xlWindows, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:= _
        Array(Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), _
        Array(40, 1), Array(49, 1), Array(53, 1), Array(55, 1), Array(57, 1), Array(59, 1), Array( _
        61, 1), Array(63, 1), Array(65, 1), Array(67, 1), Array(71, 1), Array(75, 1), Array(78, 1), _
        Array(87, 1)), TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_3****************************************************************
'***********************************************************************************************************************
Sub Разрезка_3(ByVal file As String)
        Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
        BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1 _
        ), Array(44, 1), Array(46, 1), Array(48, 1), Array(50, 1), Array(52, 1), Array(54, 1), Array _
        (56, 1), Array(58, 1), Array(61, 1), Array(70, 1), Array(73, 1), Array(76, 1), Array(79, 1)) _
        , TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_4****************************************************************
'***********************************************************************************************************************
Sub Разрезка_4(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(35, 1 _
        ), Array(44, 1), Array(48, 1), Array(50, 1), Array(52, 1), Array(54, 1), Array(56, 1), Array _
        (58, 1), Array(60, 1), Array(62, 1), Array(64, 2), Array(81, 1), Array(84, 1), Array(88, 1), _
        Array(92, 1), Array(96, 1), Array(100, 1), Array(104, 1), Array(108, 1), Array(112, 1), _
        Array(116, 1), Array(120, 1), Array(124, 1), Array(128, 1), Array(132, 1), Array(144, 1), _
        Array(147, 1), Array(151, 1), Array(162, 1), Array(175, 1), Array(179, 1), Array(183, 1), _
        Array(187, 1), Array(189, 1), Array(192, 1), Array(194, 1), Array(197, 1)), _
        TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_5****************************************************************
'***********************************************************************************************************************
Sub Разрезка_5(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=xlWindows, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:= _
        Array(Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), _
        Array(35, 1), Array(37, 1), Array(39, 1), Array(41, 1), Array(43, 1), Array(45, 1), Array( _
        47, 1), Array(49, 1), Array(51, 1), Array(71, 1), Array(81, 1), Array(95, 1), Array(104, 1), _
        Array(107, 1), Array(109, 1), Array(111, 1), Array(115, 1), Array(125, 1), Array(127, 1), _
        Array(129, 1)), TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_6****************************************************************
'***********************************************************************************************************************
Sub Разрезка_6(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(34, 1 _
        ), Array(37, 1), Array(46, 1), Array(50, 1), Array(52, 1), Array(56, 1), Array(58, 1), Array _
        (60, 1), Array(62, 1), Array(64, 1), Array(66, 1), Array(68, 1), Array(70, 1), Array(74, 1), _
        Array(77, 1), Array(79, 1), Array(81, 1), Array(83, 1), Array(91, 1), Array(95, 1), Array( _
        99, 1), Array(103, 1), Array(107, 1), Array(109, 1), Array(113, 1), Array(116, 1), Array( _
        118, 2), Array(135, 1), Array(138, 1), Array(141, 1), Array(151, 1), Array(153, 1), Array( _
        163, 1), Array(165, 1), Array(175, 1), Array(185, 1), Array(195, 1), Array(198, 1), Array( _
        208, 1), Array(210, 1), Array(220, 1), Array(226, 1), Array(235, 1), Array(238, 1), Array( _
        248, 1), Array(258, 1), Array(261, 1), Array(271, 1), Array(274, 1), Array(284, 1), Array( _
        287, 1), Array(297, 1), Array(300, 1), Array(303, 1), Array(307, 1), Array(316, 1)), _
        TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy

End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_7****************************************************************
'***********************************************************************************************************************
Sub Разрезка_7(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=xlWindows, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:= _
        Array(Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), _
        Array(40, 1), Array(42, 2), Array(59, 1), Array(62, 1), Array(71, 1), Array(74, 1), Array( _
        83, 1), Array(85, 1), Array(87, 1), Array(96, 1), Array(99, 1), Array(108, 1)), _
        TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_8****************************************************************
'***********************************************************************************************************************
Sub Разрезка_8(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1 _
        ), Array(49, 1), Array(53, 1), Array(55, 1), Array(57, 1), Array(59, 1), Array(61, 1), Array _
        (63, 1), Array(65, 1), Array(67, 1), Array(70, 1)), TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
   BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_9****************************************************************
'***********************************************************************************************************************
Sub Разрезка_9(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1 _
        ), Array(44, 1), Array(46, 1), Array(48, 1), Array(50, 1), Array(52, 1), Array(54, 1), Array _
        (56, 1), Array(58, 1), Array(61, 1), Array(70, 1), Array(75, 1), Array(85, 1)), _
        TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_10****************************************************************
'***********************************************************************************************************************
Sub Разрезка_10(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(56, 1 _
        ), Array(68, 1), Array(80, 1), Array(82, 1), Array(91, 1), Array(95, 1), Array(99, 1), Array _
        (103, 1), Array(107, 1), Array(111, 1), Array(115, 1), Array(119, 1), Array(123, 1), Array( _
        125, 1), Array(133, 1), Array(135, 1), Array(138, 1), Array(147, 1)), _
        TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub
'***********************************************************************************************************************
'*****************************************РАЗРЕЗКА_11****************************************************************
'***********************************************************************************************************************
Sub Разрезка_11(ByVal file As String)
    Dim LastRow As Long
    Const xlFixedWidth = 2
    Const xlUp = -4162
    BookXL.Workbooks.OpenText filename:=file, Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
        Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(46, 1 _
        ), Array(56, 1), Array(67, 1), Array(76, 1), Array(78, 1), Array(82, 1), Array(84, 1), Array _
        (86, 1), Array(90, 1), Array(94, 1), Array(98, 1), Array(102, 1), Array(108, 1), Array(112, _
        1), Array(116, 1), Array(120, 1), Array(124, 1), Array(127, 1), Array(131, 1), Array(139, 1 _
        ), Array(143, 1), Array(147, 1), Array(151, 1), Array(155, 1), Array(161, 1), Array(166, 1) _
        , Array(175, 1), Array(178, 1), Array(184, 1)), TrailingMinusNumbers:=True
        BookXL.Columns("B:B").Delete
    LastRow = BookXL.Range("A65536").End(xlUp).Row
    BookXL.Rows("" & LastRow - 1 & ":" & LastRow + 1).Delete
    Call Data_Copy
End Sub

