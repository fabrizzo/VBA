Attribute VB_Name = "CutAndCopytoFirst"
Option Explicit
Sub NewList()
    Application.ScreenUpdating = False
    Application.StatusBar = False
    Application.DisplayAlerts = False
    Dim V As String: Dim BrowseFolder As String
    Dim ProName, InitialPath As String
    ProName = ActiveWorkbook.Path
    InitialPath = ProName
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
    Range("A:A,D:L").Select
    Selection.NumberFormat = "@"
    Range("A1").Select
    
    Sheets("Fabula").Select: Sheets("Fabula").Copy
    If Dir(ProName & "\" & "Fabula.xlsx") = "" Then
    ActiveWorkbook.SaveAs filename:=ProName & "\" & "Fabula", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close SaveChanges:=False
    Else
    ActiveWorkbook.Close SaveChanges:=False
    End If
    Cells.ClearContents
    Range("A1").Select
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = True
    MsgBox "Готово", vbInformation
    Unload UserForm1
    ActiveWorkbook.Close SaveChanges:=False
End Sub
Sub ListFilesinFolder(ByVal SourceFolderName As String, ByVal IncludeSubFolders As Boolean)
    Dim FSO As Object: Dim SourceFolder As Object: Dim SubFolder As Object: Dim OldFolder As String: Dim FileItem As Object
    Dim maxlen, Pos As Integer: Dim Fname, strFileName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.getfolder(SourceFolderName)
'    If Right(CStr(SourceFolder), 3) = "GOD" Then
 '       DeleteFilesAndDirectory (CStr(SourceFolder))
 '       Set FileItem = Nothing
 '       Set SourceFolder = Nothing
 '       Set FSO = Nothing
 '       Exit Sub
 '   Else
        For Each FileItem In SourceFolder.Files
            maxlen = Len(FileItem)
            Pos = InStrRev(FileItem, "\")
            Fname = Right(FileItem, maxlen - Pos)
        Select Case Fname
            Case "02.txt", "02_1_02_2.txt", "09.txt", "14_1_14.txt", "15.txt", "16.txt", "17.txt", "18.txt", "19.txt", "20.txt", "21.txt", "22.txt", "23.txt", "24.txt", "29.txt", "27.txt"
                Call Разрезка1(FileItem, Fname)
            Case "11.txt", "11_1.txt"
                Call Разрезка2(FileItem)
            Case "47.txt"
                Call Разрезка3(FileItem)
            Case "12.txt", "13.txt"
                Call Разрезка4(FileItem)
            Case "52.txt", "53.txt", "56.txt"
                Call Разрезка5(FileItem)
        End Select
        Next FileItem
        If IncludeSubFolders Then
            For Each SubFolder In SourceFolder.SubFolders
                ListFilesinFolder SubFolder.Path, True
            Next SubFolder
        End If
 '   End If
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub
Sub Разрезка1(ByVal file As String, ByVal strShortName As String)
Attribute Разрезка1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Workbooks.OpenText filename:=file, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(35, 1) _
    , Array(44, 1), Array(48, 1), Array(50, 1), Array(52, 1), Array(54, 1), Array(56, 1), Array( _
    58, 1), Array(60, 1), Array(62, 1), Array(64, 1), Array(81, 1), Array(84, 1), Array(88, 1), _
    Array(92, 1), Array(96, 1), Array(100, 1), Array(104, 1), Array(108, 1), Array(112, 1), Array( _
    116, 1), Array(120, 1), Array(124, 1), Array(128, 1), Array(132, 1), Array(144, 1), Array( _
    147, 1), Array(151, 1), Array(162, 1), Array(175, 1), Array(179, 1), Array(183, 1), Array( _
    187, 1)), TrailingMinusNumbers:=True
    Columns("B:B").Delete: Columns("E:E").Delete: Columns("L:AE").Delete
    Columns("M:P").Delete: Rows("1:14").Delete: Columns("A:A").Delete:
    Columns("E:E").Select: Selection.Insert Shift:=xlLeft
    LastRow = Range("A65536").End(xlUp).Row
    Rows(LastRow).Delete: Range("A1").Select
    Select Case strShortName
    Case "02_1_02_2.txt"
        Rows("1:8").Delete
    Case "14_1_14.txt"
        Rows("1:8").Delete
    Case "09.txt"
        Rows("1:1").Delete
    End Select
    Call Data_Copy
End Sub
Sub Разрезка2(ByVal file As String)
Attribute Разрезка2.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Workbooks.OpenText filename:=file, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(49, 1), Array(51, 1), Array(54, 1), Array(58, 1), Array(61, 1), Array(63, 1), Array( _
    67, 1), Array(69, 1), Array(71, 1), Array(73, 1), Array(75, 1), Array(77, 1), Array(79, 1), _
    Array(81, 1), Array(85, 1), Array(87, 1), Array(89, 1), Array(91, 1), Array(93, 1), Array( _
    95, 1), Array(112, 1), Array(115, 1), Array(124, 1), Array(127, 1), Array(130, 1), Array(133 _
    , 1), Array(136, 1), Array(139, 1), Array(148, 1), Array(165, 1)), TrailingMinusNumbers:=True
    Columns("A:B").Delete: Columns("F:J").Delete: Columns("L:AA").Delete
    Columns("M:M").Delete: Rows("1:14").Delete: Range("A1").Select
    LastRow = Range("A65536").End(xlUp).Row
    Rows(LastRow).Delete
    Range("A1").Select
    maxlen = Len(ActiveWorkbook.Name)
    Fname = Left(ActiveWorkbook.Name, maxlen - 4)
    Fname = Right(Fname, 2)
    Call Data_Copy

End Sub
Sub Разрезка3(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Workbooks.OpenText filename:=file, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(42, 1), Array(59, 1), Array(62, 1), Array(71, 1), Array(74, 1), Array(83, 1), Array( _
    85, 1), Array(87, 1), Array(96, 1), Array(99, 1), Array(108, 1)), TrailingMinusNumbers:=True
    Columns("A:B").Delete: Columns("E:N").Delete
    Columns("E:E").Select: Selection.Insert Shift:=xlLeft
    Selection.Insert Shift:=xlLeft: Selection.Insert Shift:=xlLeft
    Selection.Insert Shift:=xlLeft: Selection.Insert Shift:=xlLeft
    Selection.Insert Shift:=xlLeft: Columns("D:D").Select
    Selection.Insert Shift:=xlLeft
    Rows("1:14").Delete: Range("A1").Select
    LastRow = Range("A65536").End(xlUp).Row
    Rows(LastRow).Delete: Range("A1").Select
    maxlen = Len(ActiveWorkbook.Name)
    Fname = Left(ActiveWorkbook.Name, maxlen - 4)
    Fname = Right(Fname, 2)
    Call Data_Copy
End Sub
Sub Разрезка4(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Workbooks.OpenText filename:=file, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(49, 1), Array(51, 1), Array(54, 1), Array(57, 1), Array(59, 1), Array(63, 1), Array( _
    65, 1), Array(67, 1), Array(69, 1), Array(71, 1), Array(73, 1), Array(75, 1), Array(79, 1), _
    Array(81, 1), Array(83, 1), Array(85, 1), Array(87, 1), Array(89, 1), Array(106, 1), Array( _
    109, 1), Array(118, 1), Array(121, 1), Array(124, 1), Array(127, 1), Array(130, 1), Array( _
    133, 1), Array(142, 1), Array(150, 1), Array(161, 1)), TrailingMinusNumbers:=True
    Columns("A:B").Delete: Columns("F:I").Delete
    Columns("L:Z").Delete: Columns("M:N").Delete
    Rows("1:14").Delete:  Range("A1").Select
    LastRow = Range("A65536").End(xlUp).Row
    Rows(LastRow).Delete: Range("A1").Select
    maxlen = Len(ActiveWorkbook.Name)
    Fname = Left(ActiveWorkbook.Name, maxlen - 4)
    Fname = Right(Fname, 2)
    Call Data_Copy

End Sub
Sub Разрезка5(ByVal file As String)
    Dim Fname As String: Dim maxlen As Integer: Dim LastRow As Long
    Workbooks.OpenText filename:=file, _
    Origin:=1251, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array( _
    Array(0, 1), Array(5, 1), Array(6, 1), Array(10, 2), Array(27, 1), Array(31, 1), Array(40, 1) _
    , Array(49, 1), Array(53, 1), Array(55, 1), Array(57, 1), Array(59, 1), Array(61, 1), Array( _
    63, 1), Array(65, 1), Array(67, 1), Array(70, 1), Array(79, 1), Array(81, 1), Array(98, 1), _
    Array(101, 1), Array(104, 1)), TrailingMinusNumbers:=True
    Columns("A:B").Delete
    Columns("L:N").Delete
    Columns("M:P").Delete
    Rows("1:14").Delete:  Range("A1").Select
    LastRow = Range("A65536").End(xlUp).Row
    Rows(LastRow).Delete: Range("A1").Select
    maxlen = Len(ActiveWorkbook.Name)
    Fname = Left(ActiveWorkbook.Name, maxlen - 4)
    Fname = Right(Fname, 2)
    Call Data_Copy

End Sub
Function Data_Copy()
    Dim LastRow, Row As Long
    Dim Pos As Integer
    Dim strFileName As String
    Dim intMax As Integer
    Dim intMaxLenFileName As Integer
    Row = 0: Row = Range("A" & Rows.Count).End(xlUp).Row
    intMaxLenFileName = Len(ActiveWorkbook.Path)
    strFileName = Right(ActiveWorkbook.Path, intMaxLenFileName - 13)
    intMax = Len(strFileName)
    Pos = InStrRev(strFileName, "\")
    strFileName = Right(strFileName, intMax - Pos)
    Range("N1:N" & Row).Value = strFileName & " " & ActiveWorkbook.Name
    Range("A1:N" & Row).Copy
    Workbooks(1).Worksheets("Fabula").Activate
    If Range("A1").Value = "" Then
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Workbooks(2).Activate
        ActiveWorkbook.Close SaveChanges:=False
    Else
        Range("A1").Select
        LastRow = 0
        LastRow = Range("A" & Rows.Count).End(xlUp).Row
        LastRow = LastRow + 1
        Range("A" & LastRow).PasteSpecial
        Application.CutCopyMode = False
        Workbooks(2).Activate
        ActiveWorkbook.Close SaveChanges:=False
    End If
End Function

