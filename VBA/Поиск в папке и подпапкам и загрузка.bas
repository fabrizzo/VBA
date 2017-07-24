Attribute VB_Name = "Module1"
Option Explicit
Sub FileList()
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
    BrowseFolder = CStr(V)
    ListFilesinFolder BrowseFolder, True
    Application.ScreenUpdating = True
    Application.StatusBar = True
    Application.DisplayAlerts = True
End Sub
Sub ListFilesinFolder(ByVal SourceFolderName As String, ByVal IncludeSubFolders As Boolean)
    Dim FSO As Object: Dim SourceFolder As Object: Dim SubFolder As Object: Dim OldFolder As String: Dim FileItem As Object
    Dim maxlen, Pos As Integer: Dim Fname, strFileName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.getfolder(SourceFolderName)
    If Right(CStr(SourceFolder), 3) = "Название папки удаления" Then
        DeleteFilesAndDirectory (CStr(SourceFolder))
        Set FileItem = Nothing
        Set SourceFolder = Nothing
        Set FSO = Nothing
        Exit Sub
    Else
        For Each FileItem In SourceFolder.Files
            maxlen = Len(FileItem)
            Pos = InStrRev(FileItem, "\")
            Fname = Right(FileItem, maxlen - Pos)
            maxlen = Len(Fname)
            strFileName = Left(Fname, maxlen - 4)
        Select Case strFileName
            Case "02", "02_1_02_2", "09", "11", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "27", "29"
                Open_word_and_Clean (FileItem)
            Case "11", "11_1"
                Open_word_and_Clean (FileItem)
            Case "12", "13"
                Open_word_and_Clean (FileItem)
            Case "47"
                Open_word_and_Clean (FileItem)
            Case Else
                If Dir(FileItem) <> "" Then
                 Kill FileItem
                End If
        End Select
        Next FileItem
        If IncludeSubFolders Then
            For Each SubFolder In SourceFolder.SubFolders
                ListFilesinFolder SubFolder.Path, True
            Next SubFolder
        End If
    End If
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub
Function DeleteFilesAndDirectory(ByVal DelPath As String)
    Dim CleanPath, RmPath As String
    CleanPath = DelPath
    RmPath = DelPath + "\*.*"
    If Dir(RmPath) <> "" Then
        Kill RmPath
        RmDir CleanPath
    Else
        RmDir CleanPath
    End If
End Function
