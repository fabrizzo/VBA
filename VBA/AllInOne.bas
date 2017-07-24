Attribute VB_Name = "AllInOne"
Option Explicit
Sub Собратьфайлы()
    Dim MyFile As String
    Dim V As String: Dim BrowseFolder As String
    Dim ProName, InitialPath As String
    MyFile = UserForm1.TextBox1.Text
    Selection.WholeStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    ProName = ActiveDocument.Path
    InitialPath = ProName
    On Error Resume Next
    With Application.FileDialog(4)
        .Title = "Выберите папку"
        .InitialFileName = InitialPath
        .Show
      Err.Clear
      V = .SelectedItems(1)
      If Err.Number > 0 Then
      MsgBox "Не выбрали ничего"
      Exit Sub
      End If

    End With
    
    BrowseFolder = CStr(V)
    ListFilesinFolder BrowseFolder, True, MyFile
MsgBox "Готово", vbInformation
UserForm1.Hide
Application.Visible = True
End Sub
Sub ListFilesinFolder(ByVal SourceFolderName As String, ByVal IncludeSubFolders As Boolean, ByVal NewFile As String)
    Dim FSO As Object: Dim SourceFolder As Object: Dim SubFolder As Object: Dim OldFolder As String: Dim FileItem As Object
    Dim maxlen, Pos As Integer: Dim Fname, strFileName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.getfolder(SourceFolderName)
'**************************************************************************
    If UserForm1.CheckBox1.Value = True And UserForm1.CheckBox2.Value = False Then
        If Right(SourceFolder, 3) = "GOD" Then
                For Each FileItem In SourceFolder.Files
                    maxlen = Len(FileItem)
                    Pos = InStrRev(FileItem, "\")
                    Fname = Right(FileItem, maxlen - Pos)
                    If Fname = NewFile Then
                        Call Open_And_Copy(FileItem)
                    End If
                Next FileItem
        End If
     ElseIf UserForm1.CheckBox2.Value = True And UserForm1.CheckBox1.Value = True Then
                For Each FileItem In SourceFolder.Files
                    maxlen = Len(FileItem)
                    Pos = InStrRev(FileItem, "\")
                    Fname = Right(FileItem, maxlen - Pos)
                    If Fname = NewFile Then
                        Call Open_And_Copy(FileItem)
                    End If
                Next FileItem
    ElseIf UserForm1.CheckBox2.Value = True And UserForm1.CheckBox1.Value = False Then
                If Right(SourceFolder, 3) = "MEC" Then
                For Each FileItem In SourceFolder.Files
                    maxlen = Len(FileItem)
                    Pos = InStrRev(FileItem, "\")
                    Fname = Right(FileItem, maxlen - Pos)
                    If Fname = NewFile Then
                        Call Open_And_Copy(FileItem)
                    End If
                Next FileItem
            End If
    End If
        
        
'**************************************************************************
     If IncludeSubFolders Then
           For Each SubFolder In SourceFolder.SubFolders
           Call ListFilesinFolder(SubFolder.Path, True, NewFile)
           Next SubFolder
        End If
     
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub
Function Open_And_Copy(ByVal file As String)
    Documents.Open filename:=file, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
    If UserForm1.CheckBox3.Value = False Then
        With Selection.Find
            .Text = "//*=Q"    ' искомый текст
            .MatchWildcards = True              ' подстановочные знаки в окне поиска (* - любой текст)
            .Wrap = wdFindStop                  ' остановиться на найденном
                If .Execute = True Then
                    Selection.WholeStory
                    Selection.Copy
                    Windows("Собрать в один.docm").Activate
                     Selection.Paste
                End If
        End With
    Else
        Selection.WholeStory
        Selection.Copy
        Windows("Собрать в один.docm").Activate
        Selection.Paste
    End If
    Documents(file).Close
    
End Function
