Attribute VB_Name = "Module1"
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Application.CutCopyMode = xlCut Then
Application.CutCopyMode = False
ElseIf Application.CutCopyMode = xlCopy Then
Application.CutCopyMode = False
End If
End Sub
