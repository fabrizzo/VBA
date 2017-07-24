Attribute VB_Name = "Module1"
Sub ShowRange()
Dim msg As String
Dim r As Integer
Dim c As Integer
msg = ""
For r = 1 To 20
For c = 1 To 8
msg = msg & Cells(r, c) & vbTab
Next c
msg = msg & vbCrLf
Next r
MsgBox msg
End Sub
