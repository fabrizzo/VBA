Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Sub ShowGraphic()
Dim FileName As String
FileName = "C:\Users\Bronxy\Pictures\25_3168.jpg"
Call ShellExecute(0&, vbNullString, FileName, _
vbNullString, vbNullString, vbNormalFocus)
End Sub
Sub OpenFileWeb()
Dim Url As String
Url = "https://www.yandex.ru"
Call ShellExecute(0&, vbNullString, Url, _
vbNullString, vbNullString, vbNormalFocus)
End Sub
Sub StartEmailOutlook()
Dim addr As String
addr = "mailto:bgates@microsoft.com"
Call ShellExecute(0&, vbNullString, addr, _
vbNullString, vbNullString, vbNormalFocus)
End Sub
