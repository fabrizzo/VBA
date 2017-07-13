Attribute VB_Name = "Module1"
Function MAXALLSHEETS(cell)
Dim MaxVal As Double
Dim Addr As String
Dim Wksht As Object
Application.Volatile
Addr = cell.Range("A1").Address
MaxVal = -9.9E+307
For Each Wksht In cell.Parent.Parent.Worksheets
If Wksht.Name = cell.Parent.Name And _
Addr = Application.Caller.Address Then
Else
If IsNumeric(Wksht.Range(Addr)) Then
If Wksht.Range(Addr) > MaxVal Then
MaxVal = Wksht.Range(Addr).Value
End If
End If
End If
Next Wksht
If MaxVal = -9.9E+307 Then MaxVal = 0
MAXALLSHEETS = MaxVal
End Function
