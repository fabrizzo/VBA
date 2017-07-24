Attribute VB_Name = "Module1"
Function ISLIKE(Text As String, pattern As String)
If Text Like pattern Then ISLIKE = True
Else
ISLIKE = False
End Function
