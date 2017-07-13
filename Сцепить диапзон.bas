Attribute VB_Name = "Concatenate_Range"
Option Explicit
Public Function яжеохрэдхюоюгнм(ByRef дхюоюгнм As Excel.Range, Optional ByVal пюгдекхрекэ As String = "") As String
Dim rCell As Range
Dim MergeText As String
For Each rCell In дхюоюгнм
    If rCell.Text <> "" Then
    MergeText = MergeText & пюгдекхрекэ & rCell.Text
    End If
Next
MergeText = Mid(MergeText, Len(пюгдекхрекэ) + 1)
яжеохрэдхюоюгнм = MergeText
End Function

