Attribute VB_Name = "Module1"
Sub test()
Dim AllCells As Range, Cell As Range
Dim Nodupes As New Collection
On Error Resume Next
For Each Cell In Range("A1:A20")
    Nodupes.Add Cell.Value, CStr(Cell.Value)
Next Cell
On Error GoTo 0
j = 1
For Each Item In Nodupes
    Range("B" & j).Value = Item
    j = j + 1
Next Item
MsgBox "”никальных значений: " & Nodupes.Count
End Sub
