Attribute VB_Name = "Module1"
Sub test()
Dim dData()
With UserForm1.ListBox1
.RowSource = ""
.AddItem "����� 1"
.AddItem "����� 2"
.AddItem "����� 3"
.AddItem "����� 4"
.AddItem "����� 5"
.AddItem "����� 6"
.AddItem "����� 7"
.AddItem "����� 8"
End With
ReDim dData(1 To UserForm1.ListBox1.ListCount)
dData = UserForm1.ListBox1.List
For row = 1 To UserForm1.ListBox1.ListCount
Cells(1, row).Value = dData(row, 1)
Next row
End Sub
