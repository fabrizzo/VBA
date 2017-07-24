Attribute VB_Name = "Module1"
Sub test()
Dim dData()
With UserForm1.ListBox1
.RowSource = ""
.AddItem "Опция 1"
.AddItem "Опция 2"
.AddItem "Опция 3"
.AddItem "Опция 4"
.AddItem "Опция 5"
.AddItem "Опция 6"
.AddItem "Опция 7"
.AddItem "Опция 8"
End With
ReDim dData(1 To UserForm1.ListBox1.ListCount)
dData = UserForm1.ListBox1.List
For row = 1 To UserForm1.ListBox1.ListCount
Cells(1, row).Value = dData(row, 1)
Next row
End Sub
