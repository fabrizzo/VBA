Attribute VB_Name = "Module1"
Function LastInColumn(rng As Range)
Application.Volatile
Set LastCell = rng.Parent.Cells(Rows.Count, rng.Column).End(xlUp)
LastInColumn = LastCell.Value
If IsEmpty(LastCell) Then LastInColumn = ""
If rng.Parent.cell(Rows.Count, rng.Column) <> "" Then _
LastInColumn = rng.Parent.Cells(Rows.Count, rng.Column)
End Function

