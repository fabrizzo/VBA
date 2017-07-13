Attribute VB_Name = "Module2"
Option Explicit

Option Explicit
Sub t()
Dim rRange As Range
Dim cell As Object
Set rRange = Range("EM2:EM705")
For Each cell In rRange
If cell.Value = "Color_171_red" Then
Range("CD" & cell.Row).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
 End With
 End If
 If cell.Value = "Color_172_red" Then
Range("CN" & cell.Row).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
 End With
 End If
 If cell.Value = "Color_173_red" Then
Range("CX" & cell.Row).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
 End With
 End If
 If cell.Value = "Color_174_red" Then
Range("DH" & cell.Row).Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
 End With
 End If
 Next cell
End Sub

