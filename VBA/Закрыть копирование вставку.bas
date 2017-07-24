Attribute VB_Name = "Module1"
Option Explicit

Private Sub Worksheet_Activate()
' ћакрос записан 13.12.2010 ( √еоргиевич)
'скрывает команды меню "вырезать", "копировать", _
"вставить", "специальна€ вставка"
'в контектном меню €чейки
With CommandBars("Cell")
.Controls(1).Enabled = False    'вырезать
.Controls(2).Enabled = False    'копировать
.Controls(3).Enabled = False    'вставить
.Controls(4).Enabled = False    'специальна€ вставка
End With
'в контектном меню столба
With CommandBars("Column")
.Controls(1).Enabled = False    'вырезать
End With
'в контектном меню строки
With CommandBars("Row")
.Controls(2).Enabled = False    'копировать
End With
'"ѕравка" наиглавнейшего меню
With CommandBars("Worksheet Menu Bar")
.Controls(2).Enabled = False
End With

Exit Sub

'вернуть всЄ обратно
With CommandBars("Cell")
.Controls(1).Enabled = True    'вырезать
.Controls(2).Enabled = True    'копировать
.Controls(3).Enabled = True    'вставить
.Controls(4).Enabled = True    'специальна€ вставка
End With
'в контектном меню столба
With CommandBars("Column")
.Controls(1).Enabled = True    'вырезать
End With
'в контектном меню строки
With CommandBars("Row")
.Controls(2).Enabled = True    'копировать
End With
'"ѕравка" наиглавнейшего меню
With CommandBars("Worksheet Menu Bar")
.Controls(2).Enabled = True
End With

End Sub




