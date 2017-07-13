Attribute VB_Name = "Module1"
Option Compare Database
Dim stVal As Integer
Dim fieldnum As Integer
Private Sub st_add_Click()
stVal = Me.st.Value
fieldnum = Me.num_stat.Value
If IsNull(Me.st.Value) Then
Exit Sub
ElseIf fieldnum < 2 Then
Exit Sub
Else
DoCmd.Close acForm, Me.Name, acSaveYes
Call NewControls
End If
End Sub
Sub NewControls()
    Dim ctl As Control
    Dim ctlLabel As Control, ctlText As Control:    Dim ctlLabel1 As Control, ctlText1 As Control:    Dim ctlLabel2 As Control, ctlText2 As Control:    Dim ctlLabel3 As Control, ctlText3 As Control
    Dim ctlLabel4 As Control, ctlText4 As Control:    Dim intDataX As Integer, intDataY As Integer:    Dim intLabelX As Integer, intLabelY As Integer
    Dim intLabelX1, intLabelX2, intLabelX3, intLabelX4 As Integer:    Dim intDataX1, intDataX2, intDataX3, intDataX4 As Integer
    Dim ctlButton As Control, ctlCheckBox As Control, ctlLabelCheck As Control, ctlButton1 As Control
    Dim ctlForm1Name As Control, ctlForm11Name As Control, ctlForm12Name As Control, ctlForm2Name As Control, ctlForm3Name As Control, ctlForm4Name As Control, ctlForm5Name As Control, ctlForm6Name As Control
    
        DoCmd.OpenForm "Form1", acDesign, , , acFormEdit, acHidden
        Do While Forms!Form1.Controls.Count > 0
            For Each ctl In Forms!Form1.Controls
                DeleteControl "Form1", ctl.Name
                Next ctl
            Loop
        j = 0
        For i = 1 To stVal
            intLabelX = 10:       intLabelY = 100 + j
            intDataX = 1000:    intDataY = 100 + j
            intLabelX1 = 2800: intDataX1 = 3630
            intLabelX2 = 5500: intDataX2 = 6350
            intLabelX3 = 8200: intDataX3 = 9100
            intLabelX4 = 11100: intDataX4 = 12150
            Set ctlText = CreateControl("Form1", acTextBox, , "", "", intDataX, intDataY)
            Set ctlLabel = CreateControl("Form1", acLabel, , ctlText.Name, "Статья " & i & ":", intLabelX, intLabelY)
            ctlText.Name = "st" & i
            Set ctlText1 = CreateControl("Form1", acTextBox, , "", "", intDataX1, intDataY)
            Set ctlLabel1 = CreateControl("Form1", acLabel, , ctlText1.Name, "Знак " & i & ":", intLabelX1, intLabelY)
            ctlText1.Name = "zn" & i
            Set ctlText2 = CreateControl("Form1", acTextBox, , "", "", intDataX2, intDataY)
            Set ctlLabel2 = CreateControl("Form1", acLabel, , ctlText2.Name, "Часть " & i & ":", intLabelX2, intLabelY)
            ctlText2.Name = "ch" & i
            Set ctlText3 = CreateControl("Form1", acTextBox, , "", "", intDataX3, intDataY)
            Set ctlLabel3 = CreateControl("Form1", acLabel, , ctlText3.Name, "Пункт " & i & ":", intLabelX3, intLabelY)
            ctlText3.Name = "p1" & i
            Set ctlText4 = CreateControl("Form1", acTextBox, , "", "", intDataX4, intDataY)
            Set ctlLabel4 = CreateControl("Form1", acLabel, , ctlText4.Name, "Пункт1" & "/" & i & ":", intLabelX4, intLabelY)
            ctlText4.Name = "p2" & i
             j = j + 300
        Next i
        
        Set ctlButton = CreateControl("Form1", acCommandButton, , , , 100, intLabelY + 500, 1500, 400)
        ctlButton.OnClick = "[Event Procedure]":    ctlButton.Caption = "Запрос":    ctlButton.Name = "Query"
        
        Set ctlButton1 = CreateControl("Form1", acCommandButton, , , , 1600, intLabelY + 500, 1500, 400)
        ctlButton1.OnClick = "[Event Procedure]"
        ctlButton1.Caption = "Назад"
        ctlButton1.Name = "Comeback"
    
    Set ctlCheckBox = CreateControl("Form1", acCheckBox, , , , 1550, intLabelY + 1150, 300, 300): ctlCheckBox.Name = "CheckBox_EGS"
    Set ctlLabelCheck = CreateControl("Form1", acLabel, , ctlCheckBox.Name, "Условия ЕГС", 150, intLabelY + 1150)
    
    Set ctlForm1Name = CreateControl("Form1", acComboBox, , , , 150, intLabelY + 1700, 1400, 300): ctlForm1Name.Name = "Form1Unselect"
    Set ctlForm11Name = CreateControl("Form1", acComboBox, , , , 1650, intLabelY + 1700, 1400, 300): ctlForm11Name.Name = "Form11Unselect"
    Set ctlForm12Name = CreateControl("Form1", acComboBox, , , , 3150, intLabelY + 1700, 1400, 300): ctlForm12Name.Name = "Form12List"
    Set ctlForm2Name = CreateControl("Form1", acComboBox, , , , 4650, intLabelY + 1700, 1400, 300): ctlForm2Name.Name = "Form2List"
    Set ctlForm3Name = CreateControl("Form1", acComboBox, , , , 6150, intLabelY + 1700, 1400, 300): ctlForm3Name.Name = "Form3List"
    Set ctlForm4Name = CreateControl("Form1", acComboBox, , , , 7650, intLabelY + 1700, 1400, 300): ctlForm4Name.Name = "Form4List"
    Set ctlForm5Name = CreateControl("Form1", acComboBox, , , , 9150, intLabelY + 1700, 1400, 300): ctlForm5Name.Name = "Form5List"
    Set ctlForm6Name = CreateControl("Form1", acComboBox, , , , 10650, intLabelY + 1700, 1400, 300): ctlForm6Name.Name = "Form6List"
    
    Dim intPositionX, intPositionY As Integer
    Dim ctlComboFieldName1, ctlComboFieldName2, ctlComboFieldName3, ctlComboFieldName4, ctlComboFieldName5, ctlComboFieldName6, ctlComboFieldName7
    Dim k, r As Integer
    Dim fieldsin As Double
    k = 0
    r = 0
    For i = 1 To fieldnum
    intPositionX = 150 + r: intPositionY = 2100 + k + intLabelY
    Set ctlComboFieldName1 = CreateControl("Form1", acComboBox, , , , intPositionX, intPositionY, 2100, 300): ctlComboFieldName1.Name = "Spisok_Field" & i
    r = r + 2200
    If i / 7 = Fix(i / 7) Then
    k = k + 350
    r = 0
    End If
    Next i
    DoCmd.Restore
    DoCmd.OpenForm "Form1", acNormal
End Sub



