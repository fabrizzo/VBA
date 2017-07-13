Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Dim BookXL As Object
'*******************************************************************************************************************
'********************************************������**************************************************************
'*******************************************************************************************************************
Private Sub Sud_Button_Click() '������ ������������ � ��� ������������
Call FolderCheck(2) '�������� ������� � ������� 2
End Sub
Private Sub Button_economic_Click() '������������ ������������� ��������������
Call FolderCheck(1) '�������� ������� � ������� 1
End Sub
'*******************************************************************************************************************
'*********************************�������� ������������� �����******************************************
'*******************************************************************************************************************
Private Function FileExists(ByVal fname2 As String) As Boolean '���� ����� ���, �� ������� ������ false
    FileExists = (Dir(fname2) = "")
End Function
'*******************************************************************************************************************
'*********************************�������� ������������� �����******************************************
'*******************************************************************************************************************
Private Function FolderExists(Fname As String) As Boolean '���� ����� ���, �� ������� ������ false
    On Error Resume Next
    Fname = GetAttr(Fname) And 0 '���������� ������� ����� ���� ��� ���� ����� ������ ������
    FolderExists = (Err = 0)
End Function
'*******************************************************************************************************************
'*****************************************�������� �����******************************************************
'*******************************************************************************************************************
Function FolderCheck(num As Integer) '�������� ������������� �����, ���� ���� ������� ��
    Dim msg As String
    On Error GoTo Error_Handler
    Select Case num '����� �������
        Case 1
                If FolderExists("C:\Reports\") = True Then '���������� �� ����� �����
                    Call CreationOfSQL(num) '�������� ������� �������� �������
                Else
                    MkDir "C:\Reports"
                    Call CreationOfSQL(num)
                End If
        Case 2
                If FolderExists("C:\Reports\") = True Then
                    Call CreationOfSQL(num) '�������� ������� �������� �������
                Else
                    MkDir "C:\Reports"
                    Call CreationOfSQL(num)
                End If
    End Select
    
Exit Function
Error_Handler: '��������� ������
msg = "������ � [" & Application.VBE.ActiveCodePane.CodeModule & "/FolderCheck]" & " ������ #" & Str(Err.Number) & " � ������� [" & Err.Source & "] ��������: " & Err.Description & Chr(13) & "� ������ ������������� ������ ���������a ���������� � ������������"
MsgBox msg, vbInformation, "��!!!"
End Function
'*******************************************************************************************************************
'*******************************�������� ������� � ���� ������********************************************
'*******************************************************************************************************************
Function CreationOfSQL(intNums As Integer)
Dim msg As String
On Error GoTo Error_Handler
    Dim qdf As QueryDef '������ ������
    Dim strSQL, strfPath, strReportName As String
    On Error Resume Next
    Select Case intNums '����� ������
    Case 1 '���������
        strfPath = "C:\Reports\Ecomonic.xls" '���� ������
        strReportName = "ecomonicQry" '��� ������
        strSQL = "SELECT Form1.f1_1kod, Form1.f1_3num, Form1.f1_4num, Form1.f1_111, Form1.f1_13s " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13z " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13ch " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13p1_1 " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13p1_2 " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13p1_3 " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13p1_4 " & Chr(38) & Chr(34) & " " & Chr(34) & Chr(38) & " Form1.f1_13p1_5 AS Article, Form11.f11_25k, Form11.f11_25d," & _
                  " Form1.f1_7d, Form1.f1_11d, Form1.f1_181 " & Chr(38) & " Form1.f1_18 AS f1_18, Form2.f2_261 " & Chr(38) & " Form2.f2_26 AS f2_26, Form4.f4_81 " & Chr(38) & " Form4.f4_8 AS f4_8, Form1.f1_20, Form2.f2_29, Form1.f1_22, Form2.f2_30, Form1.f1_24 & Form1.f1_241 " & Chr(38) & Chr(34) & " / " & Chr(34) & Chr(38) & " Form1.f1_242 " & Chr(38) & " Form1.f1_243 " & Chr(38) & Chr(34) & " / " & Chr(34) & Chr(38) & " Form1.f1_244 " & Chr(38) & " Form1.f1_245 AS f1_24, " & _
                  " Form2.f2_32_1 " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " Form2.f2_32_2 " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " Form2.f2_32_3 AS f2_32, Form5.f5_171 & Form5.f5_172 as f5_171, Form2.f2_4num, Form2.f2_34_1, Form2.f2_34_2, Form2.f2_34_3, Form2.f2_34_4, Form3.f3_8, Form3.f3_8nums, Form4.f4_10 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_101 AS f4_10_1, Form4.f4_102 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_103 AS f4_10_2, Form4.f4_104 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_105 AS f4_10_3, Form4.f4_106 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_107 AS f4_10_4, Form4.f4_11 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_111 AS f4_11_1, Form4.f4_112 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_113 AS f4_11_2, Form4.f4_114 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_115 AS " & _
                  " f4_11_3, Form4.f4_12, Form4.f4_121, Form4.f4_15, Form4.f4_16 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_161 AS f4_16_1, Form4.f4_162 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_163 AS f4_16_2, Form4.f4_164 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_165 AS f4_16_3, Form4.f4_166 " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " Form4.f4_167 AS f4_16_4, Form4.f4_32, Fabula.������ " & _
                  " FROM ((((((Form1 LEFT JOIN Form2 ON (Form1.f1_4num = Form2.f2_4num) AND (Form1.f1_3num = Form2.f2_3num)) LEFT JOIN Form3 ON (Form1.f1_4num = Form3.f1_4num) AND (Form1.f1_3num = Form3.f1_3num)) LEFT JOIN Form4 ON (Form1.f1_4num = Form4.f1_4num) AND (Form1.f1_3num = Form4.f1_3num)) LEFT JOIN Fabula ON (Form1.f1_4num = Fabula.���) AND (Form1.f1_3num = Fabula.[����� �����])) LEFT JOIN Form11 ON (Form1.f1_3num = Form11.f1_3num) AND (Form1.f1_4num = Form11.f1_4num)) LEFT JOIN Form5 ON (Form1.f1_3num = Form5.f5_3num))"
    Case 2 ' ������������ � ���
        strfPath = "C:\Reports\Obvinit.xls" '���� ������
        strReportName = "obvinQry" '��� ������
        strSQL = "SELECT Form1.f1_1kod, Form1.f1_3num, Form3.[f3_8], Form3.[f3_8num], [Form2].[f2_fam] " & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & " Left([Form2].[f2_imj],1)" & Chr(38) & Chr(34) & "." & Chr(34) & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & " Left([Form2].[f2_otc],1)" & Chr(38) & Chr(34) & "." & Chr(34) & " AS ���, [Form11].[f11_7s] " & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & " [Form11].[f11_7z] " & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & " [Form11].[f11_7ch] " & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & " [Form11].[f11_7p1_1] " & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & "[Form11].[f11_7p1_2]" & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & "[Form11].[f11_7p1_3]" & Chr(38) & Chr(34) & Chr(32) & Chr(34) & " & Form11.f11_7p1_4" & Chr(38) & Chr(34) & Chr(32) & Chr(34) & Chr(38) & "[Form11].[f11_7p1_5] AS ������, " & _
                 "IIf([Form11].[f11_18_1]=[Form2].[f2_13_1] And [Form11].[f11_18_2]=[Form2].[f2_13_2],[Form11].[f11_18_1],[Form11].[f11_18_1]" & Chr(38) & "[Form11].[f11_18_2] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [Form2].[f2_13_1] " & Chr(38) & " [Form2].[f2_13_2]) AS �����������, " & _
                 "IIf([Form11].[f11_12_1]=3 Or [Form11].[f11_12_2]=3 Or [Form11].[f11_12_3]=3 Or [f2_11]=1 Or [f2_11]=2,[Form11].[f11_12_1] " & Chr(38) & " [Form11].[f11_12_2] " & Chr(38) & "[Form11].[f11_12_3] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [f2_11]," & Chr(34) & Chr(34) & ") AS ������������������, " & _
                 "IIf([Form11].[f11_23] " & Chr(38) & " [Form11].[f11_231] In (" & Chr(34) & "02" & Chr(34) & "," & Chr(34) & "10" & Chr(34) & ") Or [Form11].[f11_232] " & Chr(38) & " [Form11].[f11_233]" & " In (" & Chr(34) & "02" & Chr(34) & "," & Chr(34) & "10" & Chr(34) & ") Or [Form2].[f2_181] " & Chr(38) & " [Form2].[f2_182]" & " In (" & Chr(34) & "02" & Chr(34) & "," & Chr(34) & "10" & Chr(34) & "), [Form11].[f11_23] " & Chr(38) & "  [Form11].[f11_231] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_232] " & Chr(38) & "  [Form11].[f11_233] " & Chr(38) & """" & "/" & """" & Chr(38) & " [Form2].[f2_181] " & Chr(38) & "  [Form2].[f2_182], [Form11].[f11_23] " & Chr(38) & "  [Form11].[f11_231] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_232] " & Chr(38) & "  [Form11].[f11_233] " & Chr(38) & """" & "/" & """" & Chr(38) & " [Form2].[f2_181] " & Chr(38) & "  [Form2].[f2_182]) AS ����_������, " & _
                 "IIf(([Form11].[f11_14_1]=1 Or [Form11].[f11_14_2]=1 Or [Form11].[f11_14_3]=1), 1," & Chr(34) & Chr(34) & ") AS �����_�����������_�11," & _
                 "IIf(([Form2].[f2_50_1]=1 Or [Form2].[f2_50_2]=1 Or [Form2].[f2_50_3]=1), 1 ," & Chr(34) & Chr(34) & ") AS �����_�����������_�2," & _
                 "IIf([Form11].[f11_15] In (1,2) Or [Form11].[f11_152] In (1,2) Or [Form11].[f11_154] In (1,2), 1 ," & Chr(34) & Chr(34) & ") AS �����_�������_�11, " & _
                 "IIf([Form2].[f2_45_1]>=1 Or [Form2].[f2_45_2]>=1, 1," & Chr(34) & Chr(34) & ") AS �����_�������_�2, " & _
                 "[Form11].[f11_13_1]" & Chr(38) & "[Form11].[f11_13_2]" & Chr(38) & """" & "/" & """" & Chr(38) & "[Form2].[f2_36_1]" & Chr(38) & "[Form2].[f2_36_2] AS [� ���������], " & _
                 "IIf([Form1].[f1_2111] & [Form1].[f1_211]<>" & Chr(34) & "00" & Chr(34) & "Or [Form11].[f11_911]" & Chr(38) & "[Form11].[f11_91]<>" & Chr(34) & "00" & Chr(34) & "Or [Form5].[f5_1811]" & Chr(38) & "[Form5].[f5_181]<>" & Chr(34) & "00" & Chr(34) & ",[Form1].[f1_2111] " & Chr(38) & " [Form1].[f1_211] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [Form11].[f11_911] " & Chr(38) & " [Form11].[f11_91] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [Form5].[f5_1811] " & Chr(38) & " [Form5].[f5_181]," & Chr(34) & Chr(34) & ") AS �����, " & _
                 "IIf([Form11].[f11_101]<>0 Or [Form2].[f2_38]<>0,[Form11].[f11_101] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [Form2].[f2_38]," & Chr(34) & Chr(34) & ") AS [� �������], " & _
                 "IIF([Form5].[f5_11]>0, [Form5].[f5_11]," & Chr(34) & Chr(34) & ") as f5_11, [Form11].[f11_28] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_281]  " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form11].[f11_282] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_283] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form11].[f11_284] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_285] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form11].[f11_286] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_287] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form11].[f11_288] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form11].[f11_289] AS f11_28, " & _
                 "[Form4].[f4_10] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_101] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_102] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_103] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_104] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_105] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_106] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_107] AS f4_10, " & _
                 "[Form4].[f4_11] " & Chr(38) & """" & "_" & """" & Chr(38) & "[Form4].[f4_111] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_112] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_113] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_114] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_115] AS f4_11, [Form4].[f4_15], " & _
                 "[Form4].[f4_16] " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " [Form4].[f4_161]" & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_162] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_163]  " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_164] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_165]  " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_166] " & Chr(38) & """" & "_" & """" & Chr(38) & " [Form4].[f4_167] AS f4_16, " & _
                 "IIF([Form4].[f4_32]>0,[Form4].[f4_32]," & Chr(34) & Chr(34) & ") as f4_32, [Form4].[f4_193] " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " [Form4].[f4_198] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_194] " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " [Form4].[f4_199] " & Chr(38) & """" & "/" & """" & Chr(38) & "[Form4].[f4_195] " & Chr(38) & Chr(34) & "_" & Chr(34) & Chr(38) & " [Form4].[f4_200] AS f4_19_3_3, " & _
                 "IIF([Form1].[f1_181] & [Form1].[f1_18]" & " In (" & Chr(34) & "02" & Chr(34) & "," & Chr(34) & "12" & Chr(34) & "," & Chr(34) & "10" & Chr(34) & "," & Chr(34) & "11" & Chr(34) & ") Or [Form4].[f4_81] & [Form4].[f4_8]" & " In (" & Chr(34) & "02" & Chr(34) & "," & Chr(34) & "12" & Chr(34) & "," & Chr(34) & "10" & Chr(34) & "," & Chr(34) & "11" & Chr(34) & ")" & " Or [Form2].[f2_261] & [Form2].[f2_26]" & " In (" & Chr(34) & "02" & Chr(34) & "," & Chr(34) & "12" & Chr(34) & "," & Chr(34) & "10" & Chr(34) & "," & Chr(34) & "11" & Chr(34) & "), [Form1].[f1_181]" & Chr(38) & "[Form1].[f1_18] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [Form2].[f2_261] " & Chr(38) & " [Form2].[f2_26] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & "[Form4].[f4_81]" & Chr(38) & "[Form4].[f4_8], " & Chr(34) & Chr(34) & ") as ECO_KOR, " & _
                 "IIf([Form1].[f1_28]" & Chr(38) & "[Form1].[f1_281]=" & Chr(34) & "211" & Chr(34) & "Or [Form1].[f1_28] &Chr(38)& [Form1].[f1_281]=" & Chr(34) & "111" & Chr(34) & ",1," & Chr(34) & Chr(34) & ") AS �_��������������, " & _
                 "IIf([f1_27_1]=131 Or [f1_27_1]=77 Or [f1_27_1]=130 Or [f1_27_2]=131 Or [f1_27_2]=130 Or [f1_27_2]=77 Or [f1_27_3]=131 Or [f1_27_3]=130 Or [f1_27_3]=77 Or [f1_27_4]=131 Or [f1_27_4]=130 Or [f1_27_4]=77,[f1_27_1] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [f1_27_2] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & " [f1_27_3] " & Chr(38) & Chr(34) & "/" & Chr(34) & Chr(38) & "[f1_27_4]," & Chr(34) & Chr(34) & ") AS f11_27, Form11.f11_25k, Form11.f11_25d" & _
                 " FROM (((((Form1 LEFT JOIN Form11 ON (Form1.[f1_4num] = Form11.[f1_4num]) AND (Form1.[f1_3num] = Form11.[f1_3num])) LEFT JOIN Form2 ON Form1.[f1_3num] = Form2.[f2_3num]) LEFT JOIN Form3 ON (Form1.[f1_4num] = Form3.[f1_4num]) AND (Form1.[f1_3num] = Form3.[f1_3num])) LEFT JOIN Form4 ON (Form1.[f1_4num] = Form4.[f1_4num]) AND (Form1.[f1_3num] = Form4.[f1_3num])) LEFT JOIN Form5 ON Form1.[f1_3num] = Form5.[f5_3num]) WHERE Form11.f11_25k In (1,41,61)"
    End Select
        Set qdf = CurrentDb.CreateQueryDef(strReportName, strSQL) '��������� ������ � ���� ������
        If FileExists(strfPath) = True Then '���� ������ ����� �� ���������� �������� ���
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strReportName, strfPath
        Else
            Kill fPath
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strReportName, strfPath
        End If
        Set qdf = Nothing
        DoCmd.DeleteObject acQuery, strReportName '������� ������, ������� �� ���������
        Select Case intNums '�������� ����������
        Case 1
            Call Eco_Excel_Interface(strfPath)
            BookXL.Visible = True
        Case 2
            Call Sud_Excel_Interface(strfPath)
            BookXL.Visible = True
        End Select
    
Exit Function
Error_Handler:
msg = "������ � [" & Application.VBE.ActiveCodePane.CodeModule & "/CreationOfSQL]" & " ������ #" & Str(Err.Number) & " � ������� [" & Err.Source & "] ��������: " & Err.Description & Chr(13) & "� ������ ������������� ������ ���������a ���������� � ������������"
MsgBox msg, vbInformation, "��!!!"
End Function
'*******************************************************************************************************************
'*********************************������� ��������� � EXCEL************************************************
'*******************************************************************************************************************
Function Eco_Excel_Interface(ByVal filePath As String)
Dim msg As String
On Error GoTo Error_Handler
    Set BookXL = CreateObject("Excel.Application")  '�������� ���������� ��������� �������� Excel
    Const xlSolid = 1: Const xlCenter = -4108: Const xlEdgeTop = 8: Const xlDown = -4121: Const xlContinuous = 1: Const xlThin = 2: Const xlUp = -4162
    Const xlInsideVertical = 11: Const xlInsideHorizontal = 12: Const xlAutomatic = -4105: Const xlContext = -5002: Const xlMedium = -4138: Const xlFormatFromLeftOrAbove = 0
    Const xlThemeColorDark1 = 1: Const xlThemeFontNone = 0: Const xlThemeColorLight1 = 2: Const xlUnderlineStyleNone = -4142: Const xlEdgeLeft = 7
    Const xlEdgeBottom = 9: Const xlPart = 2: Const xlEdgeRight = 10: Const xlByRows = 1: Const xlGeneral = 1: Dim LastRow As Long
    BookXL.Workbooks.Open filename:=filePath '��������� ����������� �������
    '*********************************����� �������********************************
    BookXL.Range("A1").Value = "���"
    BookXL.Range("B1").Value = "����� �����."
    BookXL.Range("C1").Value = "��."
    BookXL.Range("E1").Value = "C�����"
    BookXL.Range("D1").Value = "�����. ��������: � ����. / ��� ����."
    BookXL.Range("F1").Value = "�������"
    BookXL.Range("G1").Value = "���� �������"
    BookXL.Range("H1").Value = "���� ��"
    BookXL.Range("I1").Value = "���� �����������"
    BookXL.Range("J1").Value = "����. �����. ����� 1 (f1_18)"
    BookXL.Range("K1").Value = "����. �����.����� 2 (f2_26)"
    BookXL.Range("L1").Value = "����.�����.����� 4 (f4_8)"
    BookXL.Range("M1").Value = "���.-����. ����� (f1_20)"
    BookXL.Range("N1").Value = "���.-����. ����� (f2_29)"
    BookXL.Range("O1").Value = "��� ������. ����-�� (f1_22)"
    BookXL.Range("P1").Value = "��� ������. ����-�� (f2_30)"
    BookXL.Range("Q1").Value = "����� ������������� (f1_24)"
    BookXL.Range("R1").Value = "����� ������������� (f2_32)"
    BookXL.Range("S1").Value = "����� ������������� (f5_17)"
    BookXL.Range("T1").Value = "���������� ����� ���� (f2_4num)"
    BookXL.Range("U1").Value = "���. ���. �����. (f2_34_1)"
    BookXL.Range("V1").Value = "���. ���. �����. (f2_34_2)"
    BookXL.Range("W1").Value = "���. ���. �����. (f2_34_3)"
    BookXL.Range("X1").Value = "���. ���. �����. (f2_34_4)"
    BookXL.Range("Y1").Value = "����./�����."
    BookXL.Range("Z1").Value = "� ����� �"
    BookXL.Range("AA1").Value = "���. ����� (f4_10_1)"
    BookXL.Range("AB1").Value = "���. ����� (f4_10_2)"
    BookXL.Range("AC1").Value = "���. ����� (f4_10_3)"
    BookXL.Range("AD1").Value = "���. ����� (f4_10_4)"
    BookXL.Range("AE1").Value = "����������� �������� (f4_11_1)"
    BookXL.Range("AF1").Value = "����������� �������� (f4_11_2)"
    BookXL.Range("AG1").Value = "����������� �������� (f4_11_3)"
    BookXL.Range("AH1").Value = "����� �� �����(���) (f4_12)"
    BookXL.Range("AI1").Value = "����� �� �����($) (f4_121)"
    BookXL.Range("AJ1").Value = "����� ������ ��������� (f4_15)"
    BookXL.Range("AK1").Value = "������ ��������� (f4_16_1)"
    BookXL.Range("AL1").Value = "������ ��������� (f4_16_2)"
    BookXL.Range("AM1").Value = "������ ��������� (f4_16_3)"
    BookXL.Range("AN1").Value = "������ ��������� (f4_16_4)"
    BookXL.Range("AO1").Value = "����������� ��� (f4_32)"
    BookXL.Range("AP1").Value = "������"
    '*********************************�������********************************
    BookXL.Rows("1:1").RowHeight = 84.75
    BookXL.Columns("A:A").ColumnWidth = 4
    BookXL.Columns("Y:Y").ColumnWidth = 17.71
    BookXL.Columns("B:B").ColumnWidth = 17.71
    BookXL.Columns("C:C").ColumnWidth = 3
    BookXL.Columns("D:D").ColumnWidth = 9.71
    BookXL.Columns("Q:Q").ColumnWidth = 9
    '***************************�������� �������*****************************
    BookXL.Rows("2:2").Select
    BookXL.ActiveWindow.FreezePanes = True
    BookXL.Range("A1:AP1").Select
    With BookXL.Selection
        .AutoFilter: .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Orientation = 0: .AddIndent = False: .IndentLevel = 0: .ShrinkToFit = False: .ReadingOrder = xlContext: .MergeCells = False
    End With
    With BookXL.Selection.Interior
        .Pattern = xlSolid: .PatternColorIndex = xlAutomatic: .ThemeColor = xlThemeColorDark1: .TintAndShade = -0.249977111117893: .PatternTintAndShade = 0
    End With
    BookXL.Columns("AP:AP").Select
    With BookXL.Selection
        .ColumnWidth = 80: .WrapText = True: .Orientation = 0: .AddIndent = False: .IndentLevel = 0: .ShrinkToFit = False: .ReadingOrder = xlContext: .MergeCells = False
    End With
    '*********************************������ �����********************************
    BookXL.Columns("Q:Q").Select
    BookXL.Selection.Replace What:="00 / 00 / 00", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("R:R").Select
    BookXL.Selection.Replace What:="//", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Selection.Replace What:="0/0/0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Range("AA:AG").Select
    BookXL.Selection.Replace What:="0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Range("AK:AN").Select
    BookXL.Selection.Replace What:="0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    '*********************************�������*************************************
    LastRow = BookXL.Range("A" & BookXL.Rows.Count).End(xlUp).Row
    BookXL.Range("A1:AP" & LastRow).Select
    With BookXL.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
    End With
    With BookXL.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
    End With
    BookXL.Range("A1").Select
Exit Function
Error_Handler:
msg = "������ � [" & Application.VBE.ActiveCodePane.CodeModule & "/Eco_Excel_Interface]" & " ������ #" & Str(Err.Number) & " � ������� [" & Err.Source & "] ��������: " & Err.Description & Chr(13) & "� ������ ������������� ������ ���������a ���������� � ������������"
MsgBox msg, vbInformation, "��!!!"
End Function
'*******************************************************************************************************************
'*********************************������� ������������ � ���  EXCEL***************************************
'*******************************************************************************************************************
Function Sud_Excel_Interface(ByVal filePath As String)
Dim msg As String
On Error GoTo Error_Handler
    Const xlSolid = 1: Const xlCenter = -4108: Const xlEdgeTop = 8: Const xlDown = -4121: Const xlContinuous = 1: Const xlThin = 2: Const xlUp = -4162
    Const xlInsideVertical = 11: Const xlInsideHorizontal = 12: Const xlAutomatic = -4105: Const xlContext = -5002: Const xlMedium = -4138: Const xlFormatFromLeftOrAbove = 0
    Const xlThemeColorDark1 = 1: Const xlThemeFontNone = 0: Const xlThemeColorLight1 = 2: Const xlUnderlineStyleNone = -4142: Const xlEdgeLeft = 7
    Const xlEdgeBottom = 9: Const xlPart = 2: Const xlEdgeRight = 10: Const xlByRows = 1: Const xlGeneral = 1: Const xlSortOnValues = 0: Const xlAscending = 1: Const xlSortNormal = 0: Const xlSortAsNumbers = 1
    Const xlTopToBottom = 1: Const xlPinYin = 1: Const xlYes = 1
    Dim LastRow As Long
    Set BookXL = CreateObject("Excel.Application") '�������� ���������� ��������� �������� Excel
    BookXL.Workbooks.Open filename:=filePath
    '*********************************����� �������********************************
    BookXL.Range("A1").Value = "���"
    BookXL.Range("B1").Value = "����� �����."
    BookXL.Range("C1").Value = "�./�."
    BookXL.Range("D1").Value = "����� �����."
    BookXL.Range("I1").Value = "���� = 002" & vbNewLine & "����������� = 100"
    BookXL.Range("J1").Value = "����� �����. �1.1"
    BookXL.Range("K1").Value = "����� �����. �2"
    BookXL.Range("L1").Value = "����� ������� �1.1"
    BookXL.Range("M1").Value = "����� ������� �2"
    BookXL.Range("N1").Value = "� ��������� �1.1/�2"
    BookXL.Range("O1").Value = "����� �1/�1.1/�5"
    BookXL.Range("P1").Value = "� ������� 3 = ������ ���, 4 = ������ �� ����. 2 = ��� 1 = ��"
    BookXL.Range("Q1").Value = "���. ����� �.5"
    BookXL.Range("R1").Value = "���. ����� �.11"
    BookXL.Range("S1").Value = "���. ����� �.4"
    BookXL.Range("T1").Value = "���. �������� �.4"
    BookXL.Range("U1").Value = "����� ������ �.4"
    BookXL.Range("V1").Value = "������ ��������� �.4 "
    BookXL.Range("W1").Value = "�����. ��� �.4"
    BookXL.Range("X1").Value = "������ ���/������/�� �.4"
    BookXL.Range("Y1").Value = "���. = 12;02 " & vbNewLine & "����. = 12;11"
    BookXL.Range("Z1").Value = "� ���." & vbNewLine & "�1_28 = (211/111)"
    BookXL.Range("AA1").Value = "���. ���-��" & vbNewLine & "(��� = 131, ��� = 77, ��� = 130)"
    BookXL.Range("AB1").Value = "�������"
    BookXL.Range("AC1").Value = "���� �������"
    '*********************************�������********************************
    BookXL.Rows("1:1").RowHeight = 130
    BookXL.Columns("A:A").ColumnWidth = 4
    BookXL.Columns("B:B").ColumnWidth = 17.71
    BookXL.Columns("C:C").ColumnWidth = 3
    BookXL.Columns("D:D").ColumnWidth = 17.71
    BookXL.Columns("E:E").ColumnWidth = 17
    BookXL.Columns("F:F").ColumnWidth = 10
    BookXL.Columns("G:G").ColumnWidth = 11
    BookXL.Columns("I:I").ColumnWidth = 10
    BookXL.Columns("X:X").ColumnWidth = 10
    BookXL.Columns("AA:AA").ColumnWidth = 12
    BookXL.Columns("N:P").ColumnWidth = 5.71
    BookXL.Columns("J:M").ColumnWidth = 2
    BookXL.Columns("R:T").ColumnWidth = 11
    '***************************�������� �������*****************************
    BookXL.Rows("2:2").Select
    BookXL.ActiveWindow.FreezePanes = True
    BookXL.Range("A1:AC1").Select
    With BookXL.Selection
        .AutoFilter: .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Orientation = 0: .AddIndent = False: .IndentLevel = 0: .ShrinkToFit = False: .ReadingOrder = xlContext: .MergeCells = False
    End With
    With BookXL.Selection.Interior
        .Pattern = xlSolid: .PatternColorIndex = xlAutomatic: .ThemeColor = xlThemeColorDark1: .TintAndShade = -0.249977111117893: .PatternTintAndShade = 0
    End With
    BookXL.Range("G1:P1").Orientation = 90
    BookXL.Range("X1").Orientation = 90
    '*********************************������ �����********************************
    BookXL.Columns("E:E").Select
    BookXL.Selection.Replace What:=" . .", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("G:G").Select
    BookXL.Selection.Replace What:="00/", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("I:I").Select
    BookXL.Selection.Replace What:="00_00/", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("N:N").Select
    BookXL.Selection.Replace What:="00/", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("R:R").Select
    BookXL.Selection.Replace What:="0_0/0_0/0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("R:R").Select
    BookXL.Selection.Replace What:="/0_0/0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("R:R").Select
    BookXL.Selection.Replace What:="/0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("S:S").Select
    BookXL.Selection.Replace What:="0_0/0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("S:S").Select
    BookXL.Selection.Replace What:="/0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("S:S").Select
    BookXL.Selection.Replace What:="/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("T:T").Select
    BookXL.Selection.Replace What:="0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("T:T").Select
    BookXL.Selection.Replace What:="/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("T:T").Select
    BookXL.Selection.Replace What:="/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("V:V").Select
    BookXL.Selection.Replace What:="0_0/0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("V:V").Select
    BookXL.Selection.Replace What:="0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("V:V").Select
    BookXL.Selection.Replace What:="/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("V:V").Select
    BookXL.Selection.Replace What:="/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("X:X").Select
    BookXL.Selection.Replace What:="0_0/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("X:X").Select
    BookXL.Selection.Replace What:="/0_0/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    BookXL.Columns("X:X").Select
    BookXL.Selection.Replace What:="/0_0", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    LastRow = BookXL.Range("A" & BookXL.Rows.Count).End(xlUp).Row
    '*********************************����������********************************
    BookXL.Worksheets("obvinQry").Sort.SortFields.Clear
    BookXL.Worksheets("obvinQry").Sort.SortFields.Add Key:=BookXL.Range("AB2:AB" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    BookXL.Worksheets("obvinQry").Sort.SortFields.Add Key:=BookXL.Range("B2:B" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With BookXL.Worksheets("obvinQry").Sort
        .SetRange BookXL.Range("A1:AC" & LastRow): .Header = xlYes: .MatchCase = False: .Orientation = xlTopToBottom: .SortMethod = xlPinYin: .Apply
    End With
    '*********************************�������*************************************
    BookXL.Range("A1:AC" & LastRow).Select
    With BookXL.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
    End With
    With BookXL.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
    End With
    BookXL.Range("A1").Select
       
Exit Function
Error_Handler:
msg = "������ � [" & Application.VBE.ActiveCodePane.CodeModule & "/Sud_Excel_Interface]" & " ������ #" & Str(Err.Number) & " � ������� [" & Err.Source & "] ��������: " & Err.Description & Chr(13) & "� ������ ������������� ������ ���������a ���������� � ������������"
MsgBox msg, vbInformation, "��!!!"
End Function


