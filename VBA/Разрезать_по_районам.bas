Attribute VB_Name = "���������_��_�������"
Sub �������_��_�������1()
    Dim kod, a As String: Dim Fname(70, 2) As String
    Dim kol As Integer: Dim b As Integer: Dim j As Integer
    Dim i As Integer
    Dim region() As String
    Dim maxnum As Integer
    Dim num As Variant
    Dim bf As Document
    For Each bf In Documents
    bf.Close SaveChanges:=False
    Next bf
    num = InputBox("������� ������������� ����� ������ ? (�������)")
    Select Case StrPtr(num)
    Case 0
        Exit Sub
    Case Else
        maxnum = CInt(num)
    End Select
    ReDim region(1 To maxnum)
    For i = 1 To maxnum
    region(i) = InputBox("������� �������������: " & i)
    Next i
    For j = 1 To maxnum
    kod = region(j)
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    Windows(1).Activate
    kol = 0
    '12 // ���������
    kol = kol + 1: Fname(kol, 1) = "12": Fname(kol, 2) = "������"
    '13 // ���������
    kol = kol + 1: Fname(kol, 1) = "13": Fname(kol, 2) = "������"
    '14_1_14 // �������������� ����
    kol = kol + 1: Fname(kol, 1) = "14_1_14": Fname(kol, 2) = "������"
    '15 // ���
    kol = kol + 1: Fname(kol, 1) = "15": Fname(kol, 2) = "������"
    '16 // ���
    kol = kol + 1: Fname(kol, 1) = "16": Fname(kol, 2) = "������"
    '17 // ���.������
    kol = kol + 1: Fname(kol, 1) = "17": Fname(kol, 2) = "������"
    '18 //  �������� ��������� �������
    kol = kol + 1: Fname(kol, 1) = "18": Fname(kol, 2) = "������"
    '19 //  � ����� �����������������
    kol = kol + 1: Fname(kol, 1) = "19": Fname(kol, 2) = "������"
    '20 //  � ��������������, � ��� ����� � ����������� ������
    kol = kol + 1: Fname(kol, 1) = "20": Fname(kol, 2) = "������"
    '21 // ��������� � ������������ ������ (��� ����)
    kol = kol + 1: Fname(kol, 1) = "21": Fname(kol, 2) = "������"
    '22 // ��������� �� � ������������ ������ � ������
    kol = kol + 1: Fname(kol, 1) = "22": Fname(kol, 2) = "������"
    '23 // ��������� �� ������
    kol = kol + 1: Fname(kol, 1) = "23": Fname(kol, 2) = "������"
    '25 //  ��������� � ������
    kol = kol + 1: Fname(kol, 1) = "25": Fname(kol, 2) = ""
    '26 // ��������� � ������ �� �������
    kol = kol + 1: Fname(kol, 1) = "26": Fname(kol, 2) = ""
    '27 // ����������� � ������� ��� � ���
    kol = kol + 1: Fname(kol, 1) = "27": Fname(kol, 2) = ""
    '28 // ���� , ����������� ������������, �� ������� �������� � ����������� �� �.3 �.1 ��.24 � �.4 �.1 ��.24 ��� ��
    kol = kol + 1: Fname(kol, 1) = "28": Fname(kol, 2) = ""
    '29 // ����������� �������������������
    kol = kol + 1: Fname(kol, 1) = "29": Fname(kol, 2) = "������"
    '29 // ����������� �������������������
    kol = kol + 1: Fname(kol, 1) = "29_1": Fname(kol, 2) = ""
    ' 30 // � ����������� ���������
    kol = kol + 1: Fname(kol, 1) = "30": Fname(kol, 2) = ""
    ' 30_1 // � ������������� ���������
    kol = kol + 1: Fname(kol, 1) = "30_1": Fname(kol, 2) = ""
    ' 30_2 // � ����������� ���������
    kol = kol + 1: Fname(kol, 1) = "30_2": Fname(kol, 2) = ""
    ' 30_3 // ���������
    kol = kol + 1: Fname(kol, 1) = "30_3": Fname(kol, 2) = ""
    ' 30_4 // ������������
    kol = kol + 1: Fname(kol, 1) = "30_4": Fname(kol, 2) = ""
    ' 30_5 // ����������� ������������� ������� � ������������ �������
    kol = kol + 1: Fname(kol, 1) = "30_5": Fname(kol, 2) = ""
    ' 31 // ����������� ���������
    kol = kol + 1: Fname(kol, 1) = "31": Fname(kol, 2) = ""
    ' 31_1 // ����������� ������������
    kol = kol + 1: Fname(kol, 1) = "31_1": Fname(kol, 2) = ""
    ' 31_2 // �� ����������� ������������� ������� � ������������ �������
    kol = kol + 1: Fname(kol, 1) = "31_2": Fname(kol, 2) = ""
    ' 31_3 // � ����������� ���������
    kol = kol + 1: Fname(kol, 1) = "31_3": Fname(kol, 2) = ""
    ' 31_4 // � ����������� ���������
    kol = kol + 1: Fname(kol, 1) = "31_4": Fname(kol, 2) = ""
    ' 32 // �� � ����������� ���������
    kol = kol + 1: Fname(kol, 1) = "32": Fname(kol, 2) = ""
    ' 33_2 // �� ������������ ������������� �������  � ������������ �������
    kol = kol + 1: Fname(kol, 1) = "33_2": Fname(kol, 2) = ""
    '34 // ����� ����������� ������������
    kol = kol + 1: Fname(kol, 1) = "34": Fname(kol, 2) = ""
    '34_1 // ����� ����������� ������������
    kol = kol + 1: Fname(kol, 1) = "34_1":  Fname(kol, 2) = ""
    '35 // ����� ������� ����
    kol = kol + 1: Fname(kol, 1) = "35":  Fname(kol, 2) = ""
    '35_1 // ����� ������� ����
    kol = kol + 1: Fname(kol, 1) = "35_1": Fname(kol, 2) = ""
    '36 //  �����������
    kol = kol + 1: Fname(kol, 1) = "36": Fname(kol, 2) = ""
    '37 //  �����������
    kol = kol + 1: Fname(kol, 1) = "37":  Fname(kol, 2) = ""
    '38 //  ��� ����������� ��������� ������
    kol = kol + 1: Fname(kol, 1) = "38":  Fname(kol, 2) = ""
    '39 // ��������� ������������ ����������
    kol = kol + 1: Fname(kol, 1) = "39":  Fname(kol, 2) = ""
    '39_1 // ��������� ������ ��� �����������
    kol = kol + 1: Fname(kol, 1) = "39_1": Fname(kol, 2) = ""
    '39_2 // ��������� ����������
    kol = kol + 1: Fname(kol, 1) = "39_2": Fname(kol, 2) = ""
    '40 // ��������� � ��������� ����������� �������
    kol = kol + 1: Fname(kol, 1) = "40":   Fname(kol, 2) = ""
    '40_1 // ��������� � ��������� ���������
    kol = kol + 1: Fname(kol, 1) = "40_1": Fname(kol, 2) = ""
    '43 //   ������������� ��������������  �����������, ����������� �����
    kol = kol + 1: Fname(kol, 1) = "43":  Fname(kol, 2) = ""
    '43_1 // ������������ ������������� ��������������, �� ������� ���������� �������� �-4 � ��������� ����� ������
    kol = kol + 1: Fname(kol, 1) = "43_1": Fname(kol, 2) = ""
    '43_2 //   ���������  �����������, ����������� �����
    kol = kol + 1: Fname(kol, 1) = "43_2": Fname(kol, 2) = ""
    '43_3 //  ������������ �����������, ����������� �����
    kol = kol + 1: Fname(kol, 1) = "43_3": Fname(kol, 2) = ""
    '43_4 //  �� ������  �����������, ����������� �����
    kol = kol + 1: Fname(kol, 1) = "43_4": Fname(kol, 2) = ""
    '43_5 //  ���������������  �����������, ����������� �����
    kol = kol + 1: Fname(kol, 1) = "43_5": Fname(kol, 2) = ""
    For b = 1 To kol Step 1
        If Fname(b, 2) = "������" Then
            ' C ��������
            Application.ScreenUpdating = False
            Application.DisplayScrollBars = False
            
            ChangeFileOpenDirectory "C:\���������\CK\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_�_��������(kod)
            ActiveDocument.Close
            
            ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_�_��������(kod)
            ActiveDocument.Close
            
            ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_�_��������(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_�_��������(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_�_��������(kod)
            ActiveDocument.Close
        Else ' ��� �����
    
            ChangeFileOpenDirectory "C:\���������\CK\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_���_�����(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_���_�����(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_���_�����(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_���_�����(kod)
            ActiveDocument.Close

            ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
            Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
            a = �������_���_�����(kod)
            ActiveDocument.Close
        End If
    Next b
    Call ���������_��������_�����
    Call MyFileSave(kod)
    ActiveDocument.Close SaveChanges:=False
Next j
 Application.ScreenUpdating = True
 Application.DisplayScrollBars = True
End Sub
Sub ���������_��������_�����()
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(3)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(21)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    Selection.WholeStory
    Selection.Font.Size = 6
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "fssp"
        .Replacement.Text = "����"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    
  Selection.HomeKey wdStory
End Sub
Sub MyFileSave(ByVal id As String)
Dim sPath As String
sPath = "C:\"
With Dialogs(wdDialogFileSaveAs)
    .name = sPath & "��� - " & id
    .Show
    End With
End Sub
Function �������_�_��������(q) As String
ChangeWinErr:
    Dim p As String
    p = ":" & q & "*-----------------------"
    Dim r, f, flag As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
    flag = False
    Do
        With r.Find
            .ClearFormatting
            .Text = p
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            If .Execute Then
                If flag = False Then
                    flag = True
                Selection.HomeKey wdStory                   ' ���������� ������ � ������ ���������
                With Selection.Find
                    .Text = "//*=Q"    ' ������� �����
                    .MatchWildcards = True              ' �������������� ����� � ���� ������ (* - ����� �����)
                    .Wrap = wdFindStop                  ' ������������ �� ���������
                    Do
                        .Execute
                    Loop Until Not .Found               ' ������ ���� ����� (.Find) �� �������� �����
                End With '������ �� 1-� ��������� �������� a, �� 2-� (���� ����) - b; � ����� (��. �� Ctrl-Shift-F5)
                On Error GoTo ChangeWinErr
                Selection.Copy
                Windows(1).Activate
                Windows(2).Activate
                Selection.PasteAndFormat (wdPasteDefault)
                Selection.TypeParagraph
                Windows(1).Activate
                End If
                If f Then
                    If r.Start = firstOccurence Then ' MsgBox "���� �����������", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select 'MsgBox "�����...", vbInformation
    Selection.Copy
    Windows(1).Activate
    Windows(2).Activate
    Selection.PasteAndFormat (wdPasteDefault)
    Windows(1).Activate
                Set r = ActiveDocument.Range(r.End, r.End)
            Else 'MsgBox "����� �� ������", vbExclamation
                Exit Do
            End If
        End With
    Loop
End Function
Function �������_���_�����(q) As String
ChangeWinErr:
    Dim p As String
    p = ":" & q & " "
    Dim r, f, flag As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
    flag = False
    Do
        With r.Find
            .ClearFormatting
            .Text = p
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            If .Execute Then
                    If flag = False Then
                    flag = True
                    Selection.HomeKey wdStory                   ' ���������� ������ � ������ ���������
                    With Selection.Find
                        .Text = "//*==Q"    ' ������� �����
                        .MatchWildcards = True              ' �������������� ����� � ���� ������ (* - ����� �����)
                        .Wrap = wdFindStop                  ' ������������ �� ���������
                        Do
                            .Execute
                     
                        Loop Until Not .Found               ' ������ ���� ����� (.Find) �� �������� �����
                    End With '������ �� 1-� ��������� �������� a, �� 2-� (���� ����) - b; � ����� (��. �� Ctrl-Shift-F5)
                    On Error GoTo ChangeWinErr
                    Selection.Copy
                    Windows(1).Activate
                    Windows(2).Activate
                    Selection.PasteAndFormat (wdPasteDefault)
                    Selection.TypeParagraph
                    Windows(1).Activate
                    End If
                If f Then
                    If r.Start = firstOccurence Then 'MsgBox "���� �����������", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select
                Selection.MoveEndUntil Cset:=Chr(13)
    Selection.Copy
    Windows(1).Activate
    Windows(2).Activate
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.TypeParagraph
    Windows(1).Activate
                Set r = ActiveDocument.Range(r.End, r.End)
            Else 'MsgBox "����� �� ������", vbExclamation
                Exit Do
            End If
        End With
    Loop
End Function
