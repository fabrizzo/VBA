Attribute VB_Name = "NewMacros3"

Function �������_�_��������(q) As String

'
 

'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
 '   Windows(1).Activate
' MsgBox "�����.1.."


' ������2 ������
'����� �����

 Selection.HomeKey wdStory                   ' ���������� ������ � ������ ���������
    With Selection.Find
        .Text = "//*=Q"    ' ������� �����
        .MatchWildcards = True              ' �������������� ����� � ���� ������ (* - ����� �����)
        .Wrap = wdFindStop                  ' ������������ �� ���������
        Do
            .Execute
     
        Loop Until Not .Found               ' ������ ���� ����� (.Find) �� �������� �����
    End With
    '������ �� 1-� ��������� �������� a, �� 2-� (���� ����) - b; � ����� (��. �� Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate


' MsgBox "�����.1.."




'  ����� ������� ��������� � ���������
Dim p As String

p = ":" & q & "*-----------------------"

 ' MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                       ' MsgBox "���� �����������", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select
                 'MsgBox "�����...", vbInformation
                
                ' �����
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)


'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "����� �� ������", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function

Function �������_���_�����(q) As String
'
'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
'    Windows(1).Activate
' MsgBox "�����.1.."


' ������2 ������
'����� �����

 Selection.HomeKey wdStory                   ' ���������� ������ � ������ ���������
    With Selection.Find
        .Text = "//*==Q"    ' ������� �����
        .MatchWildcards = True              ' �������������� ����� � ���� ������ (* - ����� �����)
        .Wrap = wdFindStop                  ' ������������ �� ���������
        Do
            .Execute
     
        Loop Until Not .Found               ' ������ ���� ����� (.Find) �� �������� �����
    End With
    '������ �� 1-� ��������� �������� a, �� 2-� (���� ����) - b; � ����� (��. �� Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate
' MsgBox "�����.1.."



'  ����� ������� ��������� � ���������
Dim p As String

p = ":" & q & " "

' MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                        'MsgBox "���� �����������", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 5, r.End + 1).Select
                Selection.MoveEndUntil Cset:=Chr(13)

               ' MsgBox "�����...", vbInformation
                
                ' �����
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)
Selection.TypeParagraph

'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "����� �� ������", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function


Function ����_�������_�_��������(q) As String

'
'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
 '   Windows(1).Activate
' MsgBox "�����.1.."


' ������2 ������
'����� �����

 Selection.HomeKey wdStory                   ' ���������� ������ � ������ ���������
    With Selection.Find
        .Text = "//*==Q"    ' ������� �����
        .MatchWildcards = True              ' �������������� ����� � ���� ������ (* - ����� �����)
        .Wrap = wdFindStop                  ' ������������ �� ���������
        Do
            .Execute
     
        Loop Until Not .Found               ' ������ ���� ����� (.Find) �� �������� �����
    End With
    '������ �� 1-� ��������� �������� a, �� 2-� (���� ����) - b; � ����� (��. �� Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate
' MsgBox "�����.1.."



'  ����� ������� ��������� � ���������
Dim p As String

p = " " & q & " 2*-----------------------"

 'MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                        'MsgBox "���� �����������", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 22, r.End + 1).Select
         
                ' MsgBox "�����...", vbInformation
                
                ' �����
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)


'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "����� �� ������", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function


Function ����_�������_���_�����(q) As String
'
'Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
'    Windows(1).Activate
' MsgBox "�����.1.."


' ������2 ������
'����� �����

 Selection.HomeKey wdStory                   ' ���������� ������ � ������ ���������
    With Selection.Find
        .Text = "//*==Q"    ' ������� �����
        .MatchWildcards = True              ' �������������� ����� � ���� ������ (* - ����� �����)
        .Wrap = wdFindStop                  ' ������������ �� ���������
        Do
            .Execute
     
        Loop Until Not .Found               ' ������ ���� ����� (.Find) �� �������� �����
    End With
    '������ �� 1-� ��������� �������� a, �� 2-� (���� ����) - b; � ����� (��. �� Ctrl-Shift-F5)
On Error GoTo ChangeWinErr

Selection.Copy



Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)

 Selection.TypeParagraph
ChangeWinErr:

Windows(1).Activate
' MsgBox "�����.1.."



'  ����� ������� ��������� � ���������
Dim p As String

p = " " & q & " 2"


' MsgBox (p)


    Dim r, f As Boolean, firstOccurence As Long
    Set r = ActiveDocument.Range
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
                If f Then
                    If r.Start = firstOccurence Then
                        'MsgBox "���� �����������", vbInformation
                        Exit Do
                    End If
                Else
                    firstOccurence = r.Start
                    f = True
                End If
                ActiveDocument.Range(r.Start - 22, r.End + 1).Select
                Selection.MoveEndUntil Cset:=Chr(13)

               ' MsgBox "�����...", vbInformation
                
                ' �����
                '
                '
Selection.Copy
Windows(1).Activate
Windows(2).Activate
Selection.PasteAndFormat (wdPasteDefault)
Selection.TypeParagraph

'On Error GoTo ChangeWinErr
'Set bb = ActiveWindow.Next
'If Windows.Count > 0 Then
'bb.Activate
'Exit Sub
'End If
'ChangeWinErr:
Windows(1).Activate
                 
                 
                 

'
                Set r = ActiveDocument.Range(r.End, r.End)
            Else
                'MsgBox "����� �� ������", vbExclamation
                Exit Do
            End If
        End With
    Loop

End Function
Sub �������_��_�������()
Attribute �������_��_�������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������2"
    
 Dim kod, a As String
' ������� ����� ��������
Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
Windows(1).Activate
      
kod = InputBox("������� ��� ���")
        
 ' 09 // ���������� � ������������ �� ���� �� ���������� ��������� // (����� 1 �������� 9 ��� 125)
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="09.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

' 10 // ��� ������ ������������� �� ������ � ����������� ��.����
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="10.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

'10_1 // ��� ������ ������������� �� ������ � ����������� ��.���� �� ���������� ��������� // �������� 103 ��� 31
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="10_1.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'10_2 // ��� ������ ������������� �� ������ � ����������� ��.���� �� ���������� �������� // �������� 103 ��� 32
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="10_2.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'10_3 // ��� ������ ������������� �� ������ � ����������� ��.���� �� ���������� �������� // �������� 103 ��� 33
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="10_3.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

'11 // � ���������� ������������� ��� ���������� �� ����
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="11.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

'12 // ���������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="12.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'13 // ���������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="13.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'14_1_14 // �������������� ����
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="14_1_14.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'15 // ���
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="15.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


'16 // ���
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="16.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'17 // ���.������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="17.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'18 //  �������� ��������� �������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="18.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'19 //  � ����� �����������������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="19.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'20 //  � ��������������, � ��� ����� � ����������� ������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="20.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
'21 // ��������� � ������������ ������ (��� ����)
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="21.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

'22 // ��������� �� � ������������ ������ � ������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="22.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

'23 // ��������� �� ������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="23.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

'24 // ��������� �� �� ������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="24.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

'25 //  ��������� � ������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="25.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close


'26 // ��������� � ������ �� �������

 ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="26.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

'27 // ����������� � ������� ��� � ���
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="27.lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

'28 // ���� , ����������� ������������, �� ������� �������� � ����������� �� �.3 �.1 ��.24 � �.4 �.1 ��.24 ��� ��
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="28.lst", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close


'29 // ����������� �������������������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_�_��������(kod)
ActiveDocument.Close


 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close


ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="29.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close
 
' 30 // � ����������� ���������
    ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="30.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
 ' 30_1 // � ������������� ���������
     ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="30_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
 ' 30_2 // � ����������� ���������
      ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="30_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
 ' 31 // ����������� ���������
      ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="31.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
 ' 31_1 // ����������� ������������
      ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="31_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
  
  ' 31_2 // �� ����������� ������������� ������� � ������������ �������
       ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="31_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
   
 ' 32 // �� � ����������� ���������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="32.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
 
 ' 33_2 // �� ������������ ������������� �������  � ������������ �������

       ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="33_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
   
'34 // ����� ����������� ������������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="34.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'35 // ����� ������� ����
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="35.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'36 //  �����������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="36.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

'38 //  ��� ����������� ��������� ������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="38.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
 '39 // ��������� ������������ ����������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="39.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
'39_1 // ��������� ������ ��� �����������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="39_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
'39_2 // ��������� ����������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="39_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
'40 // ��������� � ��������� ����������� �������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="40.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
 
'40_1 // ��������� � ��������� ���������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="40_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

'43 //   ������������� ��������������  �����������, ����������� �����
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="43.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'43_1 // ������������ ������������� ��������������, �� ������� ���������� �������� �-4 � ��������� ����� ������
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="43_1.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'43_2 //   ���������  �����������, ����������� �����
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="43_2.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'43_3 //  ������������ �����������, ����������� �����
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="43_3.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'43_4 //  �� ������  �����������, ����������� �����
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="43_4.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close
'43_5 //  ���������������  �����������, ����������� �����
         ChangeFileOpenDirectory "C:\���������\CK\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
ActiveDocument.Close

 ChangeFileOpenDirectory "C:\���������\SLEDSTV\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\DOZNAN\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = ����_�������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:="43_5.lst", ConfirmConversions:=False, ReadOnly:= _
        False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
        "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
        Format:=wdOpenFormatAuto, XMLTransform:="", Encoding:=1251
a = �������_���_�����(kod)
 ActiveDocument.Close



' �����
 Application.Run MacroName:="���������_��������_�����"
MsgBox "���� �����������", vbInformation

End Sub
Sub ���������_��������_�����()
Attribute ���������_��������_�����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������4"
'
' ������4 ������
'
'
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



Sub �������_��_�������1()

    
 Dim kod, a As String
 
 Dim Fname(70, 2) As String
 Dim kol As Integer
 Dim b As Integer
 
' ������� ����� ��������
Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
Windows(1).Activate
      
kod = InputBox("������� ��� ���")


kol = 0



' 09 // ���������� � ������������ �� ���� �� ���������� ��������� // (����� 1 �������� 9 ��� 125)
kol = kol + 1
Fname(kol, 1) = "09"
Fname(kol, 2) = "������"

' 10 // ��� ������ ������������� �� ������ � ����������� ��.����
kol = kol + 1
Fname(kol, 1) = "10"
Fname(kol, 2) = ""

'10_1 // ��� ������ ������������� �� ������ � ����������� ��.���� �� ���������� ��������� // �������� 103 ��� 31
kol = kol + 1
Fname(kol, 1) = "10_1"
Fname(kol, 2) = ""

'10_2 // ��� ������ ������������� �� ������ � ����������� ��.���� �� ���������� �������� // �������� 103 ��� 32
kol = kol + 1
Fname(kol, 1) = "10_3"
Fname(kol, 2) = ""

'10_3 // ��� ������ ������������� �� ������ � ����������� ��.���� �� ���������� �������� // �������� 103 ��� 33
kol = kol + 1
Fname(kol, 1) = "10_3"
Fname(kol, 2) = ""

'11 // � ���������� ������������� ��� ���������� �� ����
kol = kol + 1
Fname(kol, 1) = "11"
Fname(kol, 2) = "������"

'12 // ���������
kol = kol + 1
Fname(kol, 1) = "12"
Fname(kol, 2) = "������"
 
'13 // ���������
kol = kol + 1
Fname(kol, 1) = "13"
Fname(kol, 2) = "������"
 
'14_1_14 // �������������� ����
kol = kol + 1
Fname(kol, 1) = "14_1_14"
Fname(kol, 2) = "������"
 
'15 // ���
kol = kol + 1
Fname(kol, 1) = "15"
Fname(kol, 2) = "������"

'16 // ���
kol = kol + 1
Fname(kol, 1) = "16"
Fname(kol, 2) = "������"
 
'17 // ���.������
kol = kol + 1
Fname(kol, 1) = "17"
Fname(kol, 2) = "������"
 
'18 //  �������� ��������� �������
kol = kol + 1
Fname(kol, 1) = "18"
Fname(kol, 2) = "������"
 
'19 //  � ����� �����������������
kol = kol + 1
Fname(kol, 1) = "19"
Fname(kol, 2) = "������"
 
'20 //  � ��������������, � ��� ����� � ����������� ������
 kol = kol + 1
Fname(kol, 1) = "20"
Fname(kol, 2) = "������"
 
'21 // ��������� � ������������ ������ (��� ����)
kol = kol + 1
Fname(kol, 1) = "21"
Fname(kol, 2) = "������"

'22 // ��������� �� � ������������ ������ � ������
kol = kol + 1
Fname(kol, 1) = "22"
Fname(kol, 2) = "������"

'23 // ��������� �� ������
kol = kol + 1
Fname(kol, 1) = "23"
Fname(kol, 2) = "������"

'24 // ��������� �� �� ������
kol = kol + 1
Fname(kol, 1) = "24"
Fname(kol, 2) = "������"

'40 // ��������� � ��������� ����������� �������
kol = kol + 1
Fname(kol, 1) = "40"
Fname(kol, 2) = ""
 
'40_1 // ��������� � ��������� ���������
kol = kol + 1
Fname(kol, 1) = "40_1"
Fname(kol, 2) = ""

'43 //   ������������� ��������������  �����������, ����������� �����
kol = kol + 1
Fname(kol, 1) = "43"
Fname(kol, 2) = ""

'43_1 // ������������ ������������� ��������������, �� ������� ���������� �������� �-4 � ��������� ����� ������
kol = kol + 1
Fname(kol, 1) = "43_1"
Fname(kol, 2) = ""

'43_2 //   ���������  �����������, ����������� �����
kol = kol + 1
Fname(kol, 1) = "43_2"
Fname(kol, 2) = ""

'43_3 //  ������������ �����������, ����������� �����
kol = kol + 1
Fname(kol, 1) = "43_3"
Fname(kol, 2) = ""

'43_4 //  �� ������  �����������, ����������� �����
kol = kol + 1
Fname(kol, 1) = "43_4"
Fname(kol, 2) = ""

'43_5 //  ���������������  �����������, ����������� �����
kol = kol + 1
Fname(kol, 1) = "43_5"
Fname(kol, 2) = ""



For b = 1 To kol Step 1
        
        
If Fname(b, 2) = "������" Then

        
        
 ' C ��������
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

'ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
'    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
'a = ����_�������_�_��������(kod)
' ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\FSSP\MEC\"
    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close

ChangeFileOpenDirectory "C:\���������\GPN\MEC\"
    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
a = �������_�_��������(kod)
 ActiveDocument.Close




Else

' ��� �����
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

'ChangeFileOpenDirectory "C:\���������\FSKN\MEC\"
'    Documents.Open FileName:=Fname(b, 1) & ".lst", Encoding:=1251
'a = ����_�������_���_�����(kod)
' ActiveDocument.Close

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





' �����
 Application.Run MacroName:="���������_��������_�����"
MsgBox "���� �����������", vbInformation


End Sub

