VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM�d�l�����ڂ̒��o 
   Caption         =   "���ώd�l���̍��ڒ��o"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13140
   OleObjectBlob   =   "20161027FRM�d�l�����ڂ̒��o.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRM�d�l�����ڂ̒��o"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�G�N�Z���ɂ̂����Ƃ��ɋ�؂��Ă��Ȃ��B�X�v���b�g����Ƃ��ɐ�������؂蕶�����w��ł��Ă��Ȃ����A
'�X�v���b�g���鎞�_�Ńe�L�X�g�ɐ�������؂蕶���������Ă��Ȃ����Ƃ������Ǝv���B
Dim Delimeter As String
Dim colPOS As Long
Dim ColItem As Long
Dim colEUR As Long
Dim colRemark As Long
Dim colOption As Long
Dim MatchMode As Boolean
Dim NormalizedLists As String
Dim mEventCancel As Boolean
Const StartRow As Long = 2
Const flagOption As String = "OPTION"

Private Sub ���ώd�l���̍��ږ��𒊏o()
'TODO: �I���W�i�������̃R�s�[����B�����Œu�����i1�s�ɐ��`�j�A���o����B���̈ꎞ�V�[�g�́A�����I����ɕۑ������j������
  
    Dim rng As Word.Range    'Range�I�u�W�F�N�g
    Dim copyDoc As Document '�I���W�i�������̎g���̂ăR�s�[�i�I���W�i���̓��e�ɕύX�����������Ȃ��̂Łj
    Dim NewDoc As Document
    Dim sDocName As String
    Dim IsJA As Boolean
    Delimeter = "$"
    
    '��ʂ̍X�V���I�t
    Word.application.ScreenUpdating = False
    On Error GoTo CloseDoc
    Set NewDoc = CopyOriginalTextToTempDoc
    application.WindowState = wdWindowStateMinimize
    
'   ����̔���ɕs�v�ȕ�������򂷂�
    Call �G�f�B�^�ŕ��������(NewDoc)
    
    If chkAutoRecognition Then
        Call ���[�J�[�̎����}�b�`(NewDoc)
        Call �����̌��ꔻ��(NewDoc)
    End If
    
'  ���o�����̑S�p�����𔼊p���i�S���[�J�[���ʏ����j
    Set rng = NewDoc.Range(0, 0)
    Call ���p�S�p��(rng)

'   �\��f�e�L�X�g�ɖ߂��i�\���̃e�L�X�g�������ɂ�����Ȃ��j
    Call ConvertTableToText(NewDoc)


    '�d�l���̎�ނɍ��킹�ď�����I��
    If optBerents.Value = True Then Call ���Ԃƍ��ږ���1�s��(rng)  'Berents�p
    Call �i������肾��(rng, NewDoc)
    Call DumpTempDoc(rng, NewDoc)
    If MatchMode Then Exit Sub
    
    '��ʂ̍X�V���I��
    Word.application.ScreenUpdating = True
    MsgBox "�������܂����B", vbInformation, "���m�点"
    
    Set rng = Nothing
Exit Sub

CloseDoc:
    Call DumpTempDoc(rng, NewDoc)
    Dim msg
    msg = "�G���[�I�����܂���" & vbCrLf & Err.Number & vbTab & Err.Description
    MsgBox msg
    Set rng = Nothing
End Sub

Private Sub ConvertTableToText(doc)
    Dim tbl As Word.Table
    
    For Each tbl In doc.Tables
        tbl.ConvertToText Separator:=wdSeparateByTabs
    Next
    
End Sub
Private Sub �����̌��ꔻ��(doc)
    Dim IsJA As Boolean
    IsJA = EvalJADoc(doc)
    Select Case IsJA
        Case True
            chkJA.Value = True
        Case False
            chkEN.Value = True
    End Select
    MsgBox IsJA
    
End Sub

Private Function CopyOriginalTextToTempDoc() As Word.Document
    Dim sDocName As String

    '��Ɨp�ꎞ�h�L�������g������A�e�L�X�g�{�b�N�X�ɓ��͂��ꂽ�p�X�̕�����\��t����
    If cmbDocumentName.Text = "" Then
        MsgBox "�ǂݍ��ޕ�����I�����Ă�������", , _
                vbInformation, "���m�点"
        cmbDocumentName.SetFocus
        Exit Function
    End If
    
    sDocName = cmbDocumentName.Text
    Set CopyOriginalTextToTempDoc = Documents.Add

    With CopyOriginalTextToTempDoc.Range
        .InsertFile sDocName
        .Collapse wdCollapseEnd
    End With
    
End Function

Private Sub DumpTempDoc(rng, NewDoc)
    Set rng = Nothing
    application.DisplayAlerts = False
    On Error Resume Next
    NewDoc.Saved = True
    NewDoc.Close
    Set NewDoc = Nothing
    application.DisplayAlerts = wdAlertsAll
End Sub
   
Private Sub ���Ԃƍ��ږ���1�s��(ByRef rng As Range)
'�a�������������d�l���p
'���Ԃƍ��ږ����ʂ̍s�ɂ���O��
    With rng.Find
        .Text = "�o�n�r([ �@�D.^t�����a-z�`-�yA-Z0-9�O-�X]{1,})^13" '��������͂����̕������ς��đΉ�����
        .Replacement.Text = "�o�n�r\1^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Set rng = Nothing
End Sub

Private Sub �i������肾��(ByRef rng As Range, ByRef NewDoc As Word.Document)

    Dim eApp As Object
    Dim ewkb As Object
    Dim eWks As Object
    Dim sWhat As String
    Dim Supplier As String
    Dim Lists As Variant    '���o���镶����
    
    '�d�l���̎�ނɍ��킹�Č���������������
    sWhat = ���o�������p�^�[���ݒ�(Supplier)
    Set rng = Nothing
'   ������Range�ϐ����Đݒ�i�����ϐ��𑱂��Ďg���Ȃ��Ȃ����̂�[�����s��]�j
    Dim rng2 As Word.Range
    Set rng2 = NewDoc.Range(0, 0)
'   �z��̗v�f������������
    Dim cnt As Long
    cnt = GetItemCount(sWhat, rng2)
    
    Set rng2 = Nothing
    Dim rng3 As Word.Range
    Set rng3 = NewDoc.Range(0, 0)
'   �i�Ԃƍ��ږ��̋�؂�𐮗�
    Call �e�L�X�g���K��(sWhat, rng3, Supplier)
        '�f�o�b�O�p���K���ς݃e�L�X�g�̊m�F
    NormalizedLists = NewDoc.Range.Text
    Lists = �e�L�X�g�\����(cnt, sWhat, NewDoc)
'    Lists = �I�v�V�������ڃ}�[�L���O(cnt, NewDoc)

    '�z�񂪂Ȃ��i��������v���ʂ��Ȃ��j�ꍇ�͏I��
    If IsArrayEx(Lists) <> 1 Then
        MsgBox "��v���鍀�ڂ�����܂���", vbInformation, "���m�点"
        Exit Sub
    End If
        
    Select Case MatchMode
        Case False
            '�o�͗p���[�N�V�[�g��V�K�쐬����
            Set eApp = CreateObject("Excel.Application")
            eApp.Visible = True
            eApp.application.ScreenUpdating = False
            Set ewkb = eApp.workbooks.Add
            Set eWks = ewkb.sheets(1)
            
            '���[�N�V�[�g�Ɍ�����v���ʂ�\��t����
            Dim r2 As Excel.Range
            Set r2 = eWks.Range("A2").Resize(UBound(Lists, 1) + 1, _
                                                UBound(Lists, 2) + 1)
            r2.Value = Lists
        Case True
            Dim s
            Dim k, m
            For k = LBound(Lists, 1) To UBound(Lists, 1)
                For m = LBound(Lists, 2) To UBound(Lists, 2)
                    s = Replace(Lists(k, m), vbCr, "")
                    Lists(k, m) = s
                Next m
            Next k
            lbMatch.List = Lists
            Exit Sub
    End Select
        
    Call ��ԍ��o�^(eApp, eWks)
'       �I�v�V�����}�[�L���O�̓ǂݎ��i�v����Ƀt���O�����Ă��Ȃ��̂Ŏ��P��j
    Call �I�v�V�����t���O����(eApp, eWks)
    '�f�[�^�̐�������
    Call ��������(eApp, eWks)
    
    Set eApp = Nothing
    Set ewkb = Nothing
    Set eWks = Nothing
    Set r2 = Nothing
    
End Sub
Sub �I�v�V�����t���O����(eApp, eWks)
    Dim i
    Dim LastRow
    Dim buf
    
    With eWks
        LastRow = .Cells(.Rows.Count, 1).End(-4162).Row
        For i = StartRow To LastRow
            If .Cells(i, colRemark).Value <> Empty Then
                buf = Replace(.Cells(i, colRemark), vbCr, "") '���s���������Ă�
                If buf = flagOption Then
                    .Cells(i, colRemark).Value = Empty
                    .Cells(i, colOption).Value = flagOption
                End If
            End If
        Next
    End With
End Sub
Private Function �e�L�X�g�\����(cnt, What, ByRef NewDoc As Word.Document) As Variant
    Dim j As Long
    Dim c As Long
    Dim r As Word.Range
    Dim Lists() As Variant
    Dim List As Variant
    Dim myInstr As Long

    ReDim Lists(0 To cnt, 0 To 4) As Variant
''   �����p�^�[���ɋ�؂蕶���𖄂ߍ��ށB
'    myInstr = InStr(What, " ")
'    DelimeterLEFT = Left$(What, myInstr)
'    DelimeterRIGHT = Right$(What, Len(What) - myInstr)
'    Delimeter2 = DelimeterLEFT & "$" & DelimeterRIGHT
    
    Set r = NewDoc.Range(0, 0)
    With r.Find
        .Text = What & "*^13"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    
    '�q�b�g���Ȃ��Ȃ�܂Ō����𑱂���
    Do While r.Find.Execute = True And r.Text <> ""
        Dim buf
        buf = Replace(r.Text, vbCr, "")

        List = Split(r.Text, Delimeter)
        c = c + 1
        For j = LBound(List, 1) To UBound(List, 1)
            If j > UBound(Lists, 2) Then GoTo NextItem
            Lists(c - 1, j) = List(j)
        Next j
NextItem:
    Loop
    
    Set r = Nothing
    �e�L�X�g�\���� = Lists
    
End Function


Private Sub btnCopyToClipboard_Click()
    Call ExportListsToClipboard
End Sub

'���X�g�{�b�N�X����f�[�^��CSV�ŃN���b�v�{�[�h�ɃR�s�[����
Private Sub ExportListsToClipboard()
    Dim Lists As String
    Dim i, r, c
    Dim RowCnt, ColCnt

    With lbMatch
        ColCnt = .ColumnCount - 1
        RowCnt = .ListCount - 1

        For r = 0 To RowCnt
            For c = 0 To ColCnt
                Lists = Lists & vbTab & .List(r, c)
            Next c
            Lists = Lists & vbCrLf
        Next r
    End With

    'Clipboard�Ƀf�[�^������
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText Lists
        .PutInClipboard
    End With
End Sub

'Private Function �I�v�V�������ڃ}�[�L���O(cnt, ByRef NewDoc As Word.Document) As Variant
'    Dim j As Long
'    Dim c As Long
'    Dim r As Word.Range
'    Dim What
'    Dim Lists() As Variant
'    Dim List As Variant
'
'    ReDim Lists(0 To cnt, 0 To 4) As Variant
'
'    What = "option"
'
'    Set r = NewDoc.Range(0, 0)
'    With r.Find
'        .Text = What
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindStop
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchByte = False
'        .MatchAllWordForms = False
'        .MatchSoundsLike = False
'        .MatchFuzzy = False
'        .MatchWildcards = True
'    End With
'
'    '�q�b�g���Ȃ��Ȃ�܂Ō����𑱂���
'    Do While r.Find.Execute = True And r.Text <> ""
'        List = r.Text & "$OPTION"
'        c = c + 1
'        For j = LBound(List, 1) To UBound(List, 1)
'            Lists(c - 1, j) = List(j)
'        Next j
'    Loop
'
'    Set r = Nothing
'    �e�L�X�g�\���� = Lists
'
'End Function

Sub ��ԍ��o�^(ByRef eApp As Object, ByRef wks As Object)

    With wks
        .Cells(1, 1).Name = "POS"
        .Cells(1, 2).Name = "Item"
        .Cells(1, 3).Name = "EUR"
        .Cells(1, 4).Name = "Remark"
        .Cells(1, 5).Name = "Option"
        
        colPOS = .Range("POS").Column
        ColItem = .Range("Item").Column
        colEUR = .Range("EUR").Column
        colRemark = .Range("Remark").Column
        colOption = .Range("Option").Column
    End With

End Sub
Sub ��������(ByRef eApp As Object, ByRef wks As Object)
    Dim LastRow As Long
    Dim r1 As Object
    Dim r2 As Object
    Dim rEUR As Object
    Dim r4 As Object
    Dim r5 As Object
    Dim cntPOS As Long
    Dim cntItem As Long
    Dim cntEUR As Long
    Dim cntRemark As Long
    Dim cntOption As Long
    
    With wks
        LastRow = .Cells(.Rows.Count, 1).End(-4162).Row
        Set r1 = .Range(.Cells(StartRow, colPOS), .Cells(LastRow, colPOS))
        Set r2 = .Range(.Cells(StartRow, ColItem), .Cells(LastRow, ColItem))
        Set rEUR = .Range(.Cells(StartRow, colEUR), .Cells(LastRow, colEUR))
        Set r4 = .Range(.Cells(StartRow, colRemark), .Cells(LastRow, colRemark))
        Set r5 = .Range(.Cells(StartRow, colOption), .Cells(LastRow, colOption))
        
        rEUR.NumberFormatLocal = "#,##0_ "
        
        On Error GoTo ErrHandler
        With eApp.application.WorksheetFunction
            cntPOS = .CountA(r1)
            cntItem = .CountA(r2)
            cntEUR = .CountA(rEUR)
            cntRemark = .CountA(r4)
            cntOption = .CountA(r5)
        End With
        
BackToProcedure:

        '���o���ݒ�
        .Cells(1, colPOS).Value = "POS."
        .Cells(1, ColItem).Value = "�i��"
        .Cells(1, colEUR).Value = "EUR���i"
        .Cells(1, colRemark).Value = "���l"
        .Cells(1, colOption).Value = "�I�v�V�����H"
        .Range(.Cells(1, colPOS), .Cells(1, colOption)).Font.Bold = True
        
        .Cells(1, colPOS).Value = .Cells(1, colPOS).Value & "(" & cntPOS & ")"
        .Cells(1, ColItem).Value = .Cells(1, ColItem).Value & "(" & cntItem & ")"
        .Cells(1, colEUR).Value = .Cells(1, colEUR).Value & "(" & cntEUR & ")"
        .Cells(1, colRemark).Value = .Cells(1, colRemark).Value & "(" & cntRemark & ")"
        .Cells(1, colOption).Value = .Cells(1, colOption).Value & "(" & cntOption & ")"
        
        With .Columns("B")
            .WrapText = True
            .ColumnWidth = 50
        End With
        With .Columns("C")
            .WrapText = True
            .ColumnWidth = 10
        End With
        With .Columns("D")
            .WrapText = True
            .ColumnWidth = 20
        End With
    
        Call WashPriceString(wks, LastRow)
        
        .Columns("A:E").entirecolumn.AutoFit
        .UsedRange.Rows.EntireRow.AutoFit

    End With
    eApp.application.ScreenUpdating = True
    Set eApp = Nothing
    Set r1 = Nothing
    Set r2 = Nothing
    Set rEUR = Nothing
    Set r4 = Nothing
    Set r5 = Nothing
Exit Sub

ErrHandler:
    If Err.Number = 1004 Then
        cntItem = 0
        GoTo BackToProcedure
    End If
    Set eApp = Nothing
    Set r1 = Nothing
    Set r2 = Nothing
    Set rEUR = Nothing
    Set r4 = Nothing
    Set r5 = Nothing
End Sub

Private Function GetItemCount(What, rng) As Long
    Dim c
    
    With rng.Find
        .Text = What & "*^13"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    
    Do While rng.Find.Execute
        c = c + 1
    Loop
    GetItemCount = c
End Function

Private Function ���o�������p�^�[���ݒ�(Supplier) As String
    Dim Delimeter2
    Dim DelimeterLEFT As String
    Dim DelimeterRIGHT As String
    Dim myInstr As Long
    Dim buf As String
    
    If optBerents.Value Then Supplier = "Berents"
    If optGroninger.Value Then Supplier = "groninger"
    If optIWK.Value Then Supplier = "IWK"
    
    If chkManualCriteria.Value And txtManualCriteria.Text <> "" Then
        ���o�������p�^�[���ݒ� = txtManualCriteria.Text
    Else
        Select Case Supplier
            Case "Berents"
                buf = "(^13�o�n�r[�@ ^t]{1,})"
            Case "groninger"
                buf = "(^13[0-9]{1,4})[�@ ^t]{1,}"
            Case "IWK"
                buf = "([0-9]{5,6})[ �@^t]{1,}"
        End Select
    End If

'   �����p�^�[���ɋ�؂蕶���𖄂ߍ��ށB
    myInstr = InStr(buf, " ")
    DelimeterLEFT = Left$(buf, myInstr)
    DelimeterRIGHT = Right$(buf, Len(buf) - myInstr)
    ���o�������p�^�[���ݒ� = DelimeterLEFT & "$" & DelimeterRIGHT
End Function

Sub �e�L�X�g���K��(What As String, ByRef rng As Word.Range, Supplier)

    Select Case Supplier
        Case "Berents"
            If chkEN Then �e�L�X�g���K��_Berents_EN What, rng _
            Else: �e�L�X�g���K��_Berents_JA What, rng
        Case "groninger"
            If chkEN Then �e�L�X�g���K��_groninger_EN What, rng _
            Else: �e�L�X�g���K��_groninger_JA What, rng
        Case "IWK"
            If chkEN Then �e�L�X�g���K��_IWK_EN What, rng _
            Else: �e�L�X�g���K��_IWK_JA What, rng
    End Select
End Sub
Private Sub �e�L�X�g���K��_IWK_EN(What As String, ByRef rng As Word.Range)
    Call �󔒕�����؂萮��_IWK(What, rng)
    Call Tab�폜(rng)
    Call Remark��؂萮��(rng)
    Call EUR��؂萮��(rng)
    Call �I�v�V�������ʍ폜�ƃt���O����(rng)
    Call �]���ȋ�؂蕶���폜(rng)
    Call �]���ȃJ���}�폜(rng)
End Sub

Private Sub �e�L�X�g���K��_IWK_JA(What As String, ByRef rng As Word.Range)
    Call �S�p�X�y�[�X��؂萮��(rng)
    Call �󔒕�����؂萮��_IWK(What, rng)
    Call Tab�폜(rng)
    Call �]���ȋ�؂蕶���폜(rng)
    Call �]���ȃJ���}�폜(rng)
End Sub

Private Sub �e�L�X�g���K��_groninger_EN(What As String, ByRef rng As Word.Range)
    Call �󔒕�����؂萮��_groninger(What, rng)
    Call Tab�폜(rng)
    Call Remark��؂萮��(rng)
    Call EUR��؂萮��(rng)
    Call �I�v�V�������ʍ폜�ƃt���O����(rng)
    Call �]���ȋ�؂蕶���폜(rng)
    Call �]���ȃJ���}�폜(rng)
End Sub

Private Sub �e�L�X�g���K��_groninger_JA(What As String, ByRef rng As Word.Range)
    Call POS��̉��s������_groninger(rng)
    Call ���̌�ɉ��s������_groninger(rng)
    Call �󔒕�����؂萮��_groninger(rng)
'    Call �S�p�X�y�[�X��؂萮��(rng)
'    Call �󔒕�����؂萮��_groninger(What, rng)
'    Call Tab�폜(rng)
'    Call �]���ȋ�؂蕶���폜(rng)
'    Call �]���ȃJ���}�폜(rng)
End Sub

Private Sub �e�L�X�g���K��_Berents_EN(What As String, ByRef rng As Word.Range)
    Call �󔒕�����؂萮��_Berents(What, rng)
    Call Tab�폜(rng)
    Call Remark��؂萮��(rng)
    Call EUR��؂萮��(rng)
    Call �I�v�V�������ʍ폜�ƃt���O����(rng)
    Call �]���ȋ�؂蕶���폜(rng)
    Call �]���ȃJ���}�폜(rng)
End Sub
Private Sub �e�L�X�g���K��_Berents_JA(What As String, ByRef rng As Word.Range)
    Call �󔒕�����؂萮��_Berents(What, rng)
    Call Tab�폜(rng)
    Call Remark��؂萮��(rng)
    Call EUR��؂萮��(rng)
    Call �I�v�V�������ʍ폜�ƃt���O����(rng)
    Call �]���ȋ�؂蕶���폜(rng)
    Call �]���ȃJ���}�폜(rng)
End Sub
Sub �󔒕�����؂萮��_IWK(What As String, ByRef rng As Word.Range)
    Dim sReplace As String

    sReplace = "\1" & Delimeter
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub �󔒕�����؂萮��_groninger(ByRef rng As Word.Range)
    Dim sReplace As String
    Dim What
    'groninger�a���d�l�������Ƃɍ쐬�������[�� 2016/10/27
    sReplace = "\1" & Delimeter & "\2" & Delimeter & "\3\4" '\1$\2$\3\4
    What = "^13([0-9a-zA-Z]{4,5})" 'POS
    What = What & "[ ^t�@]{1,}"     'POS�ɂÂ���
    What = What & "([0-9�O-�Xa-zA-Z��-�I��-��@-�S�A/ �@\(\)�i�j]{1,})" '�i��
    What = What & "[^13 ^t�@]{1,}" '�i���ɂÂ���
    What = What & "([0-9�O-�X]{1,})[ ^t]{1,}(��)" '���ʁi�����܂ށj
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub POS��̉��s������_groninger(ByRef rng As Word.Range)
    Dim What
    Dim sReplace As String
   
   'groninger�a���d�l�������Ƃɍ쐬�������[�� 2016/10/27
    What = "(^13[0-9a-zA-Z]{1,5})^13"
    sReplace = "\1" & Delimeter
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub ���̌�ɉ��s������_groninger(ByRef rng As Word.Range)
    Dim What
    Dim sReplace As String
   
   'groninger�a���d�l�������Ƃɍ쐬�������[�� 2016/10/27
    What = "([0-9�O-�X]{1,2})[ ]{1,}(��)"
    sReplace = "\1\2" & Delimeter
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub �󔒕�����؂萮��_Berents(What As String, ByRef rng As Word.Range)
    Dim sReplace As String

    sReplace = "\1" & Delimeter

    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub Tab�폜(ByRef rng As Word.Range)
    Dim sReplace As String
    Dim What
    
    What = "^t"
    sReplace = ","

    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub EUR��؂萮��(ByRef rng As Word.Range)
'�i���̍s��EUR���i���܂܂��ꍇ�ɋ�؂蕶����}������
'��ŉ��i���z��ɓ���邽�߁B

    Dim sReplace As String
    Dim EURWhat As String
    
    EURWhat = "(EUR)"
    sReplace = Delimeter
    
    With rng.Find
        .Text = EURWhat
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub Remark��؂萮��(ByRef rng As Word.Range)
'�i���̍s��EUR���i���܂܂��ꍇ�ɋ�؂蕶����}������
'��ŉ��i���z��ɓ���邽�߁B

    Dim sReplace As String
    Dim What As String
    
    What = "(EUR[ ^t0-9.,]{1,})*"
    sReplace = "\1" & Delimeter
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub �S�p�X�y�[�X��؂萮��(ByRef rng As Word.Range)
'���{���IWK�d�l���ŕi���̌�ɑ����X�y�[�X����؂�

    Dim sReplace As String
    Dim What As String
    
    What = "([ �@^t]{2,})"
    sReplace = Delimeter
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub �I�v�V�������ʍ폜�ƃt���O����(ByRef rng As Word.Range)
'�P�FEUR���i�̃e�L�X�g�͈͂Ɋ܂܂�銇�ʁi�j������
'�Q�F�J���}�ƃh�b�g������

    Dim sReplace As String
    Dim What As String
    
    What = "$\(([0-9.,]{1,})\)*(^13)"
    sReplace = "\1" & "$OPTION" & "\2"
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub �]���ȋ�؂蕶���폜(ByRef rng As Word.Range)
'�悭�킩��Ȃ��Ȃ��ċ�؂蕶������������ł��Ă��܂��̂�1�Ɍ��炷
    Dim sReplace As String
    Dim What As String

    What = "$$"
    sReplace = "$"
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub �]���ȃJ���}�폜(ByRef rng As Word.Range)
'�悭�킩��Ȃ��Ȃ��ċ�؂蕶������������ł��Ă��܂��̂�1�Ɍ��炷
    Dim sReplace As String
    Dim What As String

    What = ",($),"
    sReplace = "\1"
    
    With rng.Find
        .Text = What
        .Replacement.Text = sReplace
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************
Private Function IsArrayEx(varArray As Variant) As Long
    On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

Private Sub WashPriceString(wks, LastRow)
    Dim i
    Dim buf
    Dim r As Excel.Range
    Dim Lists As Variant
    Dim List As Variant

    Set r = wks.Range(wks.Cells(3, 3), wks.Cells(LastRow, 3))
    Lists = r

    On Error Resume Next
    For i = LBound(Lists) To UBound(Lists)
        If Lists(i, 1) <> Empty Then
            buf = Lists(i, 1)
            buf = Replace(buf, ".", "") '�R���}����
            buf = Replace(buf, ",", "") '�h�b�g����
            Lists(i, 1) = buf
        End If
    Next
    
    r.Value = Lists
    On Error GoTo 0
End Sub

Private Sub ���p�S�p��(rng)
    '�S�p�p�����𔼊p�p�����ֈꊇ�ϊ�
    Dim Range As Word.Range
    Set Range = rng
    
    With Range.Find
        .Text = "[�O-�X]{5,6}"
        .MatchWildcards = True
        Do While .Execute = True
          Range.CharacterWidth = wdWidthHalfWidth
          Range.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Private Sub ���[�J�[�̏Z��������_IWK(doc)
    Dim r As Word.Range
    Dim r2 As Word.Range
    
    Set r = doc.Range(0, 0)
    
    With r.Find
        .Text = "76297[ ^t]{1,}Stutensee"
        .Replacement.Text = "a" '�u�������镶���͂Ȃ�ł������i��ɂ���ƒu���ł��Ȃ��j
        .Execute Replace:=wdReplaceAll
    End With
    Set r = Nothing
    
    With r2.Find
        .Text = "76133[ ^t]{1,}Karlsruhe"
        .Replacement.Text = "a" '�u�������镶���͂Ȃ�ł������i��ɂ���ƒu���ł��Ȃ��j
        .Execute Replace:=wdReplaceAll
    End With
    Set r2 = Nothing

End Sub

Private Sub btnGetNormalizedText_Click()
    Call GetDebugText
End Sub

Sub GetDebugText()
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText NormalizedLists
        .PutInClipboard
    End With
    
End Sub

Private Sub chkAutoRecognition_Click()
    Select Case chkAutoRecognition
        Case True
            frmManufacturer.Visible = False
            frmLanguageSelect.Visible = False
        Case False
            frmManufacturer.Visible = True
            frmLanguageSelect.Visible = True
    End Select
        
End Sub

Private Sub chkManualCriteria_Click()
    If chkManualCriteria.Value Then
        txtManualCriteria.Enabled = True
        txtManualCriteria.Locked = False
    Else
        txtManualCriteria.Enabled = False
        txtManualCriteria.Locked = True
    End If
End Sub

Private Sub btnOK_Click()
    Dim flag As Boolean
    Dim c As Control
    
    MatchMode = False
    For Each c In frmManufacturer.Controls
        If TypeName(c) = "OptionButton" Then _
        If c.Value Then flag = True
    Next c
    
    If flag Then
        Call ���ώd�l���̍��ږ��𒊏o
    Else
        MsgBox "���[�J�[��I�����Ă�������", vbInformation, "���m�点"
    End If
End Sub
Private Sub btnMatch_Click()
    Dim flag As Boolean
    Dim c As Control

    Call �ŋߎg�����t�@�C���������W�X�g���ɓo�^
    
    MatchMode = True
    
    For Each c In frmManufacturer.Controls
        If TypeName(c) = "OptionButton" Then _
        If c.Value Then flag = True
    Next c
    
    If flag Then
        Call ���ώd�l���̍��ږ��𒊏o
    Else
        MsgBox "���[�J�[��I�����Ă�������", vbInformation, "���m�点"
    End If
End Sub
Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub cmbDocumentName_DropButtonClick()
    Dim myPath As String
    If mEventCancel Then Exit Sub
    With application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath = .SelectedItems(1)
    End With
End Sub

Private Sub cmbDocumentName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mEventCancel = True
    If KeyCode = 40 Then cmbDocumentName.DropDown
    mEventCancel = False
End Sub

Private Sub UserForm_Initialize()
    With lbMatch
        .ColumnCount = 5
        .ColumnWidths = "50;200;50;20;20"
    End With
    chkJA.Value = True
    Call �ŋߊJ�����t�@�C���������W�X�g������ǂݏo��
End Sub

Private Sub CommandButton1_Click()
    Call ���X�g���̗]���ȃe�L�X�g������
End Sub
Private Sub ���X�g���̗]���ȃe�L�X�g������()
    Dim j, i
    Dim Lists As Variant
    Dim buf
    Dim buf2
    Dim RemovingStr
    Const Delimeter As String = ";"
    Const ColItem As Long = 1
    
    If txtRemovingStr.Text = "" Then Exit Sub
    If lbMatch.ListCount = -1 Then Exit Sub
     
    Select Case InStr(txtRemovingStr.Text, Delimeter)
        Case Is > 0 '�������镶������������ꍇ
            Lists = Split(txtRemovingStr.Text, Delimeter)
        
            For j = LBound(Lists) To UBound(Lists)
                RemovingStr = Trim$(Lists(j))
                For i = 0 To lbMatch.ListCount - 1
                    buf = lbMatch.List(i, ColItem)
                    If buf <> "" Then
                        Do While Right$(buf, 1) = RemovingStr
                            buf = Left$(buf, Len(buf) - 1)
                            lbMatch.List(i, ColItem) = buf
                        Loop
                    End If
                Next i
            Next j
        
    Case Else '�������镶�����ЂƂ����̏ꍇ
        RemovingStr = Trim$(txtRemovingStr.Text)
        For i = 0 To lbMatch.ListCount - 1
            buf = lbMatch.List(i, ColItem)
            If buf <> "" Then
                Do While Right$(buf, 1) = RemovingStr
                    buf = Left$(buf, Len(buf) - 1)
                    lbMatch.List(i, ColItem) = buf
                Loop
            End If
        Next i
    End Select
End Sub

Private Sub �G�f�B�^�ŕ��������(doc)
    Dim List As String
    Dim myPath
    Dim i
    Dim lines As Variant
    
'   ���[�h�����Ɋ܂܂�Ă���s���ȕ�������������邽��
'  ��x�G�f�B�^�ɃR�s�[���Ă���Ăю��o��
    myPath = ThisDocument.Path & "\" & "bufItemExtraction_Specification.txt"
    List = doc.Range.Text
    If List = "" Then Exit Sub
    
    Dim buf As String
    Open myPath For Output As #1
        Print #1, List
    Close #1

    Dim buf2 As String
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(myPath).OpenAsTextStream
            buf2 = .ReadAll
            .Close
        End With
    End With
    
    doc.Range.Text = Empty
    doc.Range.Text = buf2

''   �o�͂����t�@�C�����J��
'    Shell "notepad " & myPath, vbNormalFocus
Exit Sub
Err:

End Sub

Private Sub ���[�J�[�̎����}�b�`(doc)
    Dim i
    Dim Supplier
    Dim r
    Dim f
    Dim Name
    Supplier = Array("IWK", "groninger", "Berents", "Becomix")
    
    Set r = doc.Range(0, 0)
    With r.Find
        .Text = Supplier(0)
        f = .Execute
    End With
    If f Then
        optIWK.Value = True
        Exit Sub
    End If
    Set r = Nothing
    
    Dim r2
    Set r2 = doc.Range(0, 0)
    With r2.Find
        .Text = Supplier(1)
        f = .Execute
    End With
    If f Then
        optGroninger.Value = True
        Exit Sub
    End If
    Set r2 = Nothing
    
    Dim r3
    Set r3 = doc.Range(0, 0)
    With r3.Find
        .Text = Supplier(2)
        f = .Execute
    End With
    If f Then
        optBerents.Value = True
        Exit Sub
    End If
    Set r3 = Nothing

    Dim r4
    Set r4 = doc.Range(0, 0)
    With r4.Find
        .Text = Supplier(3)
        f = .Execute
    End With
    If f Then
        optBerents.Value = True
        Exit Sub
    End If
    Set r4 = Nothing
End Sub

Sub lbMatch_Change()
    lbItemCount.Caption = lbMatch.ListCount - 1
    lbListIndex.Caption = lbMatch.ListIndex
End Sub

'Private Sub �p����_�����S()
'    '���p�p������S�p�p�����ֈꊇ�ϊ�
'    Dim myRange As Range
'    Dim blnFound As Boolean
'    application.ScreenUpdating = False
'    Set myRange = ActiveDocument.Range(0, 0)
'    With myRange.Find
'        .Text = "[0-9A-Za-z]{1,}"  '�Ώۂ̐ݒ�
'        .MatchWildcards = True
'        Do While .Execute = True
'          blnFound = True
'          myRange.HighlightColorIndex = wdTurquoise  '�F�̐ݒ�
'          myRange.CharacterWidth = wdWidthFullWidth
'          myRange.Collapse wdCollapseEnd
'        Loop
'    End With
'    Set myRange = Nothing
'
'    If blnFound = True Then
'      MsgBox "���p�p������S�p�ɕϊ����܂����B"  '���b�Z�[�W
'    End If
'    application.ScreenUpdating = True
'End Sub
'
'Private Sub �p����_�S����()
'    '�S�p�p�����𔼊p�p�����ֈꊇ�ϊ�
'    Dim myRange As Range
'    Dim blnFound As Boolean
'    application.ScreenUpdating = False
'    Set myRange = ActiveDocument.Range(0, 0)
'    With myRange.Find
'        .Text = "[�O-�X�`-�y��-��]{1,}"  '�Ώۂ̐ݒ�
'        .MatchWildcards = True
'        Do While .Execute = True
'          blnFound = True
'          myRange.HighlightColorIndex = wdBrightGreen  '�F�̐ݒ�
'          myRange.CharacterWidth = wdWidthHalfWidth
'          myRange.Collapse wdCollapseEnd
'        Loop
'    End With
'    Set myRange = Nothing
'
'    If blnFound = True Then
'      MsgBox "�S�p�p�����𔼊p�ɕϊ����܂����B"  '���b�Z�[�W
'    End If
'    application.ScreenUpdating = True
'End Sub
Private Sub spinUpDownItem_SpinDown()
    Dim strListItemDown() As String
    Dim strListItemSelected() As String
    Dim i As Long
    Dim blIsSelected As Boolean
    Dim r, r2, r3
    Dim ColCnt
    Dim col
    Dim ItemCnt
    Dim ListRow
        
'   ���X�g���I������Ă��Ȃ���ΏI��
    With lbMatch
        For i = 0 To .ListCount - 1
            If .Selected(i) Then blIsSelected = True
            ItemCnt = ItemCnt + 1
        Next i
        If Not blIsSelected Then Exit Sub
    End With
    
    ColCnt = lbMatch.ColumnCount - 1
    ReDim strListItemDown(ItemCnt - 1, ColCnt) As String
    
'   ����ȏ㉺���Ȃ���ΏI������
    With lbMatch
        If .ListIndex + 1 < .ListCount Then
            For r = 0 To .ListCount - 1
                Select Case .Selected(r)
                    Case True
                        For col = 0 To ColCnt
                            strListItemDown(col) = .List(ListRow, col)
                        Next col
                        ListRow = ListRow + 1
                    Case Else
                        '�������Ȃ��B���̃��X�g���ڂ�
                End Select
            Next r
            
            ReDim strListItemSelected(ColCnt) As String
            For r2 = 0 To ColCnt
                strListItemSelected(r2) = .List(.ListIndex, r2)
                .List(.ListIndex, r2) = strListItemDown(r2)
                .List(.ListIndex + 1, r2) = strListItemSelected(r2)
            Next r2
        End With
    lbMatch.Selected(lbMatch.ListIndex + 1) = True
    
End Sub

Private Function EvalJADoc(doc) As Boolean
    Dim List As String
    Dim para As Word.Paragraph
    Dim p As Word.Paragraph
    Dim f
    Dim Regex As Object
    
'    For Each p In doc.Paragraphs
'        list = Left$(p.Range.Text, Len(p.Range.Text) - 1)
'        f = InStr(list, vbCr)
'        If f <> 0 Then
'            Debug.Print "vbcr"
'            list = Trim(Left$(list, f - (Len(list) - f - 1)))
'        End If
'        If list <> "" Then
'            Debug.Print list
'            If StrConv(list, vbWide) = StrConv(list, vbNarrow) Then
'                Debug.Print list
'                EvalJADoc = True '�a�����聁�����S�̂̂Ȃ��ɂЂ炪�Ȃ��܂܂�Ă��邩�ǂ���
'                Debug.Print "true"
'                'Exit Function
'            End If
'        End If
'    Next

'   ���K�\���N���X�𗧂��グ��
    Set Regex = CreateObject("VBScript.RegExp")


    List = doc.Range.Text
    With Regex
        .IgnoreCase = True
        .Pattern = "[^\x00-\x7F]"
    End With
    
    EvalJADoc = Regex.test("��") '�p��
    Debug.Print EvalJADoc
    Select Case EvalJADoc
        Case True: EvalJADoc = True '�a��
        Case False: EvalJADoc = False '�p��
    End Select
    MsgBox EvalJADoc
End Function

Private Sub spinUpDownItem_SpinUp()
    Dim strListItemUp() As String
    Dim strListItemSelected() As String
    Dim i As Long
    Dim blIsSelected As Boolean
    Dim r, r2
    Dim ColCnt

    With lbMatch
        For i = 0 To .ListCount - 1
            If .Selected(i) Then blIsSelected = True
        Next i
        If Not blIsSelected Then Exit Sub        '���X�g���I������Ă��Ȃ���ΏI��
    End With
    
    ColCnt = lbMatch.ColumnCount - 1
    ReDim strListItemUp(ColCnt) As String
    
    With lbMatch
        If .ListIndex > 0 Then
            For r = 0 To ColCnt
            strListItemUp(r) = .List(.ListIndex - 1, r)
            Next r
        Else
            Exit Sub    '����ȏ�A�オ�Ȃ���ΏI������
        End If

        ReDim strListItemSelected(ColCnt) As String
        For r2 = 0 To ColCnt
        strListItemSelected(r2) = .List(.ListIndex, r2)
        .List(.ListIndex, r2) = strListItemUp(r2)
        .List(.ListIndex - 1, r2) = strListItemSelected(r2)
        Next r2
    End With
        lbMatch.Selected(lbMatch.ListIndex - 1) = True

End Sub

Sub �ŋߎg�����t�@�C���������W�X�g���ɓo�^()
    Dim i
    Dim buf
    Dim flag
    Dim msg
    Dim v
    Const Delimeter As String = vbCrLf
    msg = "No Recent Files."
    buf = GetSetting("MyMacroData", "WordExtractItemNames", "RecentFileNames", msg)

    Select Case buf
        Case msg
            SaveSetting "MyMacroData", "WordExtractItemNames", _
            "RecentFileNames", cmbDocumentName.Text & vbCrLf
        Case Else
            v = Split(buf, Delimeter)
            For i = LBound(v) To UBound(v)
                If v(i) = buf Then flag = True
            Next
            If Not flag Then _
                    SaveSetting "MyMacroData", "WordExtractItemNames", _
                    "RecentFileNames", cmbDocumentName.Text & vbCrLf
    End Select
    
'    MsgBox "registry:" & buf
End Sub

Sub �ŋߊJ�����t�@�C���������W�X�g������ǂݏo��()
    Dim i
    Dim buf
    Dim flag
    Dim msg
    Dim v
    Const Delimeter As String = vbCrLf
    msg = "No Recent Files."
    buf = GetSetting("MyMacroData", "WordExtractItemNames", "RecentFileNames", msg)

    Select Case buf
        Case msg
            Exit Sub
        Case Else
            v = Split(buf, Delimeter)
            For i = LBound(v) To UBound(v)
                cmbDocumentName.AddItem v(i)
            Next
    End Select
End Sub
