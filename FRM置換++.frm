VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM�u���}�N��Ver5 
   Caption         =   "�u��"
   ClientHeight    =   6030
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "FRM�u���}�N��Ver7.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRM�u���}�N��Ver5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim myMenu
Dim mHasMatch As Boolean
Dim mUseHistory As Boolean
Dim myHighlightColor As Long
Dim myPath As String
Dim mCancelEvent As Boolean
Dim IsMinimized  As Boolean
Public lLeft As Long
Public lTop As Long

'�J���[�s�b�J�[�_�C�A���O�̃R�[�h+++++++++++++++++++++++++++++++++++++++++++++++++
Private Declare Function ChooseColor Lib "comdlg32.dll" _
    Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
 
Private Type ChooseColor
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
 
Private Const CC_RGBINIT = &H1
Private Const CC_LFULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8
 
Public Function GetColorDlg(lngDefColor As Long) As Long
 
    Dim udtChooseColor As ChooseColor
    Dim lngRet As Long
 
    With udtChooseColor '�_�C�A���O�̐ݒ�
      .lStructSize = Len(udtChooseColor)
      .lpCustColors = String$(64, Chr$(0))
      .flags = CC_RGBINIT + CC_LFULLOPEN
      .rgbResult = lngDefColor
    End With
    lngRet = ChooseColor(udtChooseColor) '�_�C�A���O��\��
    
    If lngRet <> 0 Then '�_�C�A���O����̖߂�l���`�F�b�N
        If udtChooseColor.rgbResult > RGB(255, 255, 255) Then
          GetColorDlg = -2 '�G���[�̏ꍇ
        Else
          GetColorDlg = udtChooseColor.rgbResult '�߂�l��RGB�l����
        End If
    Else
      GetColorDlg = -1 '�L�����Z�����ꂽ�ꍇ
    End If
 
End Function

Private Sub btnMinimize_Click()
    If Not IsMinimized Then
        Me.Width = 383.5
        Me.Height = 200
    Else
        Me.Width = 383.5
        Me.Height = 329.5
    End If
End Sub

Private Sub ckReplaceAllDocs_Click()
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub ckMatchCase_Click()
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub ckUseWildCards_Click()
    If ck���K�\���I�v�V����.Value Then _
    ck���K�\���I�v�V����.Value = Not ckUseWildCards.Value
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub ck���K�\���I�v�V����_Click()
    If ckUseWildCards.Value Then _
    ckUseWildCards.Value = Not ck���K�\���I�v�V����.Value
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub cmbSelectContext_Change()
    Call UpdateHistoryListBox
End Sub

Private Sub lblSelectColor_Click()
    Dim Color As Long
   
    With FRMColorPicker
        .Show vbModal
    End With
End Sub

Private Sub CommandButton13_Click()
    Dim myDict As String
    myDict = ThisDocument.Path & "\�u������.txt"
    myDict = """" & myDict & """" 'Wscript.Shell�ɓn�������Ɋ܂܂�遏���G�X�P�[�v
    CreateObject("Wscript.Shell").Run myDict, 5
End Sub

Private Sub CommandButton6_Click()
    Unload Me
'    Me.Hide
End Sub

Private Sub Mp_Change()
    Dim indPage As Long
    indPage = Mp.Value
    
    Select Case indPage
        Case 0
            With txtReplaceWords
                .SetFocus
                .SelStart = 0
            End With
        Case 1
            lbxHistory.SetFocus
            Call UpdateHistoryListBox
    End Select

End Sub

Private Sub optChangeFontColor_Click()
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub optUseHighlight_Click()
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
    
    If Not mCancelEvent Then Call PopupColorSelect
    mCancelEvent = False
End Sub

Private Sub PopupColorSelect()
 Call lblSelectColor_Click
End Sub

Private Sub optNoHighlight_Click()
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub txtReplaceWords_Change()
    Dim Str As String: Str = txtReplaceWords.Text
    If Not Str = "" Then
        SaveInput (Str)
        SetReplaceNotFinished '�u���͌�喢�u���t���O�v�𗧂Ă�
    End If
End Sub

'********************************************************
'���W�X�g���֌W
'�e�L�X�g�{�b�N�X�̓��͂����W�X�g���Ƀo�b�N�A�b�v����B
'�d���t�@�C���𓯎��ҏW����Ƃ���Word�͂悭������̂ŁB
'********************************************************
Private Sub SaveInput(Str As String)
    SaveSetting "MyMacro", "BulkReplace", "InputWords", Str
End Sub
Private Function GetSavedStr() As String
    GetSavedStr = GetSetting("MyMacro", "BulkReplace", "InputWords")
End Function
Private Sub SetReplaceNotFinished()
    SaveSetting "MyMacro", "BulkReplace", "IsReplaced", False
End Sub
Private Sub SetReplaceDone()
'   ���W�X�g���̒u�������t���O�𗧂Ă�B
'  ���W�X�g���Ƀo�b�N�A�b�v���Ă���u����������
    SaveSetting "MyMacro", "BulkReplace", "IsReplaced", True
    SaveSetting "MyMacro", "BulkReplace", "InputWords", ""
End Sub
Private Function CheckReplaceStatus() As Boolean
    On Error GoTo Err
    CheckReplaceStatus = GetSetting("MyMacro", "BulkReplace", "IsReplaced")
    On Error GoTo 0
Err:
    '���W�X�g�����Ȃ��ꍇ�͍��'
    SaveSetting "MyMacro", "BulkReplace", "Isreplaced", False
End Function

'********************************************************
Private Sub UserForm_Initialize()

'   �C�j�V�����C�Y���ɃI�v�V�����{�^���̃N���b�N�C�x���g����������̂�����邽�߂̃C�x���g����ϐ�
    mCancelEvent = True
    
'   �u�������̏ꏊ���w��
    myPath = ThisDocument.Path & "\�u������.txt"

'   ���͂������̍ēǂݍ��݂𔻒�i�����I���Ȃǂɂ����͓r���̃f�[�^�𕜌�����j
    Dim Replaced As Boolean
    Replaced = CheckReplaceStatus
    If Not Replaced Then txtReplaceWords.Text = GetSavedStr

'   �u���������X�g�{�b�N�X�̐ݒ�
    lbxHistory.ColumnCount = 2
'    lbxHistory.MultiSelect = fmMultiSelectExtended

'   �R���g���[��ON/OFF�̏����ݒ�
    Me.ckReplaceAllDocs = False
    Me.ckMatchCase = False
    Me.ckUseWildCards = False
    
'   �n�C���C�g�F�̏����ݒ�
    lblSelectColor.BackColor = vbGreen
    lblSelectColor.Tag = 7
    optUseHighlight.Value = True

    Call SetRightClickMenu
    Call �R���g���[���v���p�e�B�Ǎ�
    Call �u���������e�L�X�g�{�b�N�X�ɕ\������
    
'   �e�L�X�g�{�b�N�X�Ƀt�H�[�J�X
    StartUpPosition = 1
    mCancelEvent = True
    Me.Show
    Me.Mp.Value = 0
    Me.txtReplaceWords.Visible = False
    Me.txtReplaceWords.Visible = True
    Me.txtReplaceWords.SetFocus
    Me.txtReplaceWords.SelStart = 0
    mCancelEvent = False
End Sub
Private Sub cmdExecute_Click()
    Dim WordLists As Variant
    WordLists = GetWhatReplace
    If Not IsArray(WordLists) Then Exit Sub

    '�u���͈́A���@�ɂ����4�ʂ�ɏ�������
    '�͈́F�A�N�e�B�u�h�L�������g�b�S�h�L�������g�G�@�u�����@�F�W���u���b���K�\��
    Select Case ckReplaceAllDocs.Value
        Case False
            If ck���K�\���I�v�V����.Value Then
                RegexReplace WordLists, ActiveDocument
            Else
                Replaces WordLists, ActiveDocument
            End If
        Case True
            Dim doc As Document
            If ck���K�\���I�v�V����.Value Then
                For Each doc In Documents
                    RegexReplace WordLists, doc
                Next doc
            Else
                For Each doc In Documents
                    Replaces WordLists, doc
                Next doc
            End If
    End Select

'   �u�����ʂ�\��(�}�b�`�����������Ȃ��������m�点��)
    ShowCompletionMsg (mHasMatch)
    mHasMatch = False
    
'   �����E�u����𗚗��ɕۑ�����
    �u�������ɓo�^ WordLists
'   �������g���Ēu������ꍇ�́A���������烊�X�g�{�b�N�X�̑I������������
    Dim i As Long
    If mUseHistory Then
        For i = 0 To lbxHistory.ListCount - 1
            If lbxHistory.Selected(i) Then lbxHistory.Selected(i) = False
        Next i
        mUseHistory = False
    End If
    SetReplaceDone
    Call �t�H�[�J�X����
End Sub

Private Sub Replaces(WordLists, doc As Word.Document)
'   �����FWordList=������E�u����,doc=���[�h����

    Dim i As Long
    Dim sWhatReplace As Variant
    Dim rng As Range
    Set rng = doc.Range(0, 0)

    '�����ݒ�
    '�u���y���F�B�F�̒萔�����������x�����Q�Ƃ���B
    Options.DefaultHighlightColorIndex = Me.lblSelectColor.Tag
    With rng.Find
        '���C���h�J�[�h
        If ckUseWildCards.Value Then _
            .MatchWildcards = True Else: .MatchWildcards = False
        If ckMatchCase.Value Then _
            .MatchCase = True Else: .MatchCase = False
        '�u���t�H���g�F
        If optChangeFontColor.Value Then .Replacement.Font.Color = wdColorRed
        If optUseHighlight.Value Then .Replacement.Highlight = True
    End With

    Dim WhatStr As String, ReplaceStr As String
    For i = LBound(WordLists) To UBound(WordLists)
        sWhatReplace = VBA.Split(WordLists(i), vbTab)
        WhatStr = sWhatReplace(0)
        ReplaceStr = sWhatReplace(1)

        '�󔒂ƒu���i������̍폜�j��0�Ƃ���ȊO
        Select Case Len(ReplaceStr)
            Case Is > 0
                With rng.Find
                    .Text = WhatStr
                    .Replacement.Text = ReplaceStr
                    If .Execute = True Then mHasMatch = True
'                   ���K�\���I�v�V�������I�t�̂Ƃ��A���͂ɐ��K�\�����g�p������G���[�ɂȂ������߈ꎞ�I�ɃG���[������
                    On Error Resume Next
                    .Execute Replace:=wdReplaceAll
                    On Error GoTo 0
                End With
            Case 0
                Call ReplaceWithEmpty(doc, WhatStr, mHasMatch)
        End Select
    Next i

    '�e�L�X�g�{�b�N�X�̒u��
    With Selection.Find
        Select Case ckMatchCase.Value
            Case True: .MatchCase = True
            Case False: .MatchCase = False
        End Select
        If optChangeFontColor.Value = True Then .Replacement.Font.Color = wdColorRed
        If optUseHighlight.Value = True Then .Replacement.Highlight = True
    End With

    Dim sp As Shape
    For Each sp In doc.Shapes
        If sp.Type = msoTextBox Then
            sp.Select

            For i = LBound(WordLists) To UBound(WordLists)
               On Error Resume Next
               sWhatReplace = VBA.Split(WordLists(i), vbTab)
               On Error GoTo 0
                WhatStr = sWhatReplace(0)
                ReplaceStr = sWhatReplace(1)
               Selection.Find.ClearFormatting
               Select Case Len(ReplaceStr)
                   Case Is > 0
                       With Selection.Find
                           .Text = WhatStr
                           .Replacement.Text = ReplaceStr
                           If .Execute = True Then mHasMatch = True
                           .Execute Replace:=wdReplaceAll
                       End With
                   Case 0
                       Call ReplaceWithEmpty(doc, WhatStr, mHasMatch)
               End Select
            Next i
        End If
    Next sp

    Set rng = Nothing
    Set doc = Nothing
Exit Sub

ErrHandler:
    Dim msg As String
    msg = "�G���[�ԍ��F" & Err.Number & vbCrLf & _
          "�G���[���e�F" & Err.Description
    MsgBox msg, vbCritical, "�G���[�I��"
    Set rng = Nothing
    Set doc = Nothing
End Sub

Private Function GetWhatReplace() As Variant
'���[�U�[�t�H�[�����猟����ƒu������擾����
'�V�K���̓y�C���Ɨ����y�C���̂ǂ��炪�I������Ă��邩�ɉ����ď������򂳂���
'�e�L�X�g�{�b�N�X����u�������擾���āA���s��؂�Ŕz��ɂ���

    Dim intPage As Long
    intPage = Mp.Value
    Dim buf As String
    Dim List As String
    Dim Lists As Variant
    'intPage=Mp.Value�F�V�K���̓y�C��=0�A�����y�C��=1
    Select Case intPage
        Case 0: buf = txtReplaceWords.Text
        Case 1
            buf = �����̑I�����擾
            mUseHistory = True
    End Select

    List = RemoveEmptyRows(buf)
    GetWhatReplace = VBA.Split(List, vbCrLf)
End Function

Private Function �����̑I�����擾() As String
    Dim f As Boolean
    Dim IsAnySelected As Boolean
    Dim c As Long, i As Long
    Dim SelItems As String
    
    If lbxHistory.ListIndex > -1 Then f = True
    If f = False Then Exit Function
    For i = 0 To lbxHistory.ListCount - 1
        If lbxHistory.Selected(i) Then
            IsAnySelected = True
            SelItems = SelItems & lbxHistory.List(i, 0) & _
                        vbTab & lbxHistory.List(i, 1) & vbCrLf
        End If
    Next
    If SelItems = "" Then Exit Function
    SelItems = Left$(SelItems, Len(SelItems) - 1)
    �����̑I�����擾 = SelItems
End Function
Private Sub �t�H�[�J�X����()
    Dim ind As Long: ind = Mp.Value
    Select Case ind
        Case 0: txtReplaceWords.SetFocus
        Case 1: lbxHistory.SetFocus
    End Select
End Sub

Private Function RemoveEmptyRows(targetStr As String)
    '�e�L�X�g�{�b�N�X�̋�s���폜����B
    '��s������ƃo�O��B
    Dim ret As String
    ret = RegularExpressions.RegexReplace(targetStr, "(\r){1,}$", "")
    ret = RegularExpressions.RegexReplace(targetStr, "(\r\n){1,}$", "")
    RemoveEmptyRows = ret
End Function

Private Sub RegexReplace(WordLists, doc As Word.Document)

    On Error GoTo ErrHandler
    With doc
    Dim rng As Range
    Dim para As Paragraph
        For Each para In .Paragraphs
            Set rng = para.Range
'           rng������s�������i��������Ȃ��Ɖ��s��������j
            rng.MoveEnd unit:=wdCharacter, Count:=-1

            Dim i As Long
            For i = LBound(WordLists) To UBound(WordLists)
'               �u�����X�g�ɋ�s���������ꍇ�A��������
                On Error Resume Next
                Dim sWhatReplace As Variant
                sWhatReplace = VBA.Split(WordLists(i), vbTab)
                On Error GoTo 0
'               �u����僊�X�g����s���������΂�
                If UBound(sWhatReplace) = -1 Then GoTo NextRow

                Dim Reg As String: Reg = sWhatReplace(0)
                Dim ReplaceStr As String: ReplaceStr = sWhatReplace(1)
                Dim ret As String
                Dim targetStr As String

                targetStr = rng.Text
                If targetStr = "" Then GoTo NextRow
                If RegularExpressions.RegexTest(targetStr, Reg) Then mHasMatch = True
                rng.Text = RegularExpressions.RegexReplace(targetStr, Reg, ReplaceStr)
            Next i
NextRow:
        Next para
    End With

    '�e�L�X�g�{�b�N�X�̒u��
    Dim sp As Shape
    With doc
        For Each sp In .Shapes
            If sp.Type = msoTextBox Then
                sp.Select
                 For i = LBound(WordLists) To UBound(WordLists)
'                   �u�����X�g�ɋ�s���������ꍇ�A��������
                    On Error Resume Next
                    sWhatReplace = VBA.Split(WordLists(i), vbTab)
                    On Error GoTo 0
'                   �u����僊�X�g����s���������΂�
                    If UBound(sWhatReplace) = 0 Then GoTo NextRow2
                    Selection.Find.ClearFormatting
                    Selection.WholeStory

                    Reg = sWhatReplace(0)
                    ReplaceStr = sWhatReplace(1)
                    targetStr = Selection.Text
                    If targetStr = "" Then GoTo NextRow2
                    If RegularExpressions.RegexTest(targetStr, Reg) Then mHasMatch = True
                    Selection.Text = RegularExpressions.RegexReplace(targetStr, Reg, ReplaceStr)
NextRow2:
                Next i
            End If
        Next sp
    End With

    Set rng = Nothing
    Set doc = Nothing
Exit Sub

ErrHandler:
    Dim msg As String
    msg = "�G���[�ԍ��F" & Err.Number & vbCrLf & _
          "�G���[���e�F" & Err.Description
    MsgBox msg, vbCritical, "�G���[�I��"
    Set rng = Nothing
    Set doc = Nothing

End Sub

Private Sub ShowCompletionMsg(HasAnyMatch)
    Select Case HasAnyMatch
        Case False
            MsgBox "���������ɓ��Ă͂܂���̂͌�����܂���ł���", vbInformation
            Exit Sub
        Case True
            MsgBox "�������܂����B", vbOKOnly + vbInformation, "����"
    End Select
End Sub

Private Sub ���C���h�J�[�h�ƃt�H���g�F�̐ݒ�(rng, HasSet)
    HasSet = True
    With rng.Find
        Select Case ckUseWildCards
            Case True: .MatchWildcards = True
            Case False: .MatchWildcards = False
        End Select
            
    '   �u����̃t�H���g�F
        If optChangeFontColor.Value Then .Replacement.Font.Color = wdColorRed
        If optUseHighlight.Value Then .Replacement.Highlight = True
    End With
End Sub

Private Sub ReplaceWithEmpty(doc As Word.Document, What As Variant, mHasMatch As Boolean)
    Dim ret As String
    Dim targetStr As String
    Dim Reg As String
    Dim ReplaceStr As String: ReplaceStr = ""
    Dim sp As Word.Shape
    Dim rng As Word.Range
    Dim para As Word.Paragraph

    '�{����u��
    For Each para In doc.Paragraphs
        Set rng = para.Range
'       ���s�������������������Q�Ƃ���i�ƂĂ���؁j
        rng.MoveEnd unit:=wdCharacter, Count:=-1
        targetStr = rng.Text
        If targetStr = "" Then GoTo NextPara
        Reg = What
        If RegularExpressions.RegexTest(targetStr, Reg) Then mHasMatch = True
        ret = RegularExpressions.RegexReplace(targetStr, Reg, ReplaceStr)
        rng.Text = ret
NextPara:
    Next para

    '�e�L�X�g�{�b�N�X��u��
    For Each sp In doc.Shapes
        If sp.Type = msoTextBox Then
            sp.Select
            Selection.Find.ClearFormatting
            Selection.WholeStory
            targetStr = Selection.Text
            If targetStr = "" Then GoTo NextPara2
            Reg = What
            If RegularExpressions.RegexTest(targetStr, Reg) Then mHasMatch = True
            ret = RegularExpressions.RegexReplace(targetStr, Reg, ReplaceStr)
            Selection.Text = ret
NextPara2:
        End If
    Next sp
End Sub

Private Sub �u�������ɓo�^(WhatReplace As Variant)

    If Not IsArray(WhatReplace) Then Exit Sub
    Dim Lines As Variant
    Dim v As Variant
    Lines = LoadHistory
    v = VBA.Split(Lines, vbCr)

    Dim HasContent As Boolean
    Select Case UBound(v)
        Case Is > 0: HasContent = True
        Case Is = -1: HasContent = False
    End Select

    Dim i As Long
    Dim wordColl As New Collection
    '�����̒��g��Collection�ɏ�������
    '�d��������ƃG���[�ɂȂ�̂�Resume�ɂ��Ă���
    On Error Resume Next
    For i = LBound(v) To UBound(v)
        wordColl.Add v(i)
    Next i
    On Error GoTo 0

'   �����E�u���オ�V������ł���Η����e�L�X�g�t�@�C���ɏ�������
    Dim j As Long, k As Long
    Dim IsDupe As Boolean
    Open myPath For Append As #1
    For j = LBound(WhatReplace) To UBound(WhatReplace)
        For k = LBound(v) To UBound(v)
            If v(k) = WhatReplace(j) Then IsDupe = True
        Next k
        If Not IsDupe Then Print #1, WhatReplace(j)
    Next j
    Close #1
End Sub

Private Sub �R���g���[���v���p�e�B�Ǎ�()
    Dim c As Long: c = 1
    Dim Properties() As String
    Dim sProperties As String
    Dim CtrlCount As Long
    Dim ctrl As Control
    Dim v As Variant
    
    CtrlCount = Me.Controls.Count
    ReDim Properties(1 To CtrlCount, 1 To 5) As String
    
    For Each ctrl In Controls
        Properties(c, 1) = ctrl.Name
        Properties(c, 2) = ctrl.BackColor
        Properties(c, 3) = ctrl.ForeColor
        Properties(c, 4) = ctrl.FontName
        Properties(c, 5) = ctrl.FontSize
        c = c + 1
    Next
    
    For Each ctrl In Controls
        sProperties = sProperties & "," & ctrl.Name
        sProperties = sProperties & "," & ctrl.BackColor
        sProperties = sProperties & "," & ctrl.ForeColor
        sProperties = sProperties & "," & ctrl.FontName
        sProperties = sProperties & "," & ctrl.FontSize & vbCr
    Next

    Me.Tag = sProperties
    v = VBA.Split(sProperties, ",")
End Sub

Private Sub �u���������e�L�X�g�{�b�N�X�ɕ\������()
    Dim DicWords() As String
    Dim v As Variant
    Dim TabSeparatedStr As Variant
    Dim c As Long: c = 0
    Dim cnt As Long
    Dim Lists As Variant, List As Variant
    Dim i As Long
    Dim ListsReduced As String
 
    On Error GoTo Err
Return1:
    
    Lists = LoadHistory
    Lists = VBA.Split(Lists, vbCr)
    
    Dim HasSeveralEntries
    Select Case UBound(Lists)
        Case Is > 1: HasSeveralEntries = True
        Case Is <= 1: HasSeveralEntries = False
    End Select
    
    Select Case HasSeveralEntries
        Case True
            ReDim Preserve Lists(LBound(Lists) To UBound(Lists) - 1) As String
        Case False
            ReDim Preserve Lists(UBound(Lists) - 1) As String
    End Select

    Dim l As Long
    Dim Line As String
    On Error Resume Next ' ������ƒu���オ������Ă��Ȃ��ƃG���[�ɂȂ�
    For l = UBound(Lists) To 0 Step -1
'       ��ɏd�����m���߂邽�߁A�����̂P�s���^�u��؂�œ񎟌��z��ɕ�����
        List = VBA.Split(Lists(l), "/")
        Line = List(0) & vbTab & List(1)
        ListsReduced = ListsReduced & vbTab & Line
    Next l

    With Mp.Pages("page2").lbxHistory
        .Text = Empty
        .Text = ListsReduced
    End With
    
Exit Sub
Err:
    If Err.Number = 53 Then
        Err.Clear
        Open myPath For Output As #2
        Close #2
        Resume Return1
    End If

End Sub

Private Sub UpdateHistoryListBox()
    Dim History As String
    Dim HistotyList As Variant

    myPath = ThisDocument.Path & "\�u������.txt"
    History = LoadHistory
    HistotyList = ConvertTo2DArray(History, vbCr)
    lbxHistory.Clear
    If IsArray(HistotyList) Then lbxHistory.List = HistotyList

End Sub

Private Function LoadHistory()
    Dim tmp As String
    Dim buf As String
    Dim Lists As String
    Const Delimeter As String = vbCr

'   �e�L�X�g�����s���ǂݍ���
    Open myPath For Input As #1
        Do Until EOF(1)
            Line Input #1, tmp
            If Len(tmp) > 0 Then
                buf = buf & tmp & Delimeter
            End If
            If Len(buf) > 3000 Then
                Lists = Lists & buf & Delimeter
                buf = Empty
            End If
        Loop
        If LenB(buf) Then LoadHistory = Lists & buf
    Close #1
End Function

Private Function ConvertTo2DArray(Arr As Variant, Delimeter) As Variant
    Dim Lists As Variant
    Dim List As Variant
    Dim i As Long, j As Long
    Dim c As Long
    Dim What() As Variant
    Dim Replace() As Variant
    
'   �����̋�؂蕶���̓X���b�V��
    Const DictDelimeter As String = "/"

'   ���s�L���ŕ���
    Lists = VBA.Split(Arr, Delimeter)

    If UBound(Lists) = 0 Then Exit Function

    For i = LBound(Lists) To UBound(Lists)
        List = VBA.Split(Lists(i), DictDelimeter)
        On Error Resume Next
        ReDim Preserve What(c) As Variant
        ReDim Preserve Replace(c) As Variant
        What(c) = List(0)
        Replace(c) = List(1)
        On Error GoTo 0
    Next i
   
    Dim Arr2 As Variant
    ReDim Arr2(0 To UBound(Lists) - 1, 1) As Variant
    Dim v As Variant
    On Error Resume Next
    For j = LBound(Lists) To UBound(Lists) - 1
        v = VBA.Split(Lists(j), DictDelimeter)
        Arr2(j, 0) = v(0)
        Arr2(j, 1) = v(1)
    Next j
    On Error GoTo 0

    ConvertTo2DArray = Arr2
Exit Function
Err:
If Err.Number = 9 Then Exit Function
End Function

Private Sub SetRightClickMenu()
    Dim myMenu As Object
    Set myMenu = application.CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
    With myMenu
        With .Controls.Add
            .Caption = "�u�����s"
            .OnAction = "DoReplace"
            .faceId = 125
        End With
        With .Controls.Add
            .Caption = "�u�����ɃR�s�["
            .OnAction = "CopyToReplaceBox"
            .faceId = 607
        End With
    End With
End Sub
'Private Sub lbxHistory_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    If Button = 2 Then myMenu.ShowPopup
'End Sub

Private Sub lbxHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim f As Boolean: f = False
    Dim c As Long, i As Long
    Dim SelItems As String
    Dim myListIndex As Long
    
    If lbxHistory.ListIndex > -1 Then f = True
    If f = False Then Exit Sub
    myListIndex = lbxHistory.ListIndex
'    SelItems = lbxHistory.List(myListIndex, 0) & vbTab & lbxHistory.List(myListIndex, 1)
    For i = 0 To lbxHistory.ListCount - 1
        If lbxHistory.Selected(i) Then
            SelItems = SelItems & lbxHistory.List(i, 0) & _
                        vbTab & lbxHistory.List(i, 1) & vbCr
        End If
    Next
    Debug.Print SelItems
    SelItems = Left$(SelItems, Len(SelItems) - 1)
    txtReplaceWords = txtReplaceWords.Text & SelItems & vbCr
End Sub


