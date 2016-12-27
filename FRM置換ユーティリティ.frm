VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM置換ユーティリティ 
   Caption         =   "置換++"
   ClientHeight    =   5880
   ClientLeft      =   110
   ClientTop       =   470
   ClientWidth     =   7050
   OleObjectBlob   =   "FRM置換ユーティリティ.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRM置換ユーティリティ"
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

'カラーピッカーダイアログのコード+++++++++++++++++++++++++++++++++++++++++++++++++
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
 
    With udtChooseColor 'ダイアログの設定
      .lStructSize = Len(udtChooseColor)
      .lpCustColors = String$(64, Chr$(0))
      .flags = CC_RGBINIT + CC_LFULLOPEN
      .rgbResult = lngDefColor
    End With
    lngRet = ChooseColor(udtChooseColor) 'ダイアログを表示
    
    If lngRet <> 0 Then 'ダイアログからの戻り値をチェック
        If udtChooseColor.rgbResult > RGB(255, 255, 255) Then
          GetColorDlg = -2 'エラーの場合
        Else
          GetColorDlg = udtChooseColor.rgbResult '戻り値にRGB値を代入
        End If
    Else
      GetColorDlg = -1 'キャンセルされた場合
    End If
 
End Function

'Private Sub btnMinimize_Click()
'    If Not IsMinimized Then
'        Me.Width = 383.5
'        Me.Height = 200
'    Else
'        Me.Width = 383.5
'        Me.Height = 329.5
'    End If
'End Sub

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
    If ck正規表現オプション.Value Then _
    ck正規表現オプション.Value = Not ckUseWildCards.Value
    Select Case Mp.Value
        Case 0: txtReplaceWords.SetFocus
        Case 1:
    End Select
End Sub

Private Sub ck正規表現オプション_Click()
    If ckUseWildCards.Value Then _
    ckUseWildCards.Value = Not ck正規表現オプション.Value
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
    myDict = ThisDocument.Path & "\置換辞書.txt"
    myDict = """" & myDict & """" 'Wscript.Shellに渡す引数に含まれる￥をエスケープ
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
        SetReplaceNotFinished '「入力語句未置換フラグ」を立てる
    End If
End Sub

'********************************************************
'レジストリ関係
'テキストボックスの入力をレジストリにバックアップする。
'重いファイルを同時編集するときにWordはよく落ちるので。
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
'   レジストリの置換完了フラグを立てる。
'  レジストリにバックアップしてある置換語句を消す
    SaveSetting "MyMacro", "BulkReplace", "IsReplaced", True
    SaveSetting "MyMacro", "BulkReplace", "InputWords", ""
End Sub
Private Function CheckReplaceStatus() As Boolean
    On Error GoTo Err
    CheckReplaceStatus = GetSetting("MyMacro", "BulkReplace", "IsReplaced")
    On Error GoTo 0
Err:
    'レジストリがない場合は作る'
    SaveSetting "MyMacro", "BulkReplace", "Isreplaced", False
End Function


Private Sub UserForm_Initialize()

'   イニシャライズ時にオプションボタンのクリックイベントが発生するのを避けるためのイベント制御変数
    mCancelEvent = True
    
'   置換辞書の場所を指定
'   保存していない文書で実行すると下記はエラーになる。
'  Err=5で判定して、デスクトップのパスを取得し、テキストファイルを作成する
    On Error Resume Next
    myPath = ThisDocument.Path & "\置換辞書.txt"
     If Err.Number = 5 Then myPath = GetDesktopPath: Err.Number = 0
    On Error GoTo 0

'   入力した語句の再読み込みを判定（強制終了などによる入力途中のデータを復元する）
    Dim Replaced As Boolean
    Replaced = CheckReplaceStatus
    If Not Replaced Then txtReplaceWords.Text = GetSavedStr

'   置換履歴リストボックスの設定
    lbxHistory.ColumnCount = 2
'    lbxHistory.MultiSelect = fmMultiSelectExtended

'   コントロールON/OFFの初期設定
    Me.ckReplaceAllDocs = False
    Me.ckMatchCase = False
    Me.ckUseWildCards = False
    
'   ハイライト色の初期設定
    lblSelectColor.BackColor = vbGreen
    lblSelectColor.Tag = 4
    optUseHighlight.Value = True

    Call 置換履歴をテキストボックスに表示する
    
'   テキストボックスにフォーカス
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
Private Function GetDesktopPath()
    Dim Path As String, WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    GetDesktopPath = WSH.SpecialFolders("Desktop") & "\"
    Set WSH = Nothing
End Function
End Function

Private Sub cmdExecute_Click()
    Dim WordLists As Variant
    WordLists = GetWhatReplace
    If Not IsArray(WordLists) Then Exit Sub

    '置換範囲、方法によって4通りに条件分岐
    '範囲：アクティブドキュメント｜全ドキュメント；　置換方法：標準置換｜正規表現
    Select Case ckReplaceAllDocs.Value
        Case False
            If ck正規表現オプション.Value Then
                RegexReplace WordLists, ActiveDocument
            Else
                Replaces WordLists, ActiveDocument
            End If
        Case True
            Dim doc As Document
            If ck正規表現オプション.Value Then
                For Each doc In Documents
                    RegexReplace WordLists, doc
                Next doc
            Else
                For Each doc In Documents
                    Replaces WordLists, doc
                Next doc
            End If
    End Select

'   置換結果を表示(マッチがあったかなかったか知らせる)
    ShowCompletionMsg (mHasMatch)
    mHasMatch = False
    
'   検索・置換後を履歴に保存する
    置換辞書に登録 WordLists
'   履歴を使って置換する場合は、完了したらリストボックスの選択を解除する
    Dim i As Long
    If mUseHistory Then
        For i = 0 To lbxHistory.ListCount - 1
            If lbxHistory.Selected(i) Then lbxHistory.Selected(i) = False
        Next i
        mUseHistory = False
    End If
    SetReplaceDone
    Call フォーカス制御
End Sub

Private Sub Replaces(WordLists, doc As Word.Document)
'   引数：WordList=検索語・置換語,doc=ワード文書

    Dim rng As Range
    Set rng = doc.Range(0, 0)

    '書式設定
    '蛍光ペン色。色の定数が入ったラベルを参照する。
    Options.DefaultHighlightColorIndex = Me.lblSelectColor.Tag
    With rng.Find
        'ワイルドカード
        If ckUseWildCards.Value Then _
            .MatchWildcards = True Else: .MatchWildcards = False
        If ckMatchCase.Value Then _
            .MatchCase = True Else: .MatchCase = False
        '置換フォント色
        If optChangeFontColor.Value Then .Replacement.Font.Color = wdColorRed
        If optUseHighlight.Value Then .Replacement.Highlight = True
    End With

    Dim sWhatReplace As Variant
    Dim WhatStr As String, ReplaceStr As String
    Dim i As Long
    
    For i = LBound(WordLists) To UBound(WordLists)
        sWhatReplace = VBA.Split(WordLists(i), vbTab)
        WhatStr = sWhatReplace(0)
        ReplaceStr = sWhatReplace(1)

        '空白と置換（検索語の削除）＝0とそれ以外
        Select Case Len(ReplaceStr)
            Case Is > 0
                With rng.Find
                    .Text = WhatStr
                    .Replacement.Text = ReplaceStr
                    If .Execute = True Then mHasMatch = True
'                   正規表現オプションがオフのとき、入力に正規表現を使用したらエラーになったため一時的にエラー無効化
                    On Error Resume Next
                    .Execute Replace:=wdReplaceAll
                    On Error GoTo 0
                End With
            Case 0
                Call ReplaceWithEmpty(doc, WhatStr, mHasMatch)
        End Select
    Next i

    'テキストボックスの置換
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
    msg = "エラー番号：" & Err.Number & vbCrLf & _
          "エラー内容：" & Err.Description
    MsgBox msg, vbCritical, "エラー終了"
    Set rng = Nothing
    Set doc = Nothing
End Sub

Private Function GetWhatReplace() As Variant
'ユーザーフォームから検索語と置換語を取得する
'新規入力ペインと履歴ペインのどちらが選択されているかに応じて条件分岐させる
'テキストボックスから置換語句を取得して、改行区切りで配列にする

    Dim buf As String
    Dim intPage As Long
    intPage = Mp.Value
    
    'intPage=Mp.Value：新規入力ペイン=0、履歴ペイン=1
    Select Case intPage
        Case 0
            buf = txtReplaceWords.Text
        Case 1
            buf = 履歴の選択を取得
            mUseHistory = True
    End Select

    Dim List As String
    List = RemoveEmptyRows(buf)
    GetWhatReplace = VBA.Split(List, vbCrLf)
End Function

Private Function 履歴の選択を取得() As String
    'リストボックスが選択されていなければ終了させる
    If lbxHistory.ListIndex = 0 Then Exit Function
    
    Dim SelItems As String
    Dim IsAnySelected As Boolean
    Dim i As Long
    For i = 0 To lbxHistory.ListCount - 1
        If lbxHistory.Selected(i) Then
            IsAnySelected = True
            SelItems = SelItems & lbxHistory.List(i, 0) & _
                        vbTab & lbxHistory.List(i, 1) & vbCrLf
        End If
    Next
    
    If SelItems <> "" Then
        SelItems = Left$(SelItems, Len(SelItems) - 1)
        履歴の選択を取得 = SelItems
    Else
        Exit Function
    End If
End Function
Private Sub フォーカス制御()
    Dim ind As Long: ind = Mp.Value
    Select Case ind
        Case 0: txtReplaceWords.SetFocus
        Case 1: lbxHistory.SetFocus
    End Select
End Sub

Private Function RemoveEmptyRows(targetStr As String)
    'テキストボックスの空行を削除する。
    '空行があると置換処理の途中でバグる。
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
'           rngから改行を除く（これをしないと改行が消える）
            rng.MoveEnd unit:=wdCharacter, Count:=-1

            Dim i As Long
            For i = LBound(WordLists) To UBound(WordLists)
'               置換リストに空行があった場合、無視する
                On Error Resume Next
                Dim sWhatReplace As Variant
                sWhatReplace = VBA.Split(WordLists(i), vbTab)
                On Error GoTo 0
'               置換語句リストが空行だったら飛ばす
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

    'テキストボックスの置換
    Dim sp As Shape
    With doc
        For Each sp In .Shapes
            If sp.Type = msoTextBox Then
                sp.Select
                 For i = LBound(WordLists) To UBound(WordLists)
'                   置換リストに空行があった場合、無視する
                    On Error Resume Next
                    sWhatReplace = VBA.Split(WordLists(i), vbTab)
                    On Error GoTo 0
'                   置換語句リストが空行だったら飛ばす
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
    msg = "エラー番号：" & Err.Number & vbCrLf & _
          "エラー内容：" & Err.Description
    MsgBox msg, vbCritical, "エラー終了"
    Set rng = Nothing
    Set doc = Nothing

End Sub

Private Sub ShowCompletionMsg(HasAnyMatch)
    Select Case HasAnyMatch
        Case False
            MsgBox "一致する項目はありません。", vbInformation
            Exit Sub
        Case True
            MsgBox "完了しました。", vbOKOnly + vbInformation, "完了"
    End Select
End Sub

Private Sub ワイルドカードとフォント色の設定(rng, HasSet)
    HasSet = True
    With rng.Find
        Select Case ckUseWildCards
            Case True: .MatchWildcards = True
            Case False: .MatchWildcards = False
        End Select
            
    '   置換後のフォント色
        If optChangeFontColor.Value Then .Replacement.Font.Color = wdColorRed
        If optUseHighlight.Value Then .Replacement.Highlight = True
    End With
End Sub

Private Sub ReplaceWithEmpty(doc As Word.Document, What As Variant, mHasMatch As Boolean)
    
    Dim ReplaceStr As String: ReplaceStr = ""
    
    '本文を置換
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    For Each para In doc.Paragraphs
        Set rng = para.Range
'       改行文字を除いた部分を参照する（とても大切）
        rng.MoveEnd unit:=wdCharacter, Count:=-1
        targetStr = rng.Text
        If targetStr = "" Then GoTo NextPara
        Dim Reg As String
        Reg = What
        If RegularExpressions.RegexTest(targetStr, Reg) Then mHasMatch = True
        Dim ret As String
        ret = RegularExpressions.RegexReplace(targetStr, Reg, ReplaceStr)
        rng.Text = ret
NextPara:
    Next para

    'テキストボックスを置換
    Dim sp As Word.Shape
    For Each sp In doc.Shapes
        If sp.Type = msoTextBox Then
            sp.Select
            Selection.Find.ClearFormatting
            Selection.WholeStory
                
            Dim targetStr As String
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

Private Sub 置換辞書に登録(WhatReplace As Variant)

    If Not IsArray(WhatReplace) Then Exit Sub
    
    Dim Lines As Variant
    Lines = LoadHistory
    
    Dim v As Variant
    v = VBA.Split(Lines, vbCr)

    Dim HasContent As Boolean
    Select Case UBound(v)
        Case Is > 0: HasContent = True
        Case Is = -1: HasContent = False
    End Select

    Dim i As Long
    Dim wordColl As New Collection
    '辞書の中身をCollectionに書き込む
    '重複があるとエラーになるのでResumeにしている
    On Error Resume Next
    For i = LBound(v) To UBound(v)
        wordColl.Add v(i)
    Next i
    On Error GoTo 0

'   検索・置換後が新しい語であれば履歴テキストファイルに書きこむ
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

Private Sub 置換履歴をテキストボックスに表示する()
 
    Dim Lists As Variant
    Lists = LoadHistory
    Lists = VBA.Split(Lists, vbCr)
 
    On Error GoTo Err
Return1:
'   配列要素を決定するために履歴の項数をしらべる
    Dim HasMultipleItems
    Select Case UBound(Lists)
        Case Is > 1: HasMultipleItems = True
        Case Is <= 1: HasMultipleItems = False
    End Select

'   最大要素数-1するのは、空要素を削除するため
    Select Case HasMultipleItems
        Case True
            ReDim Preserve Lists(LBound(Lists) To UBound(Lists) - 1) As String
        Case False
            ReDim Preserve Lists(UBound(Lists) - 1) As String
    End Select

    Dim i As Long
    Dim List As Variant
    Dim WhatReplace As String
    Dim WhatReplaces As String
    Const DictDelimeter As String = "/"
    
    On Error Resume Next ' 検索語と置換語が両方ないとエラーになる
    For i = UBound(Lists) To 0 Step -1
'       後に重複を確かめるため、辞書の１行をタブ区切りで二次元配列に分ける
        List = VBA.Split(Lists(i), DictDelimeter)
        WhatReplace = List(0) & vbTab & List(1)
        WhatReplaces = WhatReplaces & vbTab & WhatReplace
    Next i

    With Mp.Pages("page2").lbxHistory
        .Text = Empty
        .Text = WhatReplaces
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

    myPath = ThisDocument.Path & "\置換辞書.txt"
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

'   テキストから一行ずつ読み込む
    On Error GoTo Err
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
Exit Function
Err:

End Function

Private Function ConvertTo2DArray(Arr As Variant, Delimeter) As Variant
    Dim Lists As Variant
    Dim List As Variant
    Dim i As Long, j As Long
    Dim c As Long
    Dim What() As Variant
    Dim Replace() As Variant
    
    If Len(Arr) = 0 Then Exit Function
'   辞書の区切り文字はスラッシュ
    Const DictDelimeter As String = "/"

'   改行記号で分割
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
    ReDim Arr2(0 To UBound(Lists) - 1, 0 To 1) As Variant
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
