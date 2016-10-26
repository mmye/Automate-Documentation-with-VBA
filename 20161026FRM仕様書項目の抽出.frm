VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM仕様書項目の抽出 
   Caption         =   "見積仕様書の項目抽出"
   ClientHeight    =   10360
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   12570
   OleObjectBlob   =   "20161026FRM仕様書項目の抽出.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRM仕様書項目の抽出"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'エクセルにのせたときに区切られていない。スプリットするときに正しい区切り文字が指定できていないか、
'スプリットする時点でテキストに正しく区切り文字が入っていないことが原因と思う。
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

Private Sub 見積仕様書の項目名を抽出()
'TODO: オリジナル文書のコピーする。そこで置換し（1行に整形）、抽出する。この一時シートは、処理終了後に保存せず破棄する｡
  
    Dim rng As Word.Range    'Rangeオブジェクト
    Dim copyDoc As Document 'オリジナル文書の使い捨てコピー（オリジナルの内容に変更を加えたくないので）
    Dim NewDoc As Document
    Dim sDocName As String
    Delimeter = "$"
    
    '画面の更新をオフ
    Word.application.ScreenUpdating = False
    On Error GoTo CloseDoc
    Set NewDoc = CopyOriginalTextToTempDoc
    application.WindowState = wdWindowStateMinimize
    Call メーカーの自動マッチ(NewDoc)
    Set rng = NewDoc.Range(0, 0)
'  見出し中の全角文字を半角化（全メーカー共通処理）
    Call 半角全角化(rng)

    '仕様書の種類に合わせて処理を選ぶ
    If optBerents.Value = True Then Call 項番と項目名を1行化(rng)  'Berents用
    Call 品名情報取りだし(rng, NewDoc)
    Call DumpTempDoc(rng, NewDoc)
    If MatchMode Then Exit Sub
    
    '画面の更新をオン
    Word.application.ScreenUpdating = True
    MsgBox "完了しました。", vbInformation, "お知らせ"
Exit Sub

CloseDoc:
    Call DumpTempDoc(rng, NewDoc)
    Dim msg
    msg = "エラー終了しました" & vbCrLf & Err.Number & vbTab & Err.Description
    MsgBox msg
End Sub

Private Function CopyOriginalTextToTempDoc() As Word.Document
    Dim sDocName As String

    '作業用一時ドキュメントをつくり、テキストボックスに入力されたパスの文書を貼り付ける
    If cmbDocumentName.Text = "" Then
        MsgBox "読み込む文書を選択してください", , _
                vbInformation, "お知らせ"
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
   
Private Sub 項番と項目名を1行化(ByRef rng As Range)
'Ｂｅｒｅｎｔｓ仕様書用
'項番と項目名が別の行にある前提
    With rng.Find
        .Text = "ＰＯＳ([ 　．.^tａｰｚa-zＡ-ＺA-Z0-9０-９]{1,})^13" '検索もれはここの文字列を変えて対応する
        .Replacement.Text = "ＰＯＳ\1^t"
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

Private Sub 品名情報取りだし(ByRef rng As Range, ByRef NewDoc As Word.Document)

    Dim eApp As Object
    Dim ewkb As Object
    Dim eWks As Object
    Dim sWhat As String
    Dim Supplier As String
    Dim Lists As Variant    '抽出する文字列
    
    '仕様書の種類に合わせて検索文字列をえらぶ
    sWhat = 見出し文字パターン設定(Supplier)
    Set rng = Nothing
'   文書のRange変数を再設定（同じ変数を続けて使えなくなったので[原因不明]）
    Dim rng2 As Word.Range
    Set rng2 = NewDoc.Range(0, 0)
'   配列の要素数をかぞえる
    Dim cnt As Long
    cnt = GetItemCount(sWhat, rng2)
    
    Set rng2 = Nothing
    Dim rng3 As Word.Range
    Set rng3 = NewDoc.Range(0, 0)
'   品番と項目名の区切りを整理
    Call テキスト正規化(sWhat, rng3, Supplier)
        'デバッグ用正規化済みテキストの確認
    NormalizedLists = NewDoc.Range.Text
    Lists = テキスト構造化(cnt, sWhat, NewDoc)
'    Lists = オプション項目マーキング(cnt, NewDoc)

    '配列がない（＝検索一致結果がない）場合は終了
    If IsArrayEx(Lists) <> 1 Then
        MsgBox "一致する項目がありません", vbInformation, "お知らせ"
        Exit Sub
    End If
        
    Select Case MatchMode
        Case False
            '出力用ワークシートを新規作成する
            Set eApp = CreateObject("Excel.Application")
            eApp.Visible = True
            eApp.application.ScreenUpdating = False
            Set ewkb = eApp.workbooks.Add
            Set eWks = ewkb.sheets(1)
            
            'ワークシートに検索一致結果を貼り付ける
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
        
    Call 列番号登録(eApp, eWks)
'       オプションマーキングの読み取り（思う列にフラグが立てられないので次善策）
    Call オプションフラグ立て(eApp, eWks)
    'データの整理整頓
    Call 整理整頓(eApp, eWks)
    
    Set eApp = Nothing
    Set ewkb = Nothing
    Set eWks = Nothing
    Set r2 = Nothing
    
End Sub
Sub オプションフラグ立て(eApp, eWks)
    Dim i
    Dim LastRow
    Dim buf
    
    With eWks
        LastRow = .Cells(.Rows.Count, 1).End(-4162).Row
        For i = StartRow To LastRow
            If .Cells(i, colRemark).Value <> Empty Then
                buf = Replace(.Cells(i, colRemark), vbCr, "") '改行が混ざってる
                If buf = flagOption Then
                    .Cells(i, colRemark).Value = Empty
                    .Cells(i, colOption).Value = flagOption
                End If
            End If
        Next
    End With
End Sub
Private Function テキスト構造化(cnt, What, ByRef NewDoc As Word.Document) As Variant
    Dim j As Long
    Dim c As Long
    Dim r As Word.Range
    Dim Lists() As Variant
    Dim List As Variant
    Dim myInstr As Long

    ReDim Lists(0 To cnt, 0 To 4) As Variant
''   検索パターンに区切り文字を埋め込む。
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
    
    'ヒットしなくなるまで検索を続ける
    Do While r.Find.Execute = True And r.Text <> ""
        Dim buf
        buf = Replace(r.Text, vbCr, "")

        List = Split(r.Text, Delimeter)
        c = c + 1
        For j = LBound(List, 1) To UBound(List, 1)
            Lists(c - 1, j) = List(j)
        Next j
    Loop
    
    Set r = Nothing
    テキスト構造化 = Lists
    
End Function


Private Sub btnCopyToClipboard_Click()
    Call ExportListsToClipboard
End Sub

'リストボックスからデータをCSVでクリップボードにコピーする
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

    'Clipboardにデータを入れる
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText Lists
        .PutInClipboard
    End With
End Sub

'Private Function オプション項目マーキング(cnt, ByRef NewDoc As Word.Document) As Variant
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
'    'ヒットしなくなるまで検索を続ける
'    Do While r.Find.Execute = True And r.Text <> ""
'        List = r.Text & "$OPTION"
'        c = c + 1
'        For j = LBound(List, 1) To UBound(List, 1)
'            Lists(c - 1, j) = List(j)
'        Next j
'    Loop
'
'    Set r = Nothing
'    テキスト構造化 = Lists
'
'End Function

Sub 列番号登録(ByRef eApp As Object, ByRef wks As Object)

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
Sub 整理整頓(ByRef eApp As Object, ByRef wks As Object)
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

        '見出し設定
        .Cells(1, colPOS).Value = "POS."
        .Cells(1, ColItem).Value = "品名"
        .Cells(1, colEUR).Value = "EUR価格"
        .Cells(1, colRemark).Value = "備考"
        .Cells(1, colOption).Value = "オプション？"
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

Private Function 見出し文字パターン設定(Supplier) As String
    Dim Delimeter2
    Dim DelimeterLEFT As String
    Dim DelimeterRIGHT As String
    Dim myInstr As Long
    Dim buf As String
    
    If optBerents.Value Then Supplier = "Berents"
    If optGroninger.Value Then Supplier = "groninger"
    If optIWK.Value Then Supplier = "IWK"
    
    If chkManualCriteria.Value And txtManualCriteria.Text <> "" Then
        見出し文字パターン設定 = txtManualCriteria.Text
    Else
        Select Case Supplier
            Case "Berents"
                buf = "(^13ＰＯＳ[　 ^t]{1,})"
            Case "groninger"
                buf = "(^13[0-9]{1,4})[　 ^t]{1,}"
            Case "IWK"
                buf = "([0-9]{5,6})[ 　^t]{1,}"
        End Select
    End If

'   検索パターンに区切り文字を埋め込む。
    myInstr = InStr(buf, " ")
    DelimeterLEFT = Left$(buf, myInstr)
    DelimeterRIGHT = Right$(buf, Len(buf) - myInstr)
    見出し文字パターン設定 = DelimeterLEFT & "$" & DelimeterRIGHT
End Function

Sub テキスト正規化(What As String, ByRef rng As Word.Range, Supplier)

    Select Case Supplier
        Case "Berents"
            If chkEN Then テキスト正規化_Berents_EN What, rng _
            Else: テキスト正規化_Berents_JA What, rng
        Case "groninger"
            If chkEN Then テキスト正規化_groninger_EN What, rng _
            Else: テキスト正規化_groninger_JA What, rng
        Case "IWK"
            If chkEN Then テキスト正規化_IWK_EN What, rng _
            Else: テキスト正規化_IWK_JA What, rng
    End Select
End Sub
Private Sub テキスト正規化_IWK_EN(What As String, ByRef rng As Word.Range)
    Call 空白文字区切り整理_IWK(What, rng)
    Call Tab削除(rng)
    Call Remark区切り整理(rng)
    Call EUR区切り整理(rng)
    Call オプション括弧削除とフラグ立て(rng)
    Call 余分な区切り文字削除(rng)
    Call 余分なカンマ削除(rng)
End Sub

Private Sub テキスト正規化_IWK_JA(What As String, ByRef rng As Word.Range)
    Call 全角スペース区切り整理(rng)
    Call 空白文字区切り整理_IWK(What, rng)
    Call Tab削除(rng)
    Call 余分な区切り文字削除(rng)
    Call 余分なカンマ削除(rng)
End Sub

Private Sub テキスト正規化_groninger_EN(What As String, ByRef rng As Word.Range)
    Call 空白文字区切り整理_groninger(What, rng)
    Call Tab削除(rng)
    Call Remark区切り整理(rng)
    Call EUR区切り整理(rng)
    Call オプション括弧削除とフラグ立て(rng)
    Call 余分な区切り文字削除(rng)
    Call 余分なカンマ削除(rng)
End Sub

Private Sub テキスト正規化_groninger_JA(What As String, ByRef rng As Word.Range)
    Call 全角スペース区切り整理_groninger(rng)
    Call 空白文字区切り整理(What, rng)
    Call Tab削除(rng)
    Call 余分な区切り文字削除(rng)
    Call 余分なカンマ削除(rng)
End Sub

Private Sub テキスト正規化_Berents_EN(What As String, ByRef rng As Word.Range)
    Call 空白文字区切り整理_Berents(What, rng)
    Call Tab削除(rng)
    Call Remark区切り整理(rng)
    Call EUR区切り整理(rng)
    Call オプション括弧削除とフラグ立て(rng)
    Call 余分な区切り文字削除(rng)
    Call 余分なカンマ削除(rng)
End Sub
Private Sub テキスト正規化_Berents_JA(What As String, ByRef rng As Word.Range)
    Call 空白文字区切り整理_Berents(What, rng)
    Call Tab削除(rng)
    Call Remark区切り整理(rng)
    Call EUR区切り整理(rng)
    Call オプション括弧削除とフラグ立て(rng)
    Call 余分な区切り文字削除(rng)
    Call 余分なカンマ削除(rng)
End Sub
Sub 空白文字区切り整理_IWK(What As String, ByRef rng As Word.Range)
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
Sub 空白文字区切り整理_groninger(What As String, ByRef rng As Word.Range)
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
Sub 空白文字区切り整理_Berents(What As String, ByRef rng As Word.Range)
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
Sub Tab削除(ByRef rng As Word.Range)
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
Sub EUR区切り整理(ByRef rng As Word.Range)
'品名の行にEUR価格が含まれる場合に区切り文字を挿入する
'後で価格も配列に入れるため。

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
Sub Remark区切り整理(ByRef rng As Word.Range)
'品名の行にEUR価格が含まれる場合に区切り文字を挿入する
'後で価格も配列に入れるため。

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

Sub 全角スペース区切り整理(ByRef rng As Word.Range)
'日本語のIWK仕様書で品名の後に続くスペースを区切る

    Dim sReplace As String
    Dim What As String
    
    What = "([ 　^t]{2,})"
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

Sub オプション括弧削除とフラグ立て(ByRef rng As Word.Range)
'１：EUR価格のテキスト範囲に含まれる括弧（）を除去
'２：カンマとドットを除去

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
Sub 余分な区切り文字削除(ByRef rng As Word.Range)
'よくわからなくなって区切り文字がたくさんできてしまうので1つに減らす
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

Sub 余分なカンマ削除(ByRef rng As Word.Range)
'よくわからなくなって区切り文字がたくさんできてしまうので1つに減らす
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
' 機能   : 引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数   : varArray  配列
' 戻り値 : 判定結果（1:配列/0:空の配列/-1:配列じゃない）
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
            buf = Replace(buf, ".", "") 'コンマ除去
            buf = Replace(buf, ",", "") 'ドット除去
            Lists(i, 1) = buf
        End If
    Next
    
    r.Value = Lists
    On Error GoTo 0
End Sub

Private Sub 半角全角化(rng)
    '全角英数字を半角英数字へ一括変換
    Dim Range As Word.Range
    Set Range = rng
    
    With Range.Find
        .Text = "[０-９]{5,6}"
        .MatchWildcards = True
        Do While .Execute = True
          Range.CharacterWidth = wdWidthHalfWidth
          Range.Collapse wdCollapseEnd
        Loop
    End With
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
        Call 見積仕様書の項目名を抽出
    Else
        MsgBox "メーカーを選択してください", vbInformation, "お知らせ"
    End If
End Sub
Private Sub btnMatch_Click()
    Dim flag As Boolean
    Dim c As Control

    Call 最近使ったファイル名をレジストリに登録
    MatchMode = True
    For Each c In frmManufacturer.Controls
        If TypeName(c) = "OptionButton" Then _
        If c.Value Then flag = True
    Next c
    
    If flag Then
        Call 見積仕様書の項目名を抽出
    Else
        MsgBox "メーカーを選択してください", vbInformation, "お知らせ"
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

Private Sub txtDocumentName_Change()

End Sub

Private Sub optIWK_Click()

End Sub

Private Sub UserForm_Initialize()
    With lbMatch
        .ColumnCount = 5
        .ColumnWidths = "50;200;50;20;20"
    End With
    chkJA.Value = True
    Call 最近開いたファイル名をレジストリから読み出す
End Sub

Private Sub CommandButton1_Click()
    Call リスト内の余分なテキストを除去
End Sub
Private Sub リスト内の余分なテキストを除去()
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
        Case Is > 0 '除去する文字が複数ある場合
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
        
    Case Else '除去する文字がひとつだけの場合
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

Private Sub メーカーの自動マッチ(Doc)
    Dim i
    Dim Supplier
    Dim r
    Dim f
    Dim Name
    Supplier = Array("IWK", "groninger", "Berents", "Becomix")
    
    Set r = Doc.Range(0, 0)
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
    Set r2 = Doc.Range(0, 0)
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
    Set r3 = Doc.Range(0, 0)
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
    Set r4 = Doc.Range(0, 0)
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
End Sub

'Private Sub 英数字_半→全()
'    '半角英数字を全角英数字へ一括変換
'    Dim myRange As Range
'    Dim blnFound As Boolean
'    application.ScreenUpdating = False
'    Set myRange = ActiveDocument.Range(0, 0)
'    With myRange.Find
'        .Text = "[0-9A-Za-z]{1,}"  '対象の設定
'        .MatchWildcards = True
'        Do While .Execute = True
'          blnFound = True
'          myRange.HighlightColorIndex = wdTurquoise  '色の設定
'          myRange.CharacterWidth = wdWidthFullWidth
'          myRange.Collapse wdCollapseEnd
'        Loop
'    End With
'    Set myRange = Nothing
'
'    If blnFound = True Then
'      MsgBox "半角英数字を全角に変換しました。"  'メッセージ
'    End If
'    application.ScreenUpdating = True
'End Sub
'
'Private Sub 英数字_全→半()
'    '全角英数字を半角英数字へ一括変換
'    Dim myRange As Range
'    Dim blnFound As Boolean
'    application.ScreenUpdating = False
'    Set myRange = ActiveDocument.Range(0, 0)
'    With myRange.Find
'        .Text = "[０-９Ａ-Ｚａ-ｚ]{1,}"  '対象の設定
'        .MatchWildcards = True
'        Do While .Execute = True
'          blnFound = True
'          myRange.HighlightColorIndex = wdBrightGreen  '色の設定
'          myRange.CharacterWidth = wdWidthHalfWidth
'          myRange.Collapse wdCollapseEnd
'        Loop
'    End With
'    Set myRange = Nothing
'
'    If blnFound = True Then
'      MsgBox "全角英数字を半角に変換しました。"  'メッセージ
'    End If
'    application.ScreenUpdating = True
'End Sub
Private Sub spinUpDownItem_SpinDown()
    Dim strListItemDown() As String
    Dim strListItemSelected() As String
    Dim i As Long
    Dim blIsSelected As Boolean
    Dim r, r2
    Dim ColCnt
    
    With lbMatch
        For i = 0 To .ListCount - 1
            If .Selected(i) Then blIsSelected = True
        Next i
        If Not blIsSelected Then Exit Sub        'リストが選択されていなければ終了
    End With
    
    ColCnt = lbMatch.ColumnCount - 1
    ReDim strListItemDown(ColCnt) As String
    
    With lbMatch
        If .ListIndex + 1 < .ListCount Then
            For r = 0 To ColCnt
                strListItemDown(r) = .List(.ListIndex + 1, r)
            Next r
        Else
            Exit Sub    'これ以上下がなければ終了する
        End If
        
        ReDim strListItemSelected(ColCnt) As String
        For r2 = 0 To ColCnt
            strListItemSelected(r2) = .List(.ListIndex, r2)
            .List(.ListIndex, r2) = strListItemDown(r2)
            .List(.ListIndex + 1, r2) = strListItemSelected(r2)
        Next r2
    End With
    lbMatch.Selected(lbMatch.ListIndex + 1) = True
    
End Sub

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
        If Not blIsSelected Then Exit Sub        'リストが選択されていなければ終了
    End With
    
    ColCnt = lbMatch.ColumnCount - 1
    ReDim strListItemUp(ColCnt) As String
    
    With lbMatch
        If .ListIndex > 0 Then
            For r = 0 To ColCnt
            strListItemUp(r) = .List(.ListIndex - 1, r)
            Next r
        Else
            Exit Sub    'これ以上、上がなければ終了する
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

Sub 最近使ったファイル名をレジストリに登録()
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

Sub 最近開いたファイル名をレジストリから読み出す()
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
