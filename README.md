## はじめに

このユーティリティは、Wordファイル中の文字列を置換します。
仕事で使っています。



## 使い方

1. サンプルファイルにあるボタンをクリックして起動
2. 検索語と置換語をタブ区切りで入力。改行して複数の組み合わせもいけます。
3. 「置換」ボタンを押す
4. 置換されます。


## 特徴

- 一度にたくさんの語句を置換できる

- 複数のファイルを同時に置換できる（Word標準では現在選択した文書のみ）

- 正規表現が使える（Word標準ではワイルドカードのみ）

- 置換の履歴が残る（下記のような。「履歴」タブを選択）    


## その他

- 置換したテキストに蛍光ペンを引いたり、赤字にしたりできる。変更箇所がわかりやすい。
- 置換履歴はテキストファイルで保存。パスはデスクトップです。「履歴」ボタンを押すと既定のテキストエディタで置換履歴が開きます。

   

## コード

ユーザーフォーム制御、置換処理ブロックなどなど全部で1000行ほどのプログラムです。
その中から2点を下記にのせてみました。

### 1. 置換処理

文書内のパラグラフリストをループし、マッチ判定してTrueで置換します。
引数に検索語・置換語のString型変数とWord文書オブジェクト変数を取ります。

```VB
Private Sub Replaces(WordLists, doc As Word.Document)
'   引数：WordList=検索語・置換語,doc=ワード文書

    Dim i As Long
    Dim sWhatReplace As Variant
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

    Dim WhatStr As String, ReplaceStr As String
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
```



### 2. 配列変換関数

テキスト変数を指定の区切り文字で二次元配列化して返します。

```VB
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
```



## お読みいただき、ありがとうございました。


