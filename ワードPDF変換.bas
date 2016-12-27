Attribute VB_Name = "Module2"
Option Explicit

Dim FilePath As String
Dim SavePath As String

Sub Main()
'   検索ディレクトリと保存ディレクトリを指定する
    FilePath = SelectFolderDialog("PDFに変換するファイルの場所を選択…")
    SavePath = SelectFolderDialog("変換したPDFを保存する場所を選択…")
    If FilePath = "" Or SavePath = "" Then Exit Sub

'   選択したフォルダ内のファイルを探して順番に変換処理する
    FindWordDocs FilePath

    'PDFのフォルダを開く
    Shell "explorer " & SavePath, vbNormalFocus
End Sub

Private Sub FindWordDocs(FilePath)
    Dim doc As Document
    Dim File As String
    Dim c As Long: c = 1

    File = Dir(FilePath & application.PathSeparator & "*.doc*")
    Do While File <> ""
        If Left$(File, 1) <> "~" Then '隠しファイルを避ける
            Set doc = Documents.Open(FilePath & "\" & File, ReadOnly:=True) '読み取り専用で開けば安全
            ActiveWindow.Visible = False
            'PDF変換しようとするファイル名を出力
            'Debug.Print c & vbTab & doc.Name
            c = c + 1
            ConvertToPDF doc, SavePath
            CloseDoc doc
        End If
        File = Dir()
    Loop
    
'   再帰的に本プロシージャを実行し、子ディレクトリも探る
    With CreateObject("Scripting.FileSystemObject")
        Dim f As Object
        For Each f In .GetFolder(FilePath).SubFolders
            FindWordDocs (f.Path)
        Next f
    End With
    
    Set doc = Nothing
    
End Sub

Sub ConvertToPDF(doc As Document, SavePath As String)
    Dim myLen As Long
    Dim lDotLocation As Long
    lDotLocation = InStrRev(doc.Name, ".")
    myLen = Len(doc.Name)
    
    On Error GoTo Err
    Dim myFilePath As String
    myFilePath = SavePath & "\" & _
    Left$(doc.Name, myLen - (myLen - lDotLocation + 1)) & ".pdf"
    doc.ExportAsFixedFormat _
    exportformat:=wdExportFormatPDF, _
    outputfilename:=myFilePath

Exit Sub
Err:
    'PDFに変換するためには、文書を一度保存する必要がある
    If Err.Number = 5 Then MsgBox "一度文書を保存してから再度実行するとPDFが保存できます"
End Sub

Private Function SelectFolderDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String) As String
'title - the title for the dialog
'buttonName - name of the action button. Warning does not always work
'initialFileName - initial folder path for the dialog e.g. C:\
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = application.FileDialog(msoFileDialogFolderPicker)
    'Properties
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        SelectFolderDialog = fDialog.SelectedItems(1)
    End If
End Function

Private Sub CloseDoc(doc As Document)
'   文書を閉じるだけ
    doc.Saved = True
    doc.Close
End Sub
