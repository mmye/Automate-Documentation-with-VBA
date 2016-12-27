Attribute VB_Name = "ワードPDF変換"
Option Explicit

Sub Main()
    Dim File As String
    Dim doc As Document
    Dim c As Long: c = 1
    Dim Path As String
    Dim SavePath As String

    Path = SelectFolderDialog("PDFに変換するファイルの場所を選択…")
    SavePath = SelectFolderDialog("変換したPDFを保存する場所を選択…")

    If Path = "" Then Exit Sub
    File = Dir(Path & application.PathSeparator & "*.doc*")
    Do While File <> ""
        If Left$(File, 1) <> "~" Then '隠しファイルを避ける
            Set doc = Documents.Open(Path & "\" & File, ReadOnly:=True) '読み取り専用で開けば安全
            ActiveWindow.Visible = False
            'PDF変換しようとするファイル名を出力
            'Debug.Print c & vbTab & doc.Name
            c = c + 1
            ConvertToPDF doc, SavePath
            CloseDoc doc
        End If
        File = Dir()
    Loop
    
    Set doc = Nothing
    
    'PDFのフォルダを開く
    Shell "explorer " & SavePath, vbNormalFocus
End Sub

Sub ConvertToPDF(doc As Document, SavePath As String)
    Dim myFilePath As String
    Dim myLen As Long
    Dim lDotLocation As Long

    SavePath = SavePath
    lDotLocation = InStrRev(doc.Name, ".")
    myLen = Len(doc.Name)
    
    On Error GoTo Err
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

Private Sub CloseDoc(doc As Document)
    doc.Saved = True
    doc.Close
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
