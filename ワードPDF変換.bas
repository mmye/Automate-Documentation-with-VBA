Attribute VB_Name = "���[�hPDF�ϊ�"
Option Explicit

Sub Main()
    Dim File As String
    Dim doc As Document
    Dim c As Long: c = 1
    Dim Path As String
    Dim SavePath As String

    Path = SelectFolderDialog("PDF�ɕϊ�����t�@�C���̏ꏊ��I���c")
    SavePath = SelectFolderDialog("�ϊ�����PDF��ۑ�����ꏊ��I���c")

    If Path = "" Then Exit Sub
    File = Dir(Path & application.PathSeparator & "*.doc*")
    Do While File <> ""
        If Left$(File, 1) <> "~" Then '�B���t�@�C���������
            Set doc = Documents.Open(Path & "\" & File, ReadOnly:=True) '�ǂݎ���p�ŊJ���Έ��S
            ActiveWindow.Visible = False
            'PDF�ϊ����悤�Ƃ���t�@�C�������o��
            'Debug.Print c & vbTab & doc.Name
            c = c + 1
            ConvertToPDF doc, SavePath
            CloseDoc doc
        End If
        File = Dir()
    Loop
    
    Set doc = Nothing
    
    'PDF�̃t�H���_���J��
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
    'PDF�ɕϊ����邽�߂ɂ́A��������x�ۑ�����K�v������
    If Err.Number = 5 Then MsgBox "��x������ۑ����Ă���ēx���s�����PDF���ۑ��ł��܂�"
    
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
