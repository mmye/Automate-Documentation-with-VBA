Attribute VB_Name = "Module2"
Option Explicit

Dim FilePath As String
Dim SavePath As String

Sub Main()
'   �����f�B���N�g���ƕۑ��f�B���N�g�����w�肷��
    FilePath = SelectFolderDialog("PDF�ɕϊ�����t�@�C���̏ꏊ��I���c")
    SavePath = SelectFolderDialog("�ϊ�����PDF��ۑ�����ꏊ��I���c")
    If FilePath = "" Or SavePath = "" Then Exit Sub

'   �I�������t�H���_���̃t�@�C����T���ď��Ԃɕϊ���������
    FindWordDocs FilePath

    'PDF�̃t�H���_���J��
    Shell "explorer " & SavePath, vbNormalFocus
End Sub

Private Sub FindWordDocs(FilePath)
    Dim doc As Document
    Dim File As String
    Dim c As Long: c = 1

    File = Dir(FilePath & application.PathSeparator & "*.doc*")
    Do While File <> ""
        If Left$(File, 1) <> "~" Then '�B���t�@�C���������
            Set doc = Documents.Open(FilePath & "\" & File, ReadOnly:=True) '�ǂݎ���p�ŊJ���Έ��S
            ActiveWindow.Visible = False
            'PDF�ϊ����悤�Ƃ���t�@�C�������o��
            'Debug.Print c & vbTab & doc.Name
            c = c + 1
            ConvertToPDF doc, SavePath
            CloseDoc doc
        End If
        File = Dir()
    Loop
    
'   �ċA�I�ɖ{�v���V�[�W�������s���A�q�f�B���N�g�����T��
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
    'PDF�ɕϊ����邽�߂ɂ́A��������x�ۑ�����K�v������
    If Err.Number = 5 Then MsgBox "��x������ۑ����Ă���ēx���s�����PDF���ۑ��ł��܂�"
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
'   ��������邾��
    doc.Saved = True
    doc.Close
End Sub
