Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd閲覧_Click()
    Dim strFname    As String
    Dim strPdfLink  As String
    Dim strJyouken  As String
    
On Error GoTo cmd閲覧_Click_ERR
    
    strFname = "F基本情報sub"
    Me.cmd閲覧.HyperlinkAddress = ""
    srtPdflink = ""
    
    If Me.txtPDFリンク = "" Or IsNull(Me.txtPDFリンク) = True Then
        MsgBox "PDFファイルが登録されていません。", vbExclamation + vbOKOnly, "確認"
        Exit Sub
    Else
        strPdfLink = Me.txtPDFリンク
    End If
'    MsgBox strPdfLink

    If strPdfLink = "*" Then   'PDFファイルパス名が"*"のとき
        Beep
        MsgBox "PDFの登録がありません！", vbExclamation + vbOKOnly, "確認"
        Me.cmd閲覧.HyperlinkAddress = ""  'PDFのリンクを切る。
    ElseIf Dir(strPdfLink) = "" Or IsNull(strPdfLink) = True Then 'PDFファイル（パス名に一致するもの）が存在しないとき
        Beep
        MsgBox "PDFファイルが見つかりません！", vbExclamation + vbOKOnly, "確認"
        Me.cmd閲覧.HyperlinkAddress = ""  'PDFのリンクを切る。
    Else
        Me.cmd閲覧.HyperlinkAddress = strPdfLink  'PDFのリンクを設定して開く。
    End If

cmd閲覧_Click_EXIT:
    Exit Sub
    
cmd閲覧_Click_ERR:
    MsgBox Err.Description
'    Me.cmd閲覧.HyperlinkAddress = ""  'PDFのリンクを切る。
    Resume cmd閲覧_Click_EXIT

End Sub
'
'Private Sub cmd編集_Click()
'
'    Dim dataArgs    As String
'    Dim strBangou  As String
'
'    '抽出条件保存
'    strBangou = Me.txt番号
'    strNichiji = Me.登録日時
'
'    dataArgs = strBangou & "," & strNichiji
'    MsgBox dataArgs
'
'    flgForm = 0
'    DoCmd.OpenForm "F新規修正", , , , , , dataArgs
'
'    If flgHaita = 1 Then
'        DoCmd.Close acForm, "F新規修正"
'        DoCmd.OpenForm "Fメイン"
'        Forms!Fメイン.情報Sub.Requery
'    Else
'        DoCmd.Close acForm, "Fメイン"
'    End If
'
'End Sub