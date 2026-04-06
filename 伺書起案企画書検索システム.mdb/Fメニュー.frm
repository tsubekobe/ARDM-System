Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strSQL  As String
Dim intRtn  As Integer   ' ★ 追加（旧コードで未宣言だった）

' ================================================================
' フォームオープン：ボタン表示制御
' ================================================================
Private Sub Form_Open(Cancel As Integer)

    ' ログイン・ロック管理は flgHyouji>=1 または flgSYS=1 のとき表示
    Dim blnAdmin As Boolean
    blnAdmin = (flgHyouji >= 1) Or (flgSYS = 1)

    Me.cmdログイン管理.Visible = blnAdmin
    Me.cmdログイン管理.Enabled = blnAdmin
    Me.cmdロック管理.Visible = blnAdmin
    Me.cmdロック管理.Enabled = blnAdmin

    Me.cmdコード管理.Visible = (flgSYS = 1)
    Me.cmdコード管理.Enabled = (flgSYS = 1)
    Me.cmd職員管理.Visible = (flgSYS = 1)
    Me.cmd職員管理.Enabled = (flgSYS = 1)

End Sub

' ================================================================
' ★ フォームアンロード（×ボタン）：両ロック解放して終了
' ================================================================
Private Sub Form_Unload(Cancel As Integer)
    Call 全ロック解放
    ' DoCmd.Quit は呼ばない（Accessが自然に閉じる）
End Sub

' ================================================================
' タイトルダブルクリック：管理ボタンを強制表示（デバッグ用）
' ================================================================
Private Sub lblMタイトル_DblClick(Cancel As Integer)
    Me.cmdログイン管理.Enabled = True
    Me.cmdロック管理.Enabled = True
    Me.cmdコード管理.Enabled = True
    Me.cmd職員管理.Enabled = True
End Sub

' ================================================================
' 伺い書ボタン
' ================================================================
Private Sub cmd伺い書_Click()
    flgSyubetu = 1
    DoCmd.OpenForm "Fメイン"
    DoCmd.Close acForm, "Fメニュー"
End Sub

' ================================================================
' 企画書ボタン
' ================================================================
Private Sub cmd企画書_Click()
    flgSyubetu = 2
    DoCmd.OpenForm "Fメイン"
    DoCmd.Close acForm, "Fメニュー"
End Sub

' ================================================================
' 終了ボタン：ロック解放してから終了
' ================================================================
Private Sub cmd終了_Click()
    Call 全ロック解放
    DoCmd.Quit
End Sub

' ================================================================
' ログイン管理ボタン
' ================================================================
Private Sub cmdログイン管理_Click()
    Call GetTableDWH("Tログイン伺企")
    DoCmd.OpenForm "Fログイン管理"
    DoCmd.Close acForm, "Fメニュー"
End Sub

' ================================================================
' ロック管理ボタン
' ================================================================
Private Sub cmdロック管理_Click()
    Call GetTableDWH("Tロック伺企")
    DoCmd.OpenForm "Fロック管理"
    DoCmd.Close acForm, "Fメニュー"
End Sub

' ================================================================
' コード管理ボタン
' ================================================================
Private Sub cmdコード管理_Click()
    DoCmd.OpenForm "Fコード検索"
    DoCmd.Close acForm, "Fメニュー"
End Sub

' ================================================================
' 職員管理ボタン
' ================================================================
Private Sub cmd職員管理_Click()
    DoCmd.OpenForm "F職員管理"
    DoCmd.Close acForm, "Fメニュー"
End Sub

' ================================================================
' フォルダ作成ボタン
' ================================================================
Private Sub cmdフォルダ_Click()
    On Error GoTo err_handle
    Dim strNen    As String
    Dim strPath1  As String
    Dim strPath2  As String
    strNen = Format(Date, "ggge年")
    strPath1 = "\\flsv1\fsroot\みのり苑\総務data\伺い書検索\" & strNen & "度_PDF"
    strPath2 = "\\flsv1\fsroot\みのり苑\総務data\企画書検索\" & strNen & "度_PDF"
    If Dir(strPath1, vbDirectory) = vbNullString Then
        MkDir strPath1
        MsgBox "伺い書フォルダの作成が完了しました"
    Else
        MsgBox "伺い書フォルダはすでに作成されています"
    End If
    If Dir(strPath2, vbDirectory) = vbNullString Then
        MkDir strPath2
        MsgBox "起案・企画書フォルダの作成が完了しました"
    Else
        MsgBox "起案・企画書フォルダはすでに作成されています"
    End If
    Exit Sub
err_handle:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "フォルダ作成エラー"
End Sub