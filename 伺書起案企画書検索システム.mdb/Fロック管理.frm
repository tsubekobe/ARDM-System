Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private intSts      As Integer

'Private Sub cmd削除_Click()
'On Error GoTo Err_cmd削除_Click
'    Call CN_INIT(intSts)
'
'    cn.BeginTrans
'
'    strSQL = ""
'    strSQL = strSQL & "delete * from Tロック伺企 "
'    strSQL = strSQL & "where 企画番号 = '" & txt企画番号 & "'"
'
'    cn.Execute strSQL
'
'    cn.CommitTrans
'
'    Call CN_END
'
'    Call データ削除("Tロック伺企", "企画番号", CStr(Me.txt企画番号))
'
'    Me.Requery
'
'Exit_cmd削除_Click:
'    Exit Sub
'
'Err_cmd削除_Click:
'    MsgBox Err.Description
'    Resume Exit_cmd削除_Click
'    cn.RollbackTrans
'
'End Sub
' ================================================================
' 修正版 cmd削除_Click（Fロック管理.frm）
' 問題: WHERE句で "企画番号" を使っているが正しくは "伺企番号"
' ================================================================
Private Sub cmd削除_Click()
On Error GoTo Err_cmd削除_Click

    Call CN_INIT(intSts)
    cn.BeginTrans

    strSQL = ""
    strSQL = strSQL & "DELETE * FROM Tロック伺企 "
    strSQL = strSQL & "WHERE 伺企番号 = '" & txt企画番号 & "'"
    ' ← "企画番号" → "伺企番号" に修正

    cn.Execute strSQL
    cn.CommitTrans
    Call CN_END

    ' ローカルテーブルからも削除
    Call データ削除("Tロック伺企", "伺企番号", CStr(Me.txt企画番号))  ' ← フィールド名修正

    Me.Requery

Exit_cmd削除_Click:
    Exit Sub

Err_cmd削除_Click:
    If cn.State = 1 Then cn.RollbackTrans  ' ← ロールバックを追加
    MsgBox Err.Description
    Resume Exit_cmd削除_Click
End Sub

Private Sub cmd戻る_Click()

    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fロック管理", acSaveNo

End Sub

Private Sub Form_Load()
    
'    Me.Caption = cstSysK
    Me.cmd削除.SetFocus

End Sub