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

Private Sub cmd削除_Click()
On Error GoTo Err_cmd削除_Click
    Call CN_INIT(intSts)
    
    cn.BeginTrans
    
    strSQL = ""
    strSQL = strSQL & "delete * from Tログイン伺企 "
    strSQL = strSQL & "where 職員氏名 = '" & txt職員氏名 & "'"
    
    cn.Execute strSQL

    cn.CommitTrans

    Call CN_END
    
    Call データ削除("Tログイン伺企", "職員氏名", CStr(Me.txt職員氏名))
    
    Me.Requery
    
Exit_cmd削除_Click:
    Exit Sub

Err_cmd削除_Click:
    MsgBox Err.Description
    Resume Exit_cmd削除_Click
    cn.RollbackTrans

End Sub

Private Sub cmd戻る_Click()

    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fログイン管理", acSaveNo

End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_ERR

'    Me.Caption = cstSysU
    Me.cmd削除.SetFocus

Form_Load_EXIT:
    Exit Sub

Form_Load_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ログイン管理(SELECT)"
    GoTo Form_Load_EXIT
End Sub