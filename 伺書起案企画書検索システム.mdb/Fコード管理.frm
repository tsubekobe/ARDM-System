Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmd削除_Click()
On Error GoTo Err_cmd削除_Click

    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord

Exit_cmd削除_Click:
    Exit Sub

Err_cmd削除_Click:
    MsgBox Err.Description
    Resume Exit_cmd削除_Click
    
End Sub

Private Sub cmd戻る_Click()

    DoCmd.OpenForm "Fコード検索"
    DoCmd.Close acForm, "Fコード管理", acSaveNo

End Sub

Private Sub Form_Load()
    
'    Me.txtコード区分.DefaultValue = コード管理Key.コード

End Sub