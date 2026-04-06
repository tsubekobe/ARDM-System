Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmd選択_Click()

'    コード管理Key.コード = Me.コード
    DoCmd.OpenForm "Fコード管理", , , "コードID = " & Me.txtコード
    DoCmd.Close acForm, "Fコード検索"

End Sub

Private Sub cmd戻る_Click()

    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fコード検索"

End Sub

Private Sub Form_Load()

    Me.cmd選択.SetFocus

End Sub