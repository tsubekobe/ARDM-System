Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
    Dim rc As Long
    Dim strPCname As String
    
    'データベースウィンドウを非表示にする
    rc = ShowWindow(Application.hWndAccessApp, SW_SHOWMINIMIZED)
    
''''    'パソコン名取得
''''    strPCname = GetMyComputerName
''''    MsgBox strPCname
''''
    DoCmd.OpenForm "Fログイン"
    DoCmd.Close acForm, Me.Form.Name

End Sub