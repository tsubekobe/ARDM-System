Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ================================================================
' cmb職員 マウスダウン：エラーメッセージを隠す
' ================================================================
Private Sub cmb職員_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtMsg.Value = ""
    txtMsg.Visible = False
End Sub

' ================================================================
' フォームロード：マスタ同期 ＋ 端末名取得 ＋ 職員自動選択
' ================================================================
Private Sub Form_Load()

    Dim intRtn  As Integer
    Dim CN1     As New ADODB.Connection
    Dim RC1     As New ADODB.Recordset
    Dim dbDAO1  As DAO.Database
    Dim rsDAO1  As DAO.Recordset

    Me.txtMsg.Visible = False

    ' ── マスタ同期 ──────────────────────────────────────────────────
    Call GetTableDWH("Tコード管理")

    ' T施設所属 更新
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM T施設所属"
    DoCmd.SetWarnings True
    CN1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_SERVER_PATH
    RC1.Open "Q施設所属", CN1
    Set dbDAO1 = CurrentDb
    Set rsDAO1 = dbDAO1.OpenRecordset("T施設所属", dbOpenDynaset)
    Do Until RC1.EOF
        rsDAO1.AddNew
        rsDAO1!施設 = RC1!施設
        rsDAO1!所属 = RC1!所属
        rsDAO1!年度 = RC1!年度
        rsDAO1.Update
        RC1.MoveNext
    Loop
    RC1.Close:  Set RC1 = Nothing
    CN1.Close:  Set CN1 = Nothing
    rsDAO1.Close: Set rsDAO1 = Nothing
    dbDAO1.Close: Set dbDAO1 = Nothing

    ' T年度 更新
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM T年度"
    DoCmd.SetWarnings True
    CN1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_SERVER_PATH
    RC1.Open "Q年度", CN1
    Set dbDAO1 = CurrentDb
    Set rsDAO1 = dbDAO1.OpenRecordset("T年度", dbOpenDynaset)
    Do Until RC1.EOF
        rsDAO1.AddNew
        rsDAO1!年度 = RC1!年度
        rsDAO1.Update
        RC1.MoveNext
    Loop
    RC1.Close:  Set RC1 = Nothing
    CN1.Close:  Set CN1 = Nothing
    rsDAO1.Close: Set rsDAO1 = Nothing
    dbDAO1.Close: Set dbDAO1 = Nothing

    ' ── グローバル変数初期化 ────────────────────────────────────────
    flgHyouji = 0
    flgSyubetu = 0
    flgForm = 0
    flgChk = 0
    flgSYS = 0

    ' ── 端末名取得 ──────────────────────────────────────────────────
    Call 職員情報Key_INIT
    職員情報Key.処理端末 = StrConv(GetMyComputerName, vbUpperCase)

    ' ── 端末に対応する職員を自動選択 ────────────────────────────────
    intRtn = 職員管理_SEL
    If intRtn = RTN_OK Then
        Me.cmb職員.Value = 職員情報Key.職員番号
        Me.cmdログイン.SetFocus
        Exit Sub
    End If

    Me.cmb職員.Value = 0
    Me.cmb職員.SetFocus

End Sub

' ================================================================
' ログインボタン：職員選択のみ確認してメニューへ遷移
' ★ パスワード認証はここでは行わない（編集/新規ボタン押下時に実施）
' ================================================================
Private Sub cmdログイン_Click()

    Dim intRtn As Integer

    ' 職員未選択チェック
    If Nz(Me.cmb職員.Value, 0) = 0 Then
        Me.txtMsg.Visible = True
        Me.txtMsg.Value = cstMsg01   ' "職員を選択してください"
        Exit Sub
    End If

    ' 職員情報を退避
    With 職員情報Key
        .職員番号 = Me.cmb職員.Column(0, Me.cmb職員.ListIndex)
        .職員氏名 = Me.cmb職員.Column(1, Me.cmb職員.ListIndex)
        .所属部門 = Me.cmb職員.Column(2, Me.cmb職員.ListIndex)
        .使用区分 = Me.cmb職員.Column(3, Me.cmb職員.ListIndex)
    End With

    ' flgHyouji：所属部門による表示レベル設定
    '   4-5 → 中間管理 (1)、 6以上 → システム管理者 (2)
    '   ★ ここではPW確認しない。編集ボタン押下時に確認する。
    Dim strBumon As String
    strBumon = Trim$(職員情報Key.所属部門)
    If strBumon >= "6" Then
        flgHyouji = 2
        flgSYS = 1
    ElseIf strBumon >= "4" Then
        flgHyouji = 1
    Else
        flgHyouji = 0
        flgSYS = 0
    End If

    ' ── 二重ログインチェック ────────────────────────────────────────
    Call TBLログイン_INIT
    TBLログイン.職員番号 = 職員情報Key.職員番号
    intRtn = ログイン_SEL
    If intRtn = RTN_OK Then
        ' すでにログイン中
        Me.txtMsg.Visible = True
        Me.txtMsg.Value = cstMsg03   ' "選択された職員はログインしています"
        Exit Sub
    End If

    ' ── ログインレコード登録（ログインロック） ───────────────────────
    Call TBLログイン_INIT
    TBLログイン.職員番号 = 職員情報Key.職員番号
    TBLログイン.職員氏名 = Trim$(職員情報Key.職員氏名)
    TBLログイン.処理端末 = Trim$(職員情報Key.処理端末)
    TBLログイン.処理日時 = CStr(Now)
    intRtn = ログイン_INS
    If intRtn <> RTN_OK Then Exit Sub

    ' ── メニューへ遷移 ──────────────────────────────────────────────
    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fログイン"

End Sub

' ================================================================
' 終了ボタン
' ================================================================
Private Sub cmd終了_Click()
    DoCmd.Quit
End Sub

Private Sub cmdログイン_GotFocus()
    Me.cmdログイン.FontBold = True
End Sub
Private Sub cmdログイン_LostFocus()
    Me.cmdログイン.FontBold = False
End Sub
Private Sub cmd終了_GotFocus()
    Me.cmd終了.FontBold = True
End Sub
Private Sub cmd終了_LostFocus()
    Me.cmd終了.FontBold = False
End Sub