Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmb職員_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtMsg.Value = ""
    txtMsg.Visible = False
End Sub

Private Sub Form_Load()

    Dim intRtn  As Integer
    Dim strPara As String
    Dim rc As Long
    Dim CN1 As New ADODB.Connection
    Dim RC1 As New ADODB.Recordset
    Dim dbDAO1 As DAO.Database
    Dim rsDAO1 As DAO.Recordset
    Dim strFd As String
    
    Me.txtMsg.Visible = False

    Call GetTableDWH("Tコード管理")
    
    'アラートメッセージを停止
    DoCmd.SetWarnings False
    DoCmd.RunSQL "Delete * from T施設所属"
    'アラートメッセージを再開
    DoCmd.SetWarnings True
    
     'T施設所属 値更新
    strFd = "\\flsv1\fsroot\towada\福祉の里\DataBase\伺書起案企画書DB.mdb"
    CN1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFd
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

    RC1.Close: Set RC1 = Nothing
    CN1.Close: Set CN1 = Nothing
    rsDAO1.Close: Set rsDAO1 = Nothing
    dbDAO1.Close: Set dbDAO1 = Nothing

    'アラートメッセージを停止
    DoCmd.SetWarnings False
    DoCmd.RunSQL "Delete * from T年度"
    'アラートメッセージを再開
    DoCmd.SetWarnings True
    
    'T年度テーブル 値更新
    CN1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFd
    RC1.Open "Q年度", CN1
    
    Set dbDAO1 = CurrentDb
    Set rsDAO1 = dbDAO1.OpenRecordset("T年度", dbOpenDynaset)

    Do Until RC1.EOF
        rsDAO1.AddNew
        rsDAO1!年度 = RC1!年度
        rsDAO1.Update
        RC1.MoveNext
    Loop

    RC1.Close: Set RC1 = Nothing
    CN1.Close: Set CN1 = Nothing
    rsDAO1.Close: Set rsDAO1 = Nothing
    dbDAO1.Close: Set dbDAO1 = Nothing
    
    '変数初期化
    flgHyouji = 0
    flgSyubetu = 0
    flgForm = 0
'    flgTouroku = 0
    flgChk = 0
    
    
    '入力
    Me.cmb職員.Value = 0
    
    intRtn = 1 '取得(PintKBN)
    
    '端末名取得
    Call 職員情報Key_INIT
    職員情報Key.処理端末 = GetMyComputerName
    職員情報Key.処理端末 = StrConv(職員情報Key.処理端末, vbUpperCase)
    
    '職員管理
    intRtn = 職員管理_SEL
    If intRtn = RTN_OK Then
        Me.cmb職員.Value = 職員情報Key.職員番号
        Me.cmdログイン.SetFocus
        Exit Sub
    End If
    
    '初期フォーカス設定
    Me.cmb職員.SetFocus

End Sub

Private Sub cmdログイン_Click()

    Dim intRtn As Integer
    flgSYS = 0
    
    '職員が選択されていない場合
    If Nz(Me.cmb職員.Value, 0) = 0 Then
        Me.txtMsg.Value = cstMsg05
        Exit Sub
    End If
    
    '職員情報退避
    職員情報Key.職員番号 = Me.cmb職員.Column(0, Me.cmb職員.ListIndex)
    職員情報Key.職員氏名 = Me.cmb職員.Column(1, Me.cmb職員.ListIndex)
    職員情報Key.所属部門 = Me.cmb職員.Column(2, Me.cmb職員.ListIndex)
    職員情報Key.使用区分 = Me.cmb職員.Column(3, Me.cmb職員.ListIndex)
    
    If Trim$(職員情報Key.所属部門) > "3" And Trim$(職員情報Key.所属部門) < "6" Then
        If InputBoxDK("システム管理者パスワードを入力してください", "認証") <> "jimu1319s" Then
            MsgBox "パスワードが違います。"
            Exit Sub
        End If
        flgHyouji = 1
    End If
     If Trim$(職員情報Key.所属部門) > "5" Then
        If InputBoxDK("システム管理者パスワードを入力してください", "認証") <> "sys0120310272" Then
            MsgBox "パスワードが違います。"
            Exit Sub
        End If
        flgHyouji = 2
        flgSYS = 1
    End If
   
    'ログインチェック
    Call TBLログイン_INIT
    TBLログイン.職員番号 = 職員情報Key.職員番号
    intRtn = ログイン_SEL
    If intRtn = RTN_OK Then
        Me.txtMsg.Visible = True
        Me.txtMsg.Value = cstMsg03
        Exit Sub
    End If
    
    'ログイン
    Call TBLログイン_INIT
    TBLログイン.職員番号 = 職員情報Key.職員番号
    TBLログイン.職員氏名 = Trim$(職員情報Key.職員氏名)
    TBLログイン.処理端末 = Trim$(職員情報Key.処理端末)
    TBLログイン.処理日時 = CStr(Now)
    intRtn = ログイン_INS
    If intRtn <> RTN_OK Then
        Exit Sub
    End If
    
    'メニュー画面へ遷移
    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fログイン"

End Sub

Private Sub cmd終了_Click()
    
'    システム終了
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