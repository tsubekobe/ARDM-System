Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private intSts   As Integer

Private Sub cmdPDF登録_Click()

    Dim nowMonth As String
    Dim nowDay As String
    Dim intAnswer As Integer
    
On Error GoTo Err_cmdPDF登録_Click
    
    nowMonth = Format(Now(), "m")  '処理日の月
    nowDay = Format(Now(), "d")  '処理日の日
    
        If nowMonth < 4 Then
            PDF登録
        ElseIf nowMonth < 5 Then
            intAnswer = MsgBox("昨年度分の登録ですか？", vbYesNo + vbQuestion)
            If intAnswer = vbYes Then
                 PDF登録
            ElseIf intAnswer = vbNo Then
                 PDF登録
            End If
        Else  '5月～12月までの処理：今年度の処理
            PDF登録
        End If
        
    If strFname = "" Then
        Me!txtPDFリンク = "*"
    Else
        Me!txtPDFリンク = strFname  'フォルダパス名＆ファイル名
    End If
        
        
Exit_cmdPDF登録_Click:
        Exit Sub
        
Err_cmdPDF登録_Click:
        MsgBox Err.Description
        Resume Exit_cmdPDF登録_Click
End Sub

Private Sub cmd登録_Click()

On Error GoTo Err_cmd登録_Click

    Dim strTable As String

    If flgSyubetu = 1 Then
'    If 1 = 1 Then
        strTable = "T伺い書基本情報"
    ElseIf flgSyubetu = 2 Then
        strTable = "T企画書基本情報"
    End If

    '*** 登録処理
    
    If flgShinki = 1 Then
        
        'INSERT文
        strSQL = ""
        strSQL = strSQL & " INSERT INTO " & strTable
        strSQL = strSQL & " VALUES("
        strSQL = strSQL & "'" & Trim$(Me.txt処理番号) & "',"
        If Nz(Me.txt起案日, "") = "" Or Me.txt起案日 = "" Then
            strSQL = strSQL & "Null,"
        Else
            strSQL = strSQL & "#" & Trim$(Me.txt起案日) & "#,"
        End If
        If IsNull(Me.cbo種別) = True Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.cbo種別) & "',"
        End If
        If IsNull(Me.cbo施設) = True Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.cbo施設) & "',"
        End If
        If IsNull(Me.cbo所属) = True Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.cbo所属) & "',"
        End If
        If Nz(Me.txt起案者, "") = "" Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.txt起案者) & "',"
        End If
        If Nz(Me.txt件名, "") = "" Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.txt件名) & "',"
        End If
        If Nz(Me.txtPDFリンク, "") = "" Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.txtPDFリンク) & "',"
        End If
        If Nz(Me.cbo年度, "") = "" Then
            strSQL = strSQL & "'" & "" & "',"
        Else
            strSQL = strSQL & "'" & Trim$(Me.cbo年度) & "',"
        End If
        strSQL = strSQL & (Me.chk人事) & ","
        strSQL = strSQL & (Me.chk秘) & ","
        strSQL = strSQL & "#" & Trim$(Me.txtNow) & "#"
        strSQL = strSQL & " )"
        
        MsgBox "新規データを登録しました。"
    Else
        'UPDATE文
        strSQL = ""
        strSQL = "update " & strTable
        strSQL = strSQL & " SET"
        strSQL = strSQL & " 番号 = " & CStr(Nz(Me.txt処理番号, "")) & ","
        'strSQL = strSQL & " 起案日 = '" & CStr(Me.txt伺書起案日) & "',"
        strSQL = strSQL & " 起案日 = #" & Me.txt起案日 & "#,"
        If IsNull(Me.cbo種別) = True Then
            strSQL = strSQL & " 種類 = '" & "" & "',"
        Else
            strSQL = strSQL & " 種類 = '" & Trim$(CStr(Nz(Me.cbo種別, ""))) & "',"
        End If
        If IsNull(Me.cbo施設) = True Then
            strSQL = strSQL & " 施設 = '" & "" & "',"
        Else
            strSQL = strSQL & " 施設 = '" & Trim$(CStr(Nz(Me.cbo施設, ""))) & "',"
        End If
        If IsNull(Me.cbo所属) = True Then
            strSQL = strSQL & " 所属 = '" & "" & "',"
        Else
            strSQL = strSQL & " 所属 = '" & Trim$(CStr(Nz(Me.cbo所属, ""))) & "',"
        End If
        strSQL = strSQL & " 起案者 = '" & CStr(Nz(Me.txt起案者, "")) & "',"
        strSQL = strSQL & " 件名 = '" & CStr(Nz(Me.txt件名, "")) & "',"
        strSQL = strSQL & " PDFリンク = '" & CStr(Nz(Me.txtPDFリンク, "")) & "',"
        strSQL = strSQL & " 年度 = '" & CStr(Nz(Me.cbo年度, "")) & "',"
        strSQL = strSQL & " 人事 = " & (Me.chk人事) & ","
        strSQL = strSQL & " 秘 = " & (Me.chk秘)
    'WHERE句
        strSQL = strSQL & " WHERE 登録日時 = #" & Me.txt登録 & "#"
        
        MsgBox "データを修正登録しました。"

    End If
    
     Call CN_INIT(intSts)
'    MsgBox cn.State    '接続状態確認　cn=1なら接続
    Call RS_INIT(intSts)
    
    'SQLの実行（レコードの更新・追加）
    '*** トランザクション開始
    cn.BeginTrans
    cn.Execute strSQL 'SQLを実行
    '*** コミット
    cn.CommitTrans
    
    Call 排他_DEL
    
    DoCmd.Close acForm, "F新規修正"
    DoCmd.OpenForm "Fメイン"

Exit_cmd登録_Click:
    Exit Sub

Err_cmd登録_Click:
    'ODBC接続エラーの判定
    If Err.Number = 3146 Then
        MsgBox "SQLサーバーに接続できませんでした。" & vbCrLf & _
        "　システムを終了します。" & vbCritical
        'ロールバック処理
        CN1.RollbackTrans
        'システム終了
        DoCmd.Quit
    'ODBC接続以外のエラー
    Else
        MsgBox "データベース管理者に連絡してください。" & Str$(Err) & _
        Err.Description, vbCritical
        'ロールバック処理
'        cn1.RollbackTrans
        'システム終了
'        DoCmd.Quit
    End If
    Resume Exit_cmd登録_Click

End Sub

Private Sub cmd戻る_Click()
    If flgShinki = 0 Then
        排他_DEL
    End If
    
    flgSyubetu = txt種別flg
    
    DoCmd.OpenForm "Fメイン"
    DoCmd.Close acForm, "F新規修正"
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim strLine      As String
    Dim strTitle     As String
    Dim aryField()   As String
    Dim lngArySize   As Long
    Dim strJyouken2  As String
    Dim strNichiji2  As String
    Dim strTable     As String
    Dim strTouroku   As String
    Dim strNonen     As String

    flgHaita = 0

''    MsgBox Me.OpenArgs
''    MsgBox "flgHyouji: " & flgHyouji

    'パラメータ読み込み
    If IsNull(OpenArgs) Then
        Exit Sub
    Else
        strLine = OpenArgs
    End If

    'OpenArgsからパラメータ取り出しカンマ区切り分割で配列格納
    aryField = Split(strLine, ",")

    '配列数確認（開始は0）
    lngArySize = UBound(aryField)

    Me.txt抽出番号 = CStr(aryField(0))
    strNichiji2 = CStr(aryField(1))
    
    If flgSyubetu = 1 Then
        strTitle = "伺い書"
    ElseIf flgSyubetu = 2 Then
        strTitle = "企画書"
    End If

    txt種別flg = flgSyubetu

    'flgShinkiのDefaultは「0」(編集)
    If flgShinki = 0 Then
        '登録修正
        Me.txt登録 = strNichiji2
        Me.lblタイトル.Caption = "～ " & strTitle & "編集～"
        基本情報_SEL
        基本情報_DSP
    ElseIf flgShinki = 1 Then
        '新規登録
        基本情報_INIT
        基本情報_DSP
        Me.lblタイトル.Caption = "～ " & strTitle & "新規登録～"
        'フラグによりテーブルの分岐
        If flgSyubetu = 1 Then
            strTable = "T伺い書基本情報"
        ElseIf flgSyubetu = 2 Then
            strTable = "T企画書基本情報"
        Else
            Exit Sub
        End If
        
        If 受付年度(strTable) = True Then
'            MsgBox "今年度"
            strTouroku = rs!番号 + 1
'            MsgBox strTouroku
        Else
'            MsgBox "昨年度"
            strNonen = Me.txt年度2ケタ
            strTouroku = strNonen * 10000 + 1
'            MsgBox strTouroku
        End If
        Me.txt処理番号 = strTouroku
        Nendo
    End If
    
    If flgSyubetu = 1 Then
    
        Me.cbo種別.Visible = False
        Me.lbl種別.Visible = False
    ElseIf flgSyubetu = 2 Then
        Me.cbo種別.Visible = True
        Me.lbl種別.Visible = True
        
    End If
    
    If flgShinki = 0 Then
       '排他制御
        排他情報Key_INIT
        排他情報Key.職員番号 = 職員情報Key.職員番号
        排他情報Key.伺企番号 = Me.txt抽出番号
        If 排他_CHK() = True Then
            MsgBox (排他情報Key.メッセージ)
            Exit Sub
        End If
    End If
   
    
End Sub

Private Sub cbo施設_AfterUpdate()

    Me.cbo所属.RowSource = 所属リスト
    
End Sub

Private Sub cbo所属_AfterUpdate()

    Me.cbo所属.RowSource = 所属リスト
   
End Sub

Function 所属リスト()

    Dim strlist As String
    strlist = "SELECT コード名称, コード施設 FROM Tコード管理 WHERE Tコード管理.コードID = 2"
'    strSQL = "SELECT コード施設 FROM Tコード管理 WHERE (Tコード管理.コード名称)=(Forms![F新規修正]!cbo所属);"
    If IsNull(Me.cbo施設) = True Then
        所属リスト = strlist
    ElseIf Me.cbo施設 = "のへじ" Then
        所属リスト = strlist & " AND Tコード管理.コード施設 = '老健のへじ' OR Tコード管理.コード施設 = '福祉センター'"
    Else
        所属リスト = strlist & " AND Tコード管理.コード施設 = '" & Me.cbo施設 & "'"
    End If

End Function

Private Sub 修正完了_Click()
   MsgBox "修正完了"
   DoCmd.Close
End Sub

Private Sub 終了_Click()

On Error GoTo Err_終了_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    '排他制御
    排他_DEL
    
    DoCmd.Close
    
Exit_終了_Click:
    Exit Sub

Err_終了_Click:
    MsgBox Err.Description
    Resume Exit_終了_Click
    
End Sub

Sub 基本情報_DSP()

    With Me
        .txt処理番号.Value = Trim$(CStr(TBL基本情報.処理番号))
        .txt起案日.Value = Trim$(CStr(TBL基本情報.起案日))
        .cbo種別.Value = TBL基本情報.種類
        .cbo施設.Value = TBL基本情報.施設
        .cbo所属.Value = TBL基本情報.所属
        .txt起案者.Value = Trim$(CStr(TBL基本情報.起案者))
        .txt件名.Value = Trim$(CStr(TBL基本情報.件名))
        .txtPDFリンク.Value = Trim$(CStr(TBL基本情報.PDFリンク))
        .cbo年度.Value = Trim$(CStr(TBL基本情報.年度))
        .chk人事.Value = TBL基本情報.人事
        .chk秘.Value = TBL基本情報.秘
        .txt登録.Value = Trim$(CStr(TBL基本情報.登録日時))
    End With
    
End Sub

Sub 基本情報_SEL()

    If flgSyubetu = 1 Then
        strTable = "T伺い書基本情報"
    ElseIf flgSyubetu = 2 Then
        strTable = "T企画書基本情報"
    End If

    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Sub
    End If
    
    'SELECT文
    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM " & strTable
    'WHERE句
    strSQL = strSQL & " WHERE 登録日時 = #" & Me.txt登録 & "#;"
    
    'RSオープン
    Call RS_INIT(intSts)
    rs.Open strSQL, cn, adOpenStatic, adLockOptimistic
    If intSts <> DB_OK Then
        GoTo Exit_基本情報_SEL
    End If
    
    With TBL基本情報
        .処理番号 = Nz(rs.Fields("番号").Value, "")
        .起案日 = Nz(rs.Fields("起案日").Value, "")
        .種類 = Nz(rs.Fields("種類").Value, "")
        .施設 = Nz(rs.Fields("施設").Value, "")
        .所属 = Nz(rs.Fields("所属").Value, "")
        .起案者 = Nz(rs.Fields("起案者").Value, "")
        .件名 = Nz(rs.Fields("件名").Value, "")
        .PDFリンク = Nz(rs.Fields("PDFリンク").Value, "")
        .年度 = Nz(rs.Fields("年度").Value, "")
        .人事 = Nz(rs.Fields("人事").Value, 0)
        .秘 = Nz(rs.Fields("秘").Value, 0)
        .登録日時 = Nz(rs.Fields("登録日時").Value, "")
    End With
    
    RS_END
    CN_END
   
Exit_基本情報_SEL:
    Exit Sub

Err_基本情報_SEL:
    MsgBox Err.Description

End Sub

Sub Nendo()

    Dim strSakunen As String
    Dim strKotoshi As String
    Dim strlist    As String
    
    strSakunen = Year(Me.txtNow) - 1
    strKotoshi = Year(Me.txtNow)
    strlist = strlist & strKotoshi & ";"
    strlist = strlist & strSakunen
    If Month(Me.txtNow) < 4 Then
        Me.cbo年度 = strSakunen
    ElseIf Month(Me.txtNow) > 4 Then
        Me.cbo年度 = strKotoshi
    Else
        Me.cbo年度.RowSource = strlist
        Me.cbo年度.DefaultValue = Me.cbo年度.ItemData(0)
    End If

End Sub

Function 受付年度(strTable)
    Dim strNen     As String
    Dim strNonen     As String
    
    受付年度 = False
    
    Call CN_INIT(intSts)
'    MsgBox cn.State  '接続状態 接続時：1、非接続時：0
    strSQL = "SELECT max(番号) as 番号 FROM " & strTable 'SQL文
    Call RS_INIT(intSts)
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic        'SQLを実行
    
    strNen = Left(CStr(rs!番号), 2)
    strNonen = Right(Year(DateAdd("m", -3, Date)), 2)
    
    Me.txt年度2ケタ = strNonen
    If strNen <> strNonen Then
        Exit Function
    End If
    受付年度 = True
    
End Function