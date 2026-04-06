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
' モジュールレベル変数
' ================================================================
Dim strShisetu   As String
Dim strSyozoku   As String
Dim strBangou    As String
Public strTable  As String
Private intSts   As Integer
Private rs2      As ADODB.Recordset   ' ★ 追加（未宣言だった）
Private strQuery As String            ' ★ 追加（未宣言だった）


' ================================================================
' cbo年度 更新後
' ================================================================
Private Sub cbo年度_AfterUpdate()

    strSQL = "SELECT DISTINCT 年度 FROM T年度 ORDER BY 年度 DESC"
    Me.cbo年度.RowSource = strSQL

    strNendo = Me.cbo年度
    flgChk = 0
    検索詳細

    Me.cbo施設検索.RowSource = 施設_SEL1()
    Me.cbo所属検索.RowSource = 所属_SEL1()

End Sub


' ================================================================
' cbo年度 クリック
' ================================================================
Private Sub cbo年度_Click()

    If chk全検索 = True Then
        chk全検索 = False
    End If

End Sub


' ================================================================
' 全年度検索チェックボックス 更新後
' ================================================================
Private Sub chk全検索_AfterUpdate()

    If chk全検索 = True Then

        flgChk = 1

        Me.txt氏名検索 = Null
        Me.txt件名検索 = Null
        Me.txt開始日検索 = Null
        Me.txt終了日検索 = Null
        Me.cbo年度 = ""

        検索詳細
        Me.cbo所属検索.RowSource = 所属_SEL1()
        Me.cbo施設検索.RowSource = 施設_SEL1()

    Else
        strSQL = "SELECT DISTINCT " & strTable & ".[年度] FROM " & strTable & ";"
        Me.cbo年度.RowSource = strSQL
        検索詳細
        strNendo = Me.cbo年度
    End If

End Sub


' ================================================================
' 全年度検索チェックボックス 更新前
' ================================================================
Private Sub chk全検索_BeforeUpdate(Cancel As Integer)

    Me.cbo施設検索.RowSource = ""
    Me.cbo所属検索.RowSource = ""

End Sub


' ================================================================
' 検索条件クリアボタン クリック
' ================================================================
Private Sub cmd条件クリア_Click()

    Me.cbo施設検索 = Null
    Me.cbo施設検索.RowSource = "Q施設名"
    Me.cbo所属検索 = Null
    Me.cbo所属検索.RowSource = "SELECT コード名称,コード施設,コードID, コード " & _
                               "FROM Tコード管理 " & _
                               "WHERE (Tコード管理.コードID)=2 " & _
                               "ORDER BY コード;"
    Me.txt氏名検索 = Null
    Me.txt件名検索 = Null
    Me.txt開始日検索 = Null
    Me.txt終了日検索 = Null
    Me.opt種類.Value = 3

    strShisetuEx = ""
    strSyozokuEx = ""
    strShimeiEx = ""
    strKenmeiEx = ""
    strKaishiEx = ""
    strSyuryoEx = ""

    strSQL = "SELECT * FROM T年度;"
    Me.cbo年度.RowSource = strSQL
    検索詳細

End Sub


' ================================================================
' 新規登録ボタン クリック
' ================================================================
Private Sub cmd登録_Click()

    Dim dataArgs As String

    ' ── ★ PW認証チェック ────────────────────────────────────────────
    If Not PW認証チェック() Then Exit Sub

    strBangou = ""
    strNichiji = ""
    flgShinki = 1

    dataArgs = strBangou & "," & strNichiji
    DoCmd.OpenForm "F新規修正", , , , , , dataArgs
    DoCmd.Close acForm, "Fメイン"

End Sub


' ================================================================
' 編集ボタン：PW認証を通過した場合のみF新規修正を開く
' ================================================================
Private Sub cmd編集_Click()

    Dim dataArgs As String

    ' ── ★ PW認証チェック ────────────────────────────────────────────
    If Not PW認証チェック() Then Exit Sub

    strBangou = Forms!Fメイン!情報Sub!txt番号
    strNichiji = Forms!Fメイン!情報Sub!登録日時

    dataArgs = strBangou & "," & strNichiji
    flgShinki = 0

    DoCmd.OpenForm "F新規修正", , , , , , dataArgs

    If flgOwari = 1 Then
        DoCmd.Close acForm, "F新規修正"
        Exit Sub
    End If

    If flgHaita = 1 Then
        DoCmd.Close acForm, "F新規修正"
        DoCmd.OpenForm "Fメイン"
        Forms!Fメイン.情報Sub.Requery
    Else
        DoCmd.Close acForm, "Fメイン"
    End If

End Sub


' ================================================================
' 戻るボタン：文書ロック＋ログインロックを解放してメニューへ
' ================================================================
Private Sub cmd戻る_Click()

    ' ★ 両ロック解放（共通ルーティン）
    Call 全ロック解放

    strShisetuEx = ""
    strSyozokuEx = ""
    strShimeiEx = ""
    strKenmeiEx = ""
    strKaishiEx = ""
    strSyuryoEx = ""

    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fメイン"

End Sub


' ================================================================
' フォームオープン時
' ================================================================
Private Sub Form_Open(Cancel As Integer)

    Dim flgJyouken As Integer
    flgOwari = 0
    flgShinki = 0

    Me.cbo施設検索 = Null
    Me.cbo施設検索.RowSource = 施設_SEL2()
    Me.cbo所属検索 = Null
    Me.cbo所属検索.RowSource = 所属_SEL2()

    Me.opt種類 = 3

    If CStr(flgSyubetu) = "1" Then
        strTable = "T伺い書基本情報"
        Me.lblタイトル.Caption = "伺い書検索システム"
        Me.情報Sub.SourceObject = "F伺書情報Sub"
        Me.opt種類.Visible = False
    ElseIf CStr(flgSyubetu) = "2" Then
        strTable = "T企画書基本情報"
        Me.lblタイトル.Caption = "起案・企画書検索システム"
        Me.情報Sub.SourceObject = "F企画情報Sub"
        Me.opt種類.Visible = True
    End If

    If Trim$(職員情報Key.所属部門) >= "5" Then
        Me.cmd登録.Visible = True
        Me.cmd編集.Visible = True
    Else
        Me.cmd登録.Visible = False
        Me.cmd編集.Visible = False
    End If

    If flgChk <> 1 Then
        Me.chk全検索 = False
    Else
        Me.chk全検索 = True
        chk全検索_AfterUpdate
    End If

    If strNendo = "" Then
        strNendo = Me.cbo年度
        flgJyouken = 1
    ElseIf strNendo <> "" Then
        Me.cbo年度 = strNendo
        flgJyouken = 1
    End If

    If strShisetuEx <> "" Then
        Me.cbo施設検索 = strShisetuEx
        flgJyouken = 1
    End If
    If strSyozokuEx <> "" Then
        Me.cbo所属検索 = strSyozokuEx
        flgJyouken = 1
    End If
    If strShimeiEx <> "" Then
        Me.txt氏名検索 = strShimeiEx
        flgJyouken = 1
    End If
    If strKenmeiEx <> "" Then
        Me.txt件名検索 = strKenmeiEx
        flgJyouken = 1
    End If
    If strKaishiEx <> "" Then
        Me.txt開始日検索 = strKaishiEx
        flgJyouken = 1
    End If
    If strSyuryoEx <> "" Then
        Me.txt終了日検索 = strSyuryoEx
        flgJyouken = 1
    End If

    検索詳細

    strShisetu = 施設_SEL2()
    strSyozoku = 所属_SEL2()
    Me.cbo施設検索.RowSource = strShisetu
    Me.cbo所属検索.RowSource = strSyozoku

    flgForm = 1

End Sub


' ================================================================
' 各コントロール AfterUpdate
' ================================================================
Private Sub cmd期間抽出_Click()
    検索詳細
End Sub

Private Sub cbo施設検索_AfterUpdate()
    If IsNull(Me.cbo施設検索) = True Then
        strShisetuEx = ""
    End If
    検索詳細
    Me.cbo施設検索.RowSource = 施設_SEL2()
    Me.cbo所属検索.RowSource = 所属_SEL2()
End Sub

Private Sub cbo所属検索_AfterUpdate()
    If IsNull(Me.cbo所属検索) = True Then
        strSyozokuEx = ""
    End If
    検索詳細
    Me.cbo施設検索.RowSource = 施設_SEL2()
    Me.cbo所属検索.RowSource = 所属_SEL2()
End Sub

Private Sub opt種類_AfterUpdate()
    検索詳細
End Sub

Private Sub txt開始日検索_AfterUpdate()
    検索詳細
End Sub

Private Sub txt終了日検索_AfterUpdate()
    検索詳細
End Sub

Private Sub txt氏名検索_AfterUpdate()
    検索詳細
End Sub

Private Sub txt件名検索_AfterUpdate()
    検索詳細
End Sub


' ================================================================
' 検索詳細
' ================================================================
Sub 検索詳細(Optional strAddWhere As String = "")

    Dim strExtract1 As String
    Dim strExtract2 As String
    Dim strSyurui   As String
    Dim strFilter   As String
    Dim strm        As ADODB.Stream

On Error GoTo 検索詳細_ERR

    strExtract1 = ""
    strExtract2 = ""
    strSyurui = ""
    strFilter = ""

    ' ── クエリ名の決定 ──────────────────────────────────────────────
    If CStr(flgSyubetu) = "1" Then
        If flgHyouji >= 1 Then
            strQuery = "Q伺い書基本"
        Else
            strQuery = "Q伺い書制限"
        End If
    ElseIf CStr(flgSyubetu) = "2" Then
        strQuery = "Q企画書基本"
    End If

    ' ── 年度条件 ────────────────────────────────────────────────────
    If flgForm = 0 Then
        strNendo = Year(DateAdd("m", -3, Date))
        Me.cbo年度 = strNendo
        strExtract1 = " 年度 Like '*" & Me.cbo年度 & "*' "
    Else
        If IsNull(Me.cbo年度) = False And Me.cbo年度 <> "" Then
            strExtract1 = " 年度 Like '*" & Me.cbo年度 & "*' "
            strNendo = Me.cbo年度
        End If
    End If

    ' ── 施設・所属・氏名・件名 ──────────────────────────────────────
    If IsNull(Me.cbo施設検索) = False Then
        If strExtract2 <> "" Then strExtract2 = strExtract2 & " And "
        strExtract2 = strExtract2 & " 施設 Like '*" & Me.cbo施設検索 & "*' "
        strShisetuEx = Me.cbo施設検索
    End If

    If IsNull(Me.cbo所属検索) = False Then
        If strExtract2 <> "" Then strExtract2 = strExtract2 & " And "
        strExtract2 = strExtract2 & " 所属 Like '*" & Me.cbo所属検索 & "*' "
        strSyozokuEx = Me.cbo所属検索
    End If

    If IsNull(Me.txt氏名検索) = False Then
        If strExtract2 <> "" Then strExtract2 = strExtract2 & " And "
        strExtract2 = strExtract2 & " 起案者 Like '*" & Me.txt氏名検索 & "*' "
        strShimeiEx = Me.txt氏名検索
    End If

    If IsNull(Me.txt件名検索) = False Then
        If strExtract2 <> "" Then strExtract2 = strExtract2 & " And "
        strExtract2 = strExtract2 & " 件名 Like '*" & Me.txt件名検索 & "*' "
        strKenmeiEx = Me.txt件名検索
    End If

    ' ── 日付範囲 ────────────────────────────────────────────────────
    If IsNull(Me.txt開始日検索) = False Then
        strExtract1 = ""
        If strExtract2 <> "" Then strExtract2 = strExtract2 & " And "
        strExtract2 = strExtract2 & " 起案日 >= #" & Me.txt開始日検索 & "# "
        strKaishiEx = Me.txt開始日検索
    End If

    If IsNull(Me.txt終了日検索) = False Then
        strExtract1 = ""
        If strExtract2 <> "" Then strExtract2 = strExtract2 & " And "
        strExtract2 = strExtract2 & " 起案日 <= #" & Me.txt終了日検索 & "# "
        strSyuryoEx = Me.txt終了日検索
    End If

    ' ── 種類フィルタ ────────────────────────────────────────────────
    Select Case Me.opt種類
        Case 1:    strSyurui = " And 種類 = '起案書'"
        Case 2:    strSyurui = " And 種類 = '企画書'"
        Case Else: strSyurui = ""
    End Select

    ' ── フィルタ文字列の最終組み立て ────────────────────────────────
    If strExtract1 <> "" And strExtract2 <> "" Then
        strFilter = strExtract1 & " And " & strExtract2
    ElseIf strExtract1 <> "" Then
        strFilter = strExtract1
    Else
        strFilter = strExtract2
    End If

    If strSyurui <> "" Then
        If strFilter <> "" Then
            strFilter = strFilter & strSyurui
        Else
            strFilter = Mid(strSyurui, 6)   ' 先頭の " And " 5文字を除去
        End If
    End If

    ' ── ADOでレコードセット取得しサブフォームにセット ────────────────
    Call CN_INIT(intSts)
    strSQL = "SELECT * FROM " & strQuery & " ORDER BY 番号 DESC"
    Call RS_INIT(intSts)

    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic

    Set strm = New ADODB.Stream
    strm.Open

    If strFilter <> "" Then
        rs.Filter = strFilter
    End If

    rs.Save strm, adPersistADTG

    Set rs2 = New ADODB.Recordset
    rs2.ActiveConnection = "Provider=MSPersist"
    rs2.Open strm

    Set Me.情報Sub.Form.Recordset = rs2

    rs2.Close: Set rs2 = Nothing
    Call RS_END
    Call CN_END

    Exit Sub

検索詳細_ERR:
    If Not rs2 Is Nothing Then rs2.Close: Set rs2 = Nothing
    Call RS_END
    Call CN_END
    MsgBox "検索処理でエラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation, "検索エラー"

End Sub


' ================================================================
' 年度コンボ選択肢を再構築
' ================================================================
Sub 年度取得()

    With Me.cbo年度
        .RowSource = "SELECT DISTINCT " & strTable & ".[年度] FROM " & strTable & ";"
        .Requery
        If .ListCount > 0 Then
            .Value = .ItemData(0)
        Else
            .Value = Null
        End If
    End With

End Sub


' ================================================================
' 施設・所属 RowSource SQL 生成関数
'
' 【修正の核心】
'   旧コードの問題点:
'     施設_SEL1 は INNER JOIN を使っているが、
'     施設_SEL2 / 所属_SEL2 はその末尾に " AND 所属 = '...' " を単純連結していた。
'     INNER JOIN を含む SELECT 文に WHERE 句なしで AND を追記すると
'     「FROM句の構文エラー」になる。
'     正しくは、JOIN文の後に WHERE を付けてから AND で条件を追加しなければならない。
'
'   修正方針:
'     ① 施設_SEL1 / 所属_SEL1 は「WHERE なし」バージョンと「WHERE あり」バージョンを
'        内部で使い分け、追加条件が来たときに WHERE / AND を正しく付けられるようにする。
'     ② 施設_SEL2 / 所属_SEL2 は条件の有無で WHERE / AND を切り替える。
'     ③ 全年度検索(chk全検索=True)時はJOINなし・年度条件なしのシンプルなSQLを返す。
' ================================================================

' ──────────────────────────────────────────────────────────────────
' 施設_SEL1:
'   施設コンボの基本リスト（所属による絞り込みなし）
'   戻り値: SELECT文文字列
' ──────────────────────────────────────────────────────────────────
Function 施設_SEL1() As String

    '全年度検索の場合 → INNER JOIN あり・WHERE なし
    If Me.chk全検索 = True Then
        施設_SEL1 = "SELECT DISTINCT T施設所属.施設 " & _
                    "FROM T施設所属 INNER JOIN Q施設名 " & _
                    "ON T施設所属.施設 = Q施設名.コード名称"

    '年度を指定した通常検索の場合 → INNER JOIN あり・WHERE あり
    Else
        Dim str年度条件 As String

        If Nz(Me.cbo年度, "") = "" Then
            '年度未選択 → 年度条件なし（JOINだけ）
            施設_SEL1 = "SELECT DISTINCT T施設所属.施設 " & _
                        "FROM T施設所属 INNER JOIN Q施設名 " & _
                        "ON T施設所属.施設 = Q施設名.コード名称"
        Else
            '年度選択あり → WHERE で絞る
            If CInt(Nz(Me.cbo年度, 9999)) < 2006 Then
                str年度条件 = "T施設所属.年度 = '2005以前'"
            Else
                str年度条件 = "T施設所属.年度 = '" & Me.cbo年度 & "'"
            End If

            施設_SEL1 = "SELECT DISTINCT T施設所属.施設 " & _
                        "FROM T施設所属 INNER JOIN Q施設名 " & _
                        "ON T施設所属.施設 = Q施設名.コード名称 " & _
                        "WHERE " & str年度条件
        End If
    End If

End Function


' ──────────────────────────────────────────────────────────────────
' 施設_SEL2:
'   施設コンボのリスト（所属で絞り込みあり）
'   所属コンボが選択されているときだけ条件を追加する
' ──────────────────────────────────────────────────────────────────
Function 施設_SEL2() As String

    Dim strBase As String
    strBase = 施設_SEL1()   ' ← 括弧必須（関数として呼ぶ）

    If IsNull(Me.cbo所属検索) = True Or Nz(Me.cbo所属検索, "") = "" Then
        '所属未選択 → 基本SQLをそのまま返す
        施設_SEL2 = strBase
    Else
        '所属が選択されている
        '基本SQLにWHEREが含まれているかどうかで AND/WHERE を切り替える
        If InStr(UCase(strBase), "WHERE") > 0 Then
            施設_SEL2 = strBase & " AND T施設所属.所属 = '" & Me.cbo所属検索 & "'"
        Else
            施設_SEL2 = strBase & " WHERE T施設所属.所属 = '" & Me.cbo所属検索 & "'"
        End If
    End If

End Function


' ──────────────────────────────────────────────────────────────────
' 所属_SEL1:
'   所属コンボの基本リスト（施設による絞り込みなし）
' ──────────────────────────────────────────────────────────────────
Function 所属_SEL1() As String

    If Me.chk全検索 = True Then
        所属_SEL1 = "SELECT DISTINCT 所属 FROM T施設所属"

    Else
        If Nz(Me.cbo年度, "") = "" Then
            所属_SEL1 = "SELECT DISTINCT 所属 FROM T施設所属"

        ElseIf CInt(Nz(Me.cbo年度, 9999)) < 2006 Then
            所属_SEL1 = "SELECT DISTINCT 所属 FROM T施設所属 " & _
                        "WHERE 年度 = '2005以前'"
        Else
            所属_SEL1 = "SELECT DISTINCT 所属 FROM T施設所属 " & _
                        "WHERE 年度 = '" & Me.cbo年度 & "'"
        End If
    End If

End Function


' ──────────────────────────────────────────────────────────────────
' 所属_SEL2:
'   所属コンボのリスト（施設で絞り込みあり）
' ──────────────────────────────────────────────────────────────────
Function 所属_SEL2() As String

    Dim strBase As String
    strBase = 所属_SEL1()   ' ← 括弧必須（関数として呼ぶ）

    If IsNull(Me.cbo施設検索) = True Or Nz(Me.cbo施設検索, "") = "" Then
        所属_SEL2 = strBase
    Else
        If InStr(UCase(strBase), "WHERE") > 0 Then
            所属_SEL2 = strBase & " AND 施設 = '" & Me.cbo施設検索 & "'"
        Else
            所属_SEL2 = strBase & " WHERE 施設 = '" & Me.cbo施設検索 & "'"
        End If
    End If

End Function

' ================================================================
' ★ 新規追加：×ボタンでフォームを閉じたとき両ロック解放
' ================================================================
Private Sub Form_Unload(Cancel As Integer)
    Call 全ロック解放
End Sub

' ================================================================
' ★ 新規追加：PW認証チェック関数
'
' 役割：
'   - flgPwOk が既に認証済みなら即 True を返す（セッション中は再入力不要）
'   - 所属部門に応じたパスワードを確認し、成功なら flgPwOk を更新して True
'   - 失敗 or キャンセルなら False
'
' 呼び出し元：cmd登録_Click、cmd編集_Click
' ================================================================
Private Function PW認証チェック() As Boolean

    PW認証チェック = False

    Dim strBumon As String
    strBumon = Trim$(職員情報Key.所属部門)

    ' 所属部門が "3" 以下（一般職員）は編集権限なし
    If strBumon <= "3" Then
        MsgBox "編集・登録権限がありません。", vbExclamation, cstSys
        Exit Function
    End If

    ' ── すでにこのセッションで認証済みならスキップ ──────────────────
    '   flgPwOk >= 1 かつ 所属部門レベルに合致していれば認証不要
    If strBumon >= "6" Then
        If flgPwOk >= 2 Then
            PW認証チェック = True
            Exit Function
        End If
    ElseIf strBumon >= "4" Then
        If flgPwOk >= 1 Then
            PW認証チェック = True
            Exit Function
        End If
    End If

    ' ── パスワード入力 ───────────────────────────────────────────────
    Dim strInput As String
    strInput = InputBoxDK("編集・登録パスワードを入力してください", "認証")

    If strBumon >= "6" Then
        ' システム管理者パスワード
        If strInput = cstPwSys Then
            flgPwOk = 2
            flgHyouji = 2
            flgSYS = 1
            PW認証チェック = True
        Else
            MsgBox cstMsg05, vbExclamation, cstSys
        End If
    ElseIf strBumon >= "4" Then
        ' 中間管理者パスワード
        If strInput = cstPwJimu Then
            flgPwOk = 1
            flgHyouji = 1
            PW認証チェック = True
        Else
            MsgBox cstMsg05, vbExclamation, cstSys
        End If
    End If

End Function