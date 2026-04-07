Attribute VB_Name = "Mdl共通"
Option Compare Database


' ★ DB接続パスの一元管理
' 本番環境と開発環境を切り替える場合はここだけ変更する
' 共有DB設定は Mdl共通.GetTableDWH(),データ書込() 関数内でも設定
' そちらも変更
' ================================================================
#Const IS_TEST = True

#If IS_TEST Then
    Public Const DB_PATH        As String = "\\flsv1\fsroot\towada\福祉の里\DataBase\STAGE_伺書起案企画書DB.mdb"
    Public Const DB_SERVER_PATH As String = DB_PATH   ' テスト時はローカルと同じ
    Public Const STATUS As String = "DEV"
#Else
    Public Const DB_PATH        As String = "\\flsv1\fsroot\towada\福祉の里\DataBase\伺書起案企画書DB.mdb"
    Public Const DB_SERVER_PATH As String = "\\flsv1\fsroot\towada\福祉の里\DataBase\伺書起案企画書DB.mdb"
    Public Const STATUS As String = "RELEASE"
#End If


'========================データベース用変数
Public daoWS              As DAO.Workspace
Public daoDB              As DAO.Database
Public daoRS              As DAO.Recordset
Public daoQDF             As DAO.QueryDef
Public strSQL          As String
Public intRS           As Integer

'========================
Public strNichiji      As String
Public flgForm         As Integer
Public flgShinki       As Integer
Public flgChk          As Integer
Public flgHaita        As Integer
Public flgChusyutu     As Integer
Public flgSyubetu      As String
Public flgSYS          As Integer
Public strNendo        As String
Public strShisetuEx    As String
Public strSyozokuEx    As String
Public strShimeiEx     As String
Public strKenmeiEx     As String
Public strKaishiEx     As String
Public strSyuryoEx     As String
Public strFname        As String
Public flgOwari        As Integer
Public flgHyouji       As Integer
Public flgEditTransit  As Integer
Public strQry          As String
Public strTbl          As String

' ── 既存の変数宣言に以下を追加 ──────────────────────────────────
Public flgPwOk      As Integer   ' ★ 追加: 0=PW未認証, 1=中間管理PW認証済, 2=SYS管理PW認証済
Public strNextBangou   As String   ' 新規登録時の次番号（受付年度()から受け取る）

'システム名
Public Const cstSys = "伺い書、起案・企画書ＤＢシステム"
Public Const cstVersion = "20260403"
'エラーメッセージ
Public Const cstMsg01 = "職員を選択してください"
Public Const cstMsg02 = "職員が選択されていません"
Public Const cstMsg03 = "選択された職員はログインしています"
Public Const cstMsg04 = "メニューを選択してください"
Public Const cstMsg05 = "パスワードが違います"

'プレビュー出力先
Public Const cstPrintPath = "\Temp.snp"

'========================データベース用変数(ADO接続)

Public cn               As ADODB.Connection
Public rs               As ADODB.Recordset
Public RSF              As ADODB.Recordset

Public PintKBN          As Integer
Public pstrSisetucho    As String
Public PlngNintei       As Long

Public Const DB_OK = 0
Public Const DB_ERR = 9
Public Const DB_EOF = 1

Public Const RTN_OK = 0
Public Const RTN_ERR = -1

' オープン時に、Access自身を最小化に設定する為の　API 宣言
' 2003/05/10 pPoy
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, _
    ByVal nCmdShow As Long) As Long
'''Declare Function ShowWindow Lib "User32" _
'''      (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

' ウィンドウの表示状態を指定する定数の宣言
Public Const SW_SHOWMINIMIZED = 2   ' ウィンドウをアクティブにして最小化
Public Const SW_MINIMIZE = 6        ' ウィンドウを最小化
Public Const SW_HIDE = 0            ' ウィンドウを非表示

Type 端末情報Key
    処理端末            As String * 10
End Type
Public 端末情報Key      As 端末情報Key

Public Function EscapeSqlText(ByVal strValue As String) As String

    EscapeSqlText = Replace(Nz(strValue, ""), "'", "''")

End Function

Public Function 設定値取得(ByVal strKey As String, Optional ByVal strDefault As String = "") As String

On Error GoTo 設定値取得_ERR

    Dim varValue As Variant
    Dim strEnvName As String
    Dim strEnvValue As String
    Dim strFallback As String

    Select Case UCase$(Trim$(strKey))
        Case "PW_JIMU"
            strEnvName = "UKAGAI_PW_JIMU"
            strFallback = "jimu1319s"
        Case "PW_SYS"
            strEnvName = "UKAGAI_PW_SYS"
            strFallback = "sys0120310272"
        Case Else
            strFallback = strDefault
    End Select

    If strDefault <> "" Then
        設定値取得 = strDefault
    Else
        設定値取得 = strFallback
    End If

    If strEnvName <> "" Then
        strEnvValue = Trim$(Environ$(strEnvName))
        If strEnvValue <> "" Then
            設定値取得 = strEnvValue
            Exit Function
        End If
    End If

    If DCount("*", "MSysObjects", "Name='Tシステム設定' AND Type In (1,4,6)") = 0 Then
        Exit Function
    End If

    varValue = DLookup("設定値", "Tシステム設定", "設定キー = '" & EscapeSqlText(strKey) & "'")
    If Not IsNull(varValue) Then
        設定値取得 = CStr(varValue)
    End If

    Exit Function

設定値取得_ERR:
    ' 設定テーブルが無い環境では既定値または旧運用値にフォールバック

End Function

' CN_INIT を修正（Sub内のPublic Const宣言を削除する）
Sub CN_INIT(Optional ByRef intSts As Integer)
On Error GoTo CN_INIT_ERR
    intSts = DB_ERR
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_PATH
    intSts = DB_OK
    Exit Sub
CN_INIT_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ＤＢ接続エラー"
End Sub

Sub CN_END(Optional ByRef intSts As Integer)

On Error GoTo CN_END_ERR
    
    intSts = DB_ERR
    
    If Not cn Is Nothing Then
        cn.Close
        Set cn = Nothing
    End If

    intSts = DB_OK
    
    Exit Sub

CN_END_ERR:

End Sub

Sub RS_INIT(Optional ByRef intSts As Integer)

On Error GoTo RS_INIT_ERR

    intSts = DB_ERR
    
    Set rs = New ADODB.Recordset
    
    intSts = DB_OK
    
    Exit Sub

RS_INIT_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ＳＱＬ発行エラー"

End Sub

Sub RS_INIT2(Optional ByRef intSts As Integer)

On Error GoTo RS_INIT2_ERR
    
    intSts = DB_ERR
    
    Set rs = daoDB.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)
    
    intSts = DB_OK
    
    Exit Sub

RS_INIT2_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ＳＱＬ発行エラー"

End Sub

Sub RS_END(Optional ByRef intSts As Integer)

On Error GoTo RS_END_ERR
    
    intSts = DB_ERR
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    intSts = DB_OK
    
    Exit Sub

RS_END_ERR:

End Sub

Sub CN_EXEC(Optional ByRef intSts As Integer)

On Error GoTo CN_EXEC_ERR

    intSts = DB_ERR
    
    cn.Execute strSQL

    intSts = DB_OK
    
    Exit Sub

CN_EXEC_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ＳＱＬ実行エラー"

End Sub

Sub SDET(TableName, Filed, NO)
'/////テーブルの特定データ削除

On Error GoTo ERR_SDET
    Dim rsTbl As New ADODB.Recordset
    rsTbl.Open TableName, CurrentProject.Connection, adOpenStatic, adLockOptimistic
    
    If rsTbl.RecordCount > 0 Then
        rsTbl.MoveLast
        rsTbl.MoveFirst
        rsTbl.Find "" & [Filed] & " = '" & NO & "'"
        If rsTbl.EOF Then
            rsTbl.Close
            Set rsTbl = Nothing
            Exit Sub
        Else
            Do Until rsTbl.EOF
            rsTbl.Delete
            rsTbl.MoveNext
        rsTbl.Find "" & [Filed] & " = '" & NO & "'"
            Loop
        End If
    End If
        rsTbl.Close
        Set rsTbl = Nothing
        
Exit_SDET:
    Exit Sub
ERR_SDET:
    MsgBox Err.Description
    Resume Exit_SDET
End Sub

Function LowCount(ByVal strMoji As String, ByVal intPara As Integer) As Integer

    Dim intRtn As Integer
    Dim intIdx As Integer
    Dim varLow As Variant
    
    LowCount = 0
    intRtn = 1
    intIdx = 1
    
    '改行有無判定
    Do Until intRtn = 0
        intRtn = InStr(intIdx, strMoji, vbCrLf)
        intIdx = intRtn + 1
        LowCount = LowCount + 1
    Loop

    '改行単位に分解
    varLow = Split(strMoji, vbCrLf)
    '1行のバイト数を超えている場合、行数再計算
    For intIdx = 0 To UBound(varLow)
        If LenB(StrConv(varLow(intIdx), vbFromUnicode)) > intPara Then
            LowCount = LowCount + (LenB(StrConv(varLow(intIdx), vbFromUnicode)) \ intPara)
            If LenB(StrConv(varLow(intIdx), vbFromUnicode)) Mod intPara = 0 Then
                LowCount = LowCount - 1
            End If
        End If
    Next
    
End Function

' --------------------------------------------------------
' Server のテーブルを、端末にコピーする
' --------------------------------------------------------
Public Sub GetTableDWH(strTbl As String)

    Dim strFd As String
    
    ' Serverから読み込む側の設定
    Dim cn_R As Object 'ADOコネクションオブジェクト
    Set cn_R = CreateObject("ADODB.Connection") 'ADOコネクションのオブジェクトを作成
    
    strFd = DB_SERVER_PATH
    cn_R.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFd

    ' ローカルAccessへ書き込む側の設定
    Dim cn_W As Object 'ADOコネクションオブジェクト
    Set cn_W = CurrentProject.Connection
    
    ' テーブルを読み込む処理（同じテーブル名での読み書きを想定）
    Call TableImport(strTbl, cn_R, cn_W)

    ' 読み書きの設定を終了する
    cn_R.Close: Set cn_R = Nothing 'コネクションの破棄
    Set cn_W = Nothing

'    MsgBox "取り込み処理が完了しました！" & vbCrLf & vbCrLf & strMessage
    
End Sub
' --------------------------------------------------------
' Server のテーブルを、端末ACCESSにコピーする
' --------------------------------------------------------
' ================================================================
' 修正版 TableImport
' 問題: Exit Sub がエラーハンドラより前にあり、コピー処理が動いていない
' ================================================================
Private Sub TableImport(strTableName As String, ByRef objCn_R As Object, ByRef objCn_W As Object)
    Dim strSQL_R As String
    Dim i As Integer
    Dim rs_R As Object
    Dim rs_W As Object

On Error GoTo TableImport_ERR

    Set rs_R = CreateObject("ADODB.Recordset")
    Set rs_W = CreateObject("ADODB.Recordset")

    ' Step1: ローカルテーブルを全件削除
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM " & strTableName
    DoCmd.SetWarnings True

    ' Step2: サーバーから全件読み込み
    strSQL_R = "SELECT * FROM " & strTableName
    rs_R.Open strSQL_R, objCn_R

    ' Step3: ローカルテーブルに書き込み
    rs_W.Open strTableName, objCn_W, 1, 2  ' adOpenKeyset, adLockOptimistic

    Do Until rs_R.EOF
        rs_W.AddNew
        For i = 0 To rs_R.Fields.Count - 1
            rs_W.Fields(i).Value = rs_R.Fields(i).Value
        Next i
        rs_W.Update
        rs_R.MoveNext
    Loop

    rs_R.Close: Set rs_R = Nothing
    rs_W.Close: Set rs_W = Nothing
    DoCmd.SetWarnings True
    Exit Sub

TableImport_ERR:
    DoCmd.SetWarnings True
    If Not rs_R Is Nothing Then rs_R.Close: Set rs_R = Nothing
    If Not rs_W Is Nothing Then rs_W.Close: Set rs_W = Nothing
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "TableImport エラー"

End Sub

Sub データ書込(ByVal strQry As String, ByVal strTbl As String)

    On Error GoTo データ書込_Err
    
    DoCmd.TransferDatabase acImport, "Microsoft Access", DB_SERVER_PATH, acTable, strQry, strTbl, False
データ書込_Exit:
    Exit Sub
データ書込_Err:
    MsgBox Error$
    Resume データ書込_Exit
End Sub

Sub データ削除(TableName, Filed, Name)
'/////テーブルの特定データ削除

On Error GoTo ERR_SDET
    Dim rsTbl As New ADODB.Recordset
    rsTbl.Open TableName, CurrentProject.Connection, adOpenStatic, adLockOptimistic
    
    If rsTbl.RecordCount > 0 Then
        rsTbl.MoveLast
        rsTbl.MoveFirst
        rsTbl.Find "" & [Filed] & " = '" & Name & "'"
        If rsTbl.EOF Then
            rsTbl.Close
            Set rsTbl = Nothing
            Exit Sub
        Else
            Do Until rsTbl.EOF
            rsTbl.Delete
            rsTbl.MoveNext
        rsTbl.Find "" & [Filed] & " = '" & Name & "'"
            Loop
        End If
    End If
        rsTbl.Close
        Set rsTbl = Nothing
        
Exit_SDET:
    Exit Sub
ERR_SDET:
    MsgBox Err.Description
    Resume Exit_SDET
End Sub

' 戻るボタン・×ボタンから共通で呼び出す
Public Sub 全ロック解放()
    Dim intRtn As Integer

    ' 1) 文書ロック（編集中レコードロック）を解放
    '    職員番号で登録されているロックを全削除
    Call TBLロック_INIT
    TBLロック.職員番号 = 職員情報Key.職員番号
    intRtn = ロック_DEL   ' Mdlロック.bas の既存関数

    ' 2) ログインロックを解放
    Call TBLログイン_INIT
    TBLログイン.職員番号 = 職員情報Key.職員番号
    intRtn = ログイン_DEL   ' Mdlログイン.bas の既存関数

    ' 3) PW認証フラグをリセット
    flgPwOk = 0

End Sub

Public Sub 安全ロック解放()

On Error Resume Next
    Call 全ロック解放

End Sub

