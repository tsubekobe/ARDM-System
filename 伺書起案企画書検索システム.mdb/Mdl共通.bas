Attribute VB_Name = "Mdl共通"
Option Compare Database

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
Public strQry          As String
Public strTbl          As String

'システム名
Public Const cstSys = "伺い書、起案・企画書ＤＢシステム"
'エラーメッセージ
Public Const cstMsg01 = "職員を選択してください"
Public Const cstMsg02 = "職員が選択されていません"
Public Const cstMsg03 = "選択された職員はログインしています"
Public Const cstMsg04 = "メニューを選択してください"
Public Const cstMsg05 = "パスワードが違います"

'プレビュー出力先
Public Const cstPrintPath = "\Temp.snp"

'========================データベース用変数(ADO接続)

Public cn               As New ADODB.Connection
Public rs               As New ADODB.Recordset
Public RSF              As New ADODB.Recordset

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

Sub CN_INIT(Optional ByRef intSts As Integer)

    Dim strFd As String

On Error GoTo CN_INIT_ERR
    
    intSts = DB_ERR
    
    Set cn = New ADODB.Connection
    strFd = "\\flsv1\fsroot\towada\福祉の里\DataBase\伺書起案企画書DB.mdb"
'    strFd = "z:\DataBase\伺書起案企画書DB.mdb"
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFd

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

    Dim strSQL_R As String
    Dim strMessage As String
    Dim strFd As String
    
    ' Serverから読み込む側の設定
    Dim cn_R As Object 'ADOコネクションオブジェクト
    Set cn_R = CreateObject("ADODB.Connection") 'ADOコネクションのオブジェクトを作成
     'テスト環境
'    strFd = "z:\DataBase\伺書起案企画書DB.mdb"
'    本番環境
    strFd = "\\flsv1\fsroot\towada\福祉の里\DataBase\伺書起案企画書DB.mdb"
    cn_R.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFd

    ' ローカルAccessへ書き込む側の設定
    Dim cn_W As Object 'ADOコネクションオブジェクト
    Set cn_W = CurrentProject.Connection
    
    ' テーブルを読み込む処理（同じテーブル名での読み書きを想定）
    Call TableImport(strTbl, cn_R, cn_W)

    ' 読み書きの設定を終了する
    cn_R.Close: Set cn_R = Nothing 'コネクションの破棄
    cn_W.Close: Set cn_W = Nothing 'コネクションの破棄

'    MsgBox "取り込み処理が完了しました！" & vbCrLf & vbCrLf & strMessage
    
End Sub

' --------------------------------------------------------
' Server のテーブルを、端末ACCESSにコピーする
' --------------------------------------------------------
Private Sub TableImport(strTableName As String, ByRef objCn_R As Object, ByRef objCn_W As Object)
    Dim strSQL_R As String
    Dim i As Integer

    ' SQLServerから読み込む側の設定
    Dim rs_R As Object 'ADOレコードセットオブジェクト
    Set rs_R = CreateObject("ADODB.Recordset") 'ADOレコードセットのオブジェクトを作成

    ' Accessへ書き込む側の設定
    Dim rs_W As Object 'ADOレコードセットオブジェクト
    Set rs_W = CreateObject("ADODB.Recordset") 'ADOレコードセットのオブジェクトを作成
    
    ' 削除メッセージを出さなくして、Access上のテーブル内容を全件削除
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM " & strTableName
    
    ' Access側のテーブルが、クリアされて無かったら、エラー表示
    If DCount("*", strTableName) > 0 Then
        MsgBox "Access側の「" & strTableName & "」テーブルの内容、削除しきれてないでー"
    End If
        
    ' SQLserver側のテーブル内容を全件持ってくるため、読み込み、書き込みのオープン
    strSQL_R = "SELECT * FROM " & strTableName
    rs_R.Open strSQL_R, objCn_R
    rs_W.Open strTableName, objCn_W, 1, 2
    
    ' レコードごとに読み書きし、全レコードをなめる
    Do Until rs_R.EOF
        rs_W.AddNew
        For i = 0 To rs_R.Fields.Count - 1
            rs_W.Fields(i).Value = rs_R.Fields(i).Value
        Next i
        rs_W.Update
        rs_R.MoveNext
    Loop
    
     'レコードセットの破棄
    rs_R.Close: Set rs_R = Nothing
    rs_W.Close: Set rs_W = Nothing

    ' メッセージを出すモードに戻す
    DoCmd.SetWarnings True

End Sub

Sub データ書込(ByVal strQry As String, ByVal strTbl As String)

    On Error GoTo データ書込_Err
    
     'テスト環境
'    DoCmd.TransferDatabase acImport, "Microsoft Access", "z:\DataBase\伺書起案企画書DB.mdb", acTable, strQry, strTbl, False
'    本番環境
    DoCmd.TransferDatabase acImport, "Microsoft Access", "\\flsv1\fsroot\towada\福祉の里\DataBase\伺書起案企画書DB.mdb", acTable, strQry, strTbl, False
    
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

