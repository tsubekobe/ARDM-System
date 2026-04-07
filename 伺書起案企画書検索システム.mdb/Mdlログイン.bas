Attribute VB_Name = "Mdlログイン"
Option Compare Database
Option Explicit

Type TBLログイン
    職員番号        As Long
    職員氏名        As String * 50
    所属部門        As Integer
    使用区分        As Integer
    処理端末        As String * 10
    処理日時        As String * 19
End Type
Public TBLログイン  As TBLログイン

Private intSts      As Integer

Sub TBLログイン_INIT()
    
    With TBLログイン
        .職員番号 = 0
        .職員氏名 = ""
        .所属部門 = 0
        .使用区分 = 0
        .処理端末 = ""
        .処理日時 = ""
    End With

End Sub

Function ログイン_SEL() As Integer

On Error GoTo ログイン_SEL_ERR
    
    ログイン_SEL = RTN_ERR
    
    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'SELECT文
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM Tログイン伺企"
    'WHERE句
    strSQL = strSQL & " WHERE 職員番号 = " & CStr(TBLログイン.職員番号)
    'RSオープン
    Call RS_INIT(intSts)
    If intSts <> DB_OK Then
        GoTo ログイン_SEL_EXIT
    End If
    rs.Open strSQL, cn, adOpenStatic, adLockOptimistic
    
    'RSなし
    If rs.EOF Then
        ログイン_SEL = DB_EOF
        GoTo ログイン_SEL_EXIT
    Else
        TBLログイン.職員番号 = Nz(rs.Fields("職員番号").Value, 0)
        TBLログイン.職員氏名 = Nz(rs.Fields("職員氏名").Value, "")
        TBLログイン.所属部門 = Nz(rs.Fields("所属部門").Value, 0)
        TBLログイン.使用区分 = Nz(rs.Fields("使用区分").Value, 0)
        TBLログイン.処理端末 = Nz(rs.Fields("処理端末").Value, "")
        TBLログイン.処理日時 = Nz(rs.Fields("処理日時").Value, "")
    End If
    
    ログイン_SEL = RTN_OK
    
ログイン_SEL_EXIT:
    'RSクローズ
    Call RS_END
    'DBクローズ
    Call CN_END
    
    Exit Function

ログイン_SEL_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ログイン(SELECT)"
    GoTo ログイン_SEL_EXIT

End Function

Function ログイン_DEL() As Integer

On Error GoTo ログイン_DEL_ERR
    
    ログイン_DEL = RTN_ERR
    
    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'DELETE文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM Tログイン伺企"
    'WHERE句
    strSQL = strSQL & " WHERE 職員番号 = " & CStr(TBLログイン.職員番号)
    '実行
    Call CN_EXEC(intSts)
    If intSts <> DB_OK Then
        GoTo ログイン_DEL_EXIT
    End If

    ログイン_DEL = RTN_OK

ログイン_DEL_EXIT:
    'DBクローズ
    Call CN_END
    
    Exit Function

ログイン_DEL_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ログイン(DELETE)"
    GoTo ログイン_DEL_EXIT

End Function

Function ログイン_INS() As Integer

On Error GoTo ログイン_INS_ERR
    
    ログイン_INS = RTN_ERR
    
    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'INSERT文
    strSQL = ""
    strSQL = strSQL & " INSERT INTO Tログイン伺企"
    strSQL = strSQL & " VALUES("
    strSQL = strSQL & CStr(TBLログイン.職員番号) & ","
    strSQL = strSQL & "'" & EscapeSqlText(Trim$(TBLログイン.職員氏名)) & "',"
    strSQL = strSQL & CStr(TBLログイン.所属部門) & ","
    strSQL = strSQL & CStr(TBLログイン.使用区分) & ","
    strSQL = strSQL & "'" & EscapeSqlText(Trim$(TBLログイン.処理端末)) & "',"
    strSQL = strSQL & "#" & EscapeSqlText(Trim$(TBLログイン.処理日時)) & "#"
    strSQL = strSQL & " )"
    '実行
    Call CN_EXEC(intSts)
    If intSts <> DB_OK Then
        GoTo ログイン_INS_EXIT
    End If

    ログイン_INS = RTN_OK

ログイン_INS_EXIT:
    'DBクローズ
    Call CN_END
    
    Exit Function

ログイン_INS_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ログイン(INSERT)"
    GoTo ログイン_INS_EXIT

End Function
