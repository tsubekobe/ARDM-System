Attribute VB_Name = "Mdlロック"
Option Compare Database
Option Explicit

Type TBLロック
    伺企番号        As String * 10
    職員番号        As Long
    職員氏名        As String * 50
    処理端末        As String * 10
    処理日時        As String * 19
End Type
Public TBLロック  As TBLロック

Type アクセス情報key
    伺企番号        As String * 10
    職員氏名        As String * 50
    処理端末        As String * 10
End Type
Public アクセス情報key As アクセス情報key

Private intSts      As Integer

Sub TBLロック_INIT()
    
    With TBLロック
        .伺企番号 = ""
        .職員番号 = 0
        .職員氏名 = ""
        .処理端末 = ""
        .処理日時 = ""
    End With

End Sub

Function ロック_SEL() As Integer

On Error GoTo ロック_SEL_ERR
    
    ロック_SEL = RTN_ERR
    
    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'SELECT文
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM Tロック伺企"
    'WHERE句
    strSQL = strSQL & " WHERE 伺企番号 = '" & Trim$(TBLロック.伺企番号) & "'"
    
    'RSオープン
    Call RS_INIT(intSts)
    rs.Open strSQL, cn, adOpenStatic, adLockOptimistic
    If intSts <> DB_OK Then
        GoTo ロック_SEL_EXIT
    End If
    
    'RSなし
    If rs.EOF Then
        ロック_SEL = DB_EOF
        GoTo ロック_SEL_EXIT
    Else
        TBLロック.伺企番号 = Nz(rs.Fields("伺企番号").Value, "")
        TBLロック.職員番号 = Nz(rs.Fields("職員番号").Value, 0)
        TBLロック.職員氏名 = Nz(rs.Fields("職員氏名").Value, "")
        TBLロック.処理端末 = Nz(rs.Fields("処理端末").Value, "")
        TBLロック.処理日時 = Nz(rs.Fields("処理日時").Value, "")
    End If
    
    ロック_SEL = RTN_OK
    
ロック_SEL_EXIT:
    'RSクローズ
    RS_END
    'DBクローズ
    CN_END
    
    Exit Function

ロック_SEL_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ロック(SELECT)"
    GoTo ロック_SEL_EXIT

End Function

Function ロック_DEL() As Integer

On Error GoTo ロック_DEL_ERR
    
    ロック_DEL = RTN_ERR
    
    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'DELETE文
    strSQL = ""
    strSQL = strSQL & " DELETE FROM Tロック伺企"
    'WHERE句
    strSQL = strSQL & " WHERE 職員番号 = " & CInt(TBLロック.職員番号)
    '実行
    Call CN_EXEC(intSts)
    If intSts <> DB_OK Then
        GoTo ロック_DEL_EXIT
    End If

    ロック_DEL = RTN_OK

ロック_DEL_EXIT:
    'DBクローズ
    Call CN_END
    
    Exit Function

ロック_DEL_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ロック(DELETE)"
    GoTo ロック_DEL_EXIT

End Function

Function ロック_INS() As Integer

On Error GoTo ロック_INS_ERR
    
    ロック_INS = RTN_ERR
    
    'DBオープン
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'INSERT文
    strSQL = ""
    strSQL = strSQL & " INSERT INTO Tロック伺企"
    strSQL = strSQL & " VALUES("
    strSQL = strSQL & "'" & Trim$(TBLロック.伺企番号) & "',"
    strSQL = strSQL & CInt(TBLロック.職員番号) & ","
    strSQL = strSQL & "'" & Trim$(TBLロック.職員氏名) & "',"
    strSQL = strSQL & "'" & Trim$(TBLロック.処理端末) & "',"
    strSQL = strSQL & "#" & Trim$(TBLロック.処理日時) & "#"
    strSQL = strSQL & " )"
    '実行
    Call CN_EXEC(intSts)
    If intSts <> DB_OK Then
        GoTo ロック_INS_EXIT
    End If

    ロック_INS = RTN_OK

ロック_INS_EXIT:
    'DBクローズ
    Call CN_END
    
    Exit Function

ロック_INS_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ロック(INSERT)"
    GoTo ロック_INS_EXIT

End Function

' ── Mdlロック.bas に追加 ──
' 伺企番号キーでロックを削除する（職員番号問わず）
Function ロック_DEL_BY_KEY(strKey As String) As Integer

On Error GoTo ロック_DEL_BY_KEY_ERR

    ロック_DEL_BY_KEY = RTN_ERR

    Dim intS As Integer
    Call CN_INIT(intS)
    If intS <> DB_OK Then Exit Function

    strSQL = "DELETE FROM Tロック伺企 WHERE 伺企番号 = '" & strKey & "'"
    Call CN_EXEC(intS)
    If intS = DB_OK Then ロック_DEL_BY_KEY = RTN_OK

ロック_DEL_BY_KEY_EXIT:
    Call CN_END
    Exit Function

ロック_DEL_BY_KEY_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ロック削除(KEY)エラー"
    GoTo ロック_DEL_BY_KEY_EXIT

End Function

