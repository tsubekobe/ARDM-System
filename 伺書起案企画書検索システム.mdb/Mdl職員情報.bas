Attribute VB_Name = "MdlEˆõî•ñ"
Option Compare Database
Option Explicit

Type Eˆõî•ñKey
    Eˆõ”Ô†            As Long
    Eˆõ–¼            As String * 50
    Š‘®•”–å            As String * 20
    ˆ—’[––            As String * 10
    g—p‹æ•ª            As Integer
End Type
Public Eˆõî•ñKey      As Eˆõî•ñKey

Sub Eˆõî•ñKey_INIT()

    With Eˆõî•ñKey
        .Eˆõ”Ô† = 0
        .Eˆõ–¼ = ""
        .Š‘®•”–å = ""
        .ˆ—’[–– = ""
        .g—p‹æ•ª = 0
    End With

End Sub


Function EˆõŠÇ—_SEL() As Integer

On Error GoTo EˆõŠÇ—_SEL_ERR

    Dim intSts As Integer
    
    EˆõŠÇ—_SEL = RTN_ERR

    'DBƒI[ƒvƒ“
    Call CN_INIT(intSts)
    If intSts <> DB_OK Then
        Exit Function
    End If
    
    'SELECT•¶
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM TEˆõŠÇ—"
    'WHERE‹å
    strSQL = strSQL & " WHERE ˆ—’[–– = '" & Trim$(Eˆõî•ñKey.ˆ—’[––) & "'"
    
    'RSƒI[ƒvƒ“
    Call RS_INIT(intSts)
    rs.Open strSQL, cn, adOpenStatic, adLockOptimistic
    If intSts <> DB_OK Then
        GoTo EˆõŠÇ—_SEL_EXIT
    End If
    
    'RS‚È‚µ
    If rs.EOF Then
        EˆõŠÇ—_SEL = DB_EOF
        GoTo EˆõŠÇ—_SEL_EXIT
    Else
        Eˆõî•ñKey.Eˆõ”Ô† = Nz(rs.Fields("Eˆõ”Ô†").Value, 0)
    End If
    
    EˆõŠÇ—_SEL = RTN_OK
    
EˆõŠÇ—_SEL_EXIT:
    'RSƒNƒ[ƒY
    Call RS_END
    'DBƒNƒ[ƒY
    Call CN_END
    
    Exit Function

EˆõŠÇ—_SEL_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "EˆõŠÇ—(SELECT)"
    GoTo EˆõŠÇ—_SEL_EXIT

End Function
