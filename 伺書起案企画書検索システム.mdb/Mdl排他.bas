Attribute VB_Name = "Mdl排他"
Option Compare Database
Option Explicit

Type 排他情報Key
    伺企番号            As String * 10
    職員番号            As String * 10
    処理名              As String * 10
    メッセージ          As String
End Type
Public 排他情報Key      As 排他情報Key

Sub 排他情報Key_INIT()

    With 排他情報Key
        .伺企番号 = ""
        .職員番号 = ""
        .処理名 = ""
        .メッセージ = ""
    End With

End Sub

Sub 排他_DEL()

On Error GoTo 排他_DEL_ERR
    
    Dim intRtn As Integer
    
    'TBLロック初期化
    Call TBLロック_INIT
    TBLロック.職員番号 = 職員情報Key.職員番号
    'ロック(DELETE)
    intRtn = ロック_DEL

    Exit Sub

排他_DEL_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "排他解除エラー"

End Sub

Function 排他_CHK() As Boolean

On Error GoTo 排他_CHK_ERR
    
    Dim intRtn As Integer
    
    排他_CHK = False

    'TBLロック初期化
    Call TBLロック_INIT
    TBLロック.伺企番号 = 排他情報Key.伺企番号
    'ロック(SELECT)
    intRtn = ロック_SEL
    If intRtn = RTN_OK Then
        If TBLロック.職員番号 = 職員情報Key.職員番号 Then
            Exit Function
        End If
        排他_CHK = True
    End If

    'TBLロック初期化
    If 排他_CHK = True Then
        排他情報Key.メッセージ = 排他情報Key.メッセージ & "選択された伺企番号は、"
        排他情報Key.メッセージ = 排他情報Key.メッセージ & "他の職員により使用中です"
        排他情報Key.メッセージ = 排他情報Key.メッセージ & "（"
        排他情報Key.メッセージ = 排他情報Key.メッセージ & Trim$(TBLロック.職員氏名)
        排他情報Key.メッセージ = 排他情報Key.メッセージ & "）"
        flgOwari = 1
    Else
        Call TBLロック_INIT
        TBLロック.伺企番号 = 排他情報Key.伺企番号
        TBLロック.職員番号 = 職員情報Key.職員番号
        TBLロック.職員氏名 = Trim$(職員情報Key.職員氏名)
        TBLロック.処理端末 = Trim$(職員情報Key.処理端末)
        TBLロック.処理日時 = CStr(Now)
        intRtn = ロック_INS
        排他情報Key.メッセージ = "排他中"
    End If
        
    Exit Function

排他_CHK_ERR:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "排他チェックエラー"

End Function
