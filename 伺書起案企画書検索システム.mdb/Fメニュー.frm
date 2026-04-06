Version =19
VersionRequired =19
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim strSQL     As String

Private Sub cmdコード管理_Click()

    DoCmd.OpenForm "Fコード検索"
    DoCmd.Close acForm, "Fメニュー"

End Sub

Private Sub cmdフォルダ_Click()

    On Error GoTo err_handle
    
    Dim strNen    As String
    Dim strPath1, strPath2  As String

    '年度取得
    strNen = Format(Date, "ggge年")
    
        strPath1 = "\\flsv1\fsroot\みのり苑\総務data\伺い書検索" & "\" & strNen & "度_PDF"  'フォルダパス名：本来の保存場所
'        strPath1 = "C:\Ts-002\総務data\伺い書検索" & "\" & strNen & "度_PDF"  'フォルダパス名：メンテ用
        strPath2 = "\\flsv1\fsroot\みのり苑\総務data\企画書検索" & "\" & strNen & "度_PDF"  'フォルダパス名：本来の保存場所
'        strPath2 = "C:\Ts-002\総務data\企画書検索" & "\" & strNen & "度_PDF"  'フォルダパス名：メンテ用

    If Dir(strPath1, vbDirectory) = vbNullString Then
        MkDir strPath1
        MsgBox "伺い書フォルダの作成が完了しました"
    Else
        MsgBox "伺い書フォルダすでに作成されています"
    End If
    If Dir(strPath2, vbDirectory) = vbNullString Then
        MkDir strPath2
        MsgBox "起案・企画書フォルダの作成が完了しました"
    Else
        MsgBox "起案・企画書フォルダすでに作成されています"
    End If
    
    Exit Sub

err_handle:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "フォルダ作成エラー"
    Debug.Print Err.Description

End Sub

Private Sub cmdログイン管理_Click()

    Call GetTableDWH("Tログイン伺企")
    DoCmd.OpenForm "Fログイン管理"
    DoCmd.Close acForm, "Fメニュー"

End Sub

Private Sub cmdロック管理_Click()

    Call GetTableDWH("Tロック伺企")
    DoCmd.OpenForm "Fロック管理"
    DoCmd.Close acForm, "Fメニュー"

End Sub

Private Sub cmd企画書_Click()

    flgSyubetu = 2
    
    DoCmd.OpenForm "Fメイン"
    DoCmd.Close acForm, "Fメニュー"
    
End Sub

Private Sub cmd伺い書_Click()

    flgSyubetu = 1
    
    DoCmd.OpenForm "Fメイン"
    DoCmd.Close acForm, "Fメニュー"
    
End Sub

Private Sub cmd終了_Click()

    Dim aaa As String
    Dim DB As DAO.Database
    Dim RRS As DAO.Recordset

    Set DB = CurrentDb
    Set RRS = DB.OpenRecordset("Tログイン伺企", dbOpenSnapshot)

    RRS.FindFirst "職員番号=" & TBLログイン.職員番号

    If RRS.NoMatch = False Then
        
        If IsNull(TBLログイン.職員番号) = False Then
            'ログイン情報削除
            TBLログイン_INIT
            TBLログイン.職員番号 = 職員情報Key.職員番号
            intRtn = ログイン_DEL
        End If
    End If
    
    RRS.Close: Set RRS = Nothing

'    システム終了
    DoCmd.Quit

End Sub

Private Sub cmd職員管理_Click()

    DoCmd.OpenForm "F職員管理"
    DoCmd.Close acForm, "Fメニュー"

End Sub

Private Sub Form_Open(Cancel As Integer)
    
    If flgHyouji = 1 Or flgHyouji = 2 Then
        Me.cmdログイン管理.Visible = True
        Me.cmdログイン管理.Enabled = True
        Me.cmdロック管理.Visible = True
        Me.cmdロック管理.Enabled = True
    Else
        cmdログイン管理.Enabled = False
        cmdロック管理.Enabled = False
    End If
    If flgSYS = 1 Then
        Me.cmdログイン管理.Visible = True
        Me.cmdログイン管理.Enabled = True
        Me.cmdロック管理.Visible = True
        Me.cmdロック管理.Enabled = True
        Me.cmdコード管理.Visible = True
        Me.cmdコード管理.Enabled = True
        Me.cmd職員管理.Visible = True
        Me.cmd職員管理.Enabled = True
    End If
    
End Sub

Private Sub lblMタイトル_DblClick(Cancel As Integer)

    With Me
        .cmdログイン管理.Enabled = True
        .cmdロック管理.Enabled = True
        .cmdコード管理.Enabled = True
        .cmd職員管理.Enabled = True
    End With
    
End Sub