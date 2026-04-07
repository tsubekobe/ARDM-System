Attribute VB_Name = "Mdlファイル登録・選択"
Option Compare Database
Option Explicit

Public Sub PDF登録()

On Error GoTo Err_PDFTouroku_Click

    Dim mojisuu As Integer
    Dim mojinuki As Integer
    Dim Rmoji   As String
    Dim Lmoji   As String
    Dim strSrc  As String
'    Dim strFname As String
        
    strSrc = ファイル選択
    mojisuu = Len(strSrc)
    
    If mojisuu = 0 Then
        strFname = ""
    Else
        mojinuki = mojisuu - 8
        Lmoji = StrConv(Left(strSrc, 8), vbLowerCase) '小文字に揃える処理
        Rmoji = Mid(strSrc, 9, mojinuki)
        strFname = Lmoji & Rmoji
    End If
        
Exit_PDFTouroku_Click:
        Exit Sub
        
Err_PDFTouroku_Click:
        MsgBox Err.Description
        Resume Exit_PDFTouroku_Click
        
End Sub

Function ファイル選択()

On Error GoTo ErrorHandler

'    Dim returnValue As Variant
'    Dim strmsg As String
'    returnValue = SysCmd(acSysCmdAccessVer)
'    strmsg = "Access2002、2003でないため、この機能を利用できません。"
    
    'Accessのバージョンを調べます。
    'Access2000は10.0、Access2000は9.0,Access97は8.0,Access95は7.0を返します。
    
'    If returnValue = "10.0" Or returnValue = "11.0" Or returnValue = "12.0" Or returnValue = "13.0" Then
        
    Dim intType As Integer
    Dim varSelectedFile As Variant
    
    'ファイルを選択する場合は、msofiledialogfilepicker
    intType = msoFileDialogFilePicker
    
    'ファイル参照用の設定値をセットします。
    With Application.FileDialog(intType)
    
        'ダイアログタイトル名
        .Title = "ファイル選択"
        
        'ファイルの種類を定義します。
        .Filters.Add "PDF ファイル", "*.PDF"
        .Filters.Add "pdfファイル", "*.pdf"
        .Filters.Add "すべてのファイル", "*.*"
        
        '複数ファイル選択を可能にする場合はTrue、不可の場合はFalse。
        .AllowMultiSelect = False
        
        '最初に開くホルダーを当ファイルのフォルダーとします。
        If flgSyubetu = 1 Then
            .InitialFileName = "\\flsv1\fsroot\みのり苑\総務data\伺い書検索\"   'フォルダパス名：本来の保存元
'            .InitialFileName = "C:\Ts-002\総務data\伺い書検索\"  'フォルダパス名：テスト用
        ElseIf flgSyubetu = 2 Then
            .InitialFileName = "\\flsv1\fsroot\みのり苑\総務data\企画書検索\"   'フォルダパス名：本来の保存元
'            .InitialFileName = "C:\Ts-002\総務data\伺い書検索\"  'フォルダパス名：テスト用
        End If
        
    
        If .Show = -1 Then 'ファイルが選択されれば　-1 を返します。
            For Each varSelectedFile In .SelectedItems
                ファイル選択 = varSelectedFile
            Next
        Else
             'キャンセルが押されたときにエラーなく抜けるため
            ファイル選択 = ""
        End If
    
    End With
        
'    Else
'
'        MsgBox strmsg, vbOKOnly, "Microsoft Access Club"
'
'    End If
    
Exit Function

ErrorHandler:

    MsgBox "予期せぬエラーが発生しました" & Chr(13) & _
            "エラーナンバー：" & Err.Number & Chr(13) & _
            "エラー内容：" & Err.Description, vbOKOnly
    End
    
End Function



