Attribute VB_Name = "Mdl基本情報"
Option Compare Database
Private intSts      As Integer

Type TBL基本情報
    処理番号   As String * 7
    起案日     As String * 10
    種類       As String * 6
    施設       As String * 20
    所属       As String * 100
    起案者     As String * 50
    件名       As String
    PDFリンク  As String
    年度       As String * 4
    人事       As Boolean
    秘         As Boolean
    登録日時   As String * 20
End Type
Public TBL基本情報  As TBL基本情報

Sub 基本情報_INIT()
    With TBL基本情報
        .処理番号 = ""
        .起案日 = ""
        .種類 = ""
        .施設 = ""
        .所属 = ""
        .起案者 = ""
        .件名 = ""
        .PDFリンク = ""
        .年度 = ""
        .人事 = False
        .秘 = False
        .登録日時 = ""
    End With
End Sub

