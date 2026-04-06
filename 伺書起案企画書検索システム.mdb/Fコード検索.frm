Version =19
VersionRequired =19
Checksum =99685246
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularCharSet =128
    TabularFamily =50
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13041
    DatasheetFontHeight =11
    ItemSuffix =38
    Left =4410
    Top =2820
    Right =17250
    Bottom =11145
    RecSrcDt = Begin
        0xdabc840d9d17e440
    End
    GUID = Begin
        0x72b21ccc0733774ea6adb95fe54e7f3b
    End
    NameMap = Begin
        0x0acc0e55000000005fef0cc3ded00842b4ce633f88042f67000000009fabad58 ,
        0xce2ae640006e5003045fcf005400b330fc30c930a17b0674000000000000821e ,
        0x947ec4176347b5236c9020f432d3070000005fef0cc3ded00842b4ce633f8804 ,
        0x2f67b330fc30c930000000000000d461ab6deb9f3d4ba1bc98e01d2826550700 ,
        0x00005fef0cc3ded00842b4ce633f88042f67b330fc30c9304900440000000000 ,
        0x000059a9f9f8735ea749af8b546d95d274b5070000005fef0cc3ded00842b4ce ,
        0x633f88042f67b330fc30c9300d54f07900000000000000000000000000000000 ,
        0x0000000000000c000000040000000000000000000000000000000000
    End
    RecordSource ="SELECT Tコード管理.* FROM Tコード管理 WHERE ((([Tコード管理].コードID)=9999)) ORDER BY [Tコード管理].コー"
        "ド; "
    DatasheetFontName ="ＭＳ Ｐゴシック"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x010400069c00540303ff0000010009009a0b3408640001000700580201000100 ,
        0x5802030001004134000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000010000000000000001000000 ,
        0x0200000001000000000000000000000000000000000000000000000050524956 ,
        0x4230000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000180000000000102710271027 ,
        0x0000102700000000000000008000540300000000000000000000000000000000 ,
        0x0000000000000000030000000000000000001000503403002888040000000000 ,
        0x000000000000010000000000000000000000000000000000e7b14b4c03000000 ,
        0x05000a00ff000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0100000000000000000000000000000080000000534d544a0000000010007000 ,
        0x430075006200650050004400460000005265736f6c7574696f6e003630306470 ,
        0x69005061676553697a650041340050616765526567696f6e0000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x00000000000000000000000000000000
    End
    PrtDevNames = Begin
        0x080010001800010000000000000000000000000000000000437562655044463a ,
        0x00
    End
    OnLoad ="[Event Procedure]"
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =128
            FontSize =11
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontCharSet =128
            Width =1701
            Height =435
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =128
            Width =1701
            Height =270
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin ComboBox
            TextFontCharSet =128
            Width =1701
            Height =270
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin Subform
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =1134
            BackColor =15395562
            Name ="フォームヘッダー"
            GUID = Begin
                0x9181ec6151ce0b46853763623ad74d1b
            End
            Begin
                Begin Label
                    SpecialEffect =4
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    TextFontFamily =49
                    Left =170
                    Top =170
                    Width =5670
                    Height =496
                    FontSize =20
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="lblタイトル"
                    Caption ="コード検索"
                    FontName ="ＭＳ ゴシック"
                    GUID = Begin
                        0x39023354239df04eb29ef8372c9511cf
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =49
                    Left =195
                    Top =840
                    Width =1266
                    Height =270
                    Name ="lblコード区分"
                    Caption ="コード区分"
                    FontName ="ＭＳ ゴシック"
                    GUID = Begin
                        0xf42691a6a8eaf94294c68ab7163ba641
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =49
                    Left =1521
                    Top =840
                    Width =4866
                    Height =270
                    Name ="lblコード名称"
                    Caption ="コード名称"
                    FontName ="ＭＳ ゴシック"
                    GUID = Begin
                        0xb393070f6742654bbff7cc88a0793df6
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =397
            BackColor =8421376
            Name ="詳細"
            GUID = Begin
                0x37fb4a55bddb6843b67477e28ba16871
            End
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontFamily =49
                    BackStyle =0
                    Left =195
                    Top =75
                    Width =1266
                    ForeColor =16777215
                    Name ="txtコード"
                    ControlSource ="コード"
                    FontName ="ＭＳ ゴシック"
                    GUID = Begin
                        0xc9b686722c909c4ba1788778da7421be
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontFamily =49
                    IMEMode =1
                    BackStyle =0
                    Left =1521
                    Top =75
                    Width =4866
                    TabIndex =1
                    ForeColor =16777215
                    Name ="txtコード名称"
                    ControlSource ="コード名称"
                    FontName ="ＭＳ ゴシック"
                    GUID = Begin
                        0x5fed5413b94d4949a11e1bb80a318bdb
                    End
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =1134
            BackColor =13434879
            Name ="フォームフッター"
            GUID = Begin
                0x8a79dce2cf2af642846aa07e2d8b1d82
            End
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =49
                    Left =851
                    Top =165
                    Width =1418
                    Height =397
                    Name ="cmd選択"
                    Caption ="選　択"
                    OnClick ="[Event Procedure]"
                    FontName ="ＭＳ ゴシック"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    GUID = Begin
                        0x3d47211044fd514ab264510f508d41b6
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =49
                    Left =10773
                    Top =165
                    Width =1418
                    Height =397
                    TabIndex =1
                    Name ="cmd戻る"
                    Caption ="戻　る"
                    OnClick ="[Event Procedure]"
                    FontName ="ＭＳ ゴシック"
                    GUID = Begin
                        0x5fe7cf73f6c7e847abc743fd338dea8f
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmd選択_Click()

'    コード管理Key.コード = Me.コード
    DoCmd.OpenForm "Fコード管理", , , "コードID = " & Me.txtコード
    DoCmd.Close acForm, "Fコード検索"

End Sub

Private Sub cmd戻る_Click()

    DoCmd.OpenForm "Fメニュー"
    DoCmd.Close acForm, "Fコード検索"

End Sub

Private Sub Form_Load()

    Me.cmd選択.SetFocus

End Sub