Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =10148
    DatasheetFontHeight =11
    ItemSuffix =88
    Left =5811
    Top =3026
    Right =24429
    Bottom =14117
    RecSrcDt = Begin
        0x1b36415d9252e640
    End
    Caption ="Declarations"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =4210752
            FontName ="Calibri"
            GridlineColor =10921638
            ForeTint =75.0
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackColor =14136213
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =14136213
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =15060409
            HoverTint =40.0
            PressedColor =9592887
            PressedShade =75.0
            HoverForeColor =4210752
            HoverForeTint =75.0
            PressedForeColor =4210752
            PressedForeTint =75.0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4649
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =57
                    Width =5416
                    Height =4536
                    FontSize =9
                    TabIndex =6
                    Name ="lbDictData"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2835"
                    AfterUpdate ="[Event Procedure]"
                    HorizontalAnchor =2
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =57
                    LayoutCachedWidth =5476
                    LayoutCachedHeight =4593
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5674
                    Top =1397
                    Width =4380
                    Height =1354
                    TabIndex =3
                    ForeColor =0
                    Name ="lbVariations"
                    RowSourceType ="Value List"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    HorizontalAnchor =1
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =5674
                    LayoutCachedTop =1397
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =2751
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5674
                    Top =60
                    Width =4380
                    Height =317
                    FontWeight =700
                    TabIndex =1
                    Name ="txtWord"
                    ControlSource ="=[lbDictData]"
                    Format ="@;;\"(select item)\""
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0xbfbfbf00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740057006f00720064005d0020004900730020004e0075006c00 ,
                        0x6c0000000000
                    End
                    GroupTable =1
                    HorizontalAnchor =1

                    LayoutCachedLeft =5674
                    LayoutCachedTop =60
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =377
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000bfbfbf00ffffff00110000005b00 ,
                        0x74007800740057006f00720064005d0020004900730020004e0075006c006c00 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =85
                    Width =0
                    Height =0
                    Name ="Command15"
                    Caption ="sysFirst"

                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5674
                    Top =737
                    Width =4380
                    Height =300
                    TabIndex =2
                    Name ="txtCurrentLetterCase"
                    Format ="@;;---"
                    ConditionalFormat = Begin
                        0x01000000dc000000020000000100000000000000000000001200000001000000 ,
                        0xbfbfbf00ffffff000100000000000000130000003d00000001010000ba141900 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740057006f00720064005d0020004900730020004e0075006c00 ,
                        0x6c000000000053007400720043006f006d00700028005b007400780074005700 ,
                        0x6f00720064005d002c005b00740078007400430075007200720065006e007400 ,
                        0x560061006c00750065005d002c00300029003c003e00300000000000
                    End
                    GroupTable =1
                    TopPadding =0
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    HorizontalAnchor =1

                    LayoutCachedLeft =5674
                    LayoutCachedTop =737
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =1037
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000bfbfbf00ffffff00110000005b00 ,
                        0x74007800740057006f00720064005d0020004900730020004e0075006c006c00 ,
                        0x0000000000000000000000000000000000000000000100000000000000010100 ,
                        0x00ba141900ffffff002900000053007400720043006f006d00700028005b0074 ,
                        0x007800740057006f00720064005d002c005b0074007800740043007500720072 ,
                        0x0065006e007400560061006c00750065005d002c00300029003c003e00300000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5674
                    Top =437
                    Width =4380
                    Height =300
                    Name ="Label19"
                    Caption ="Current lettercase:"
                    GroupTable =1
                    BottomPadding =0
                    HorizontalAnchor =1
                    LayoutCachedLeft =5674
                    LayoutCachedTop =437
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =737
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5674
                    Top =1097
                    Width =4380
                    Height =300
                    Name ="Label20"
                    Caption ="Variations:"
                    GroupTable =1
                    BottomPadding =0
                    HorizontalAnchor =1
                    LayoutCachedLeft =5674
                    LayoutCachedTop =1097
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =1397
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5674
                    Top =3111
                    Width =3403
                    Height =343
                    TabIndex =4
                    Name ="txtSelectedLetterCase"
                    ControlSource ="=[lbVariations]"
                    Format ="@;;---"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0xbfbfbf00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740057006f00720064005d0020004900730020004e0075006c00 ,
                        0x6c0000000000
                    End
                    GroupTable =1
                    TopPadding =0
                    HorizontalAnchor =1

                    LayoutCachedLeft =5674
                    LayoutCachedTop =3111
                    LayoutCachedWidth =9077
                    LayoutCachedHeight =3454
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000bfbfbf00ffffff00110000005b00 ,
                        0x74007800740057006f00720064005d0020004900730020004e0075006c006c00 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5674
                    Top =2777
                    Width =4380
                    Height =300
                    Name ="Label45"
                    Caption ="Change to:"
                    GroupTable =1
                    TopPadding =0
                    HorizontalAnchor =1
                    LayoutCachedLeft =5674
                    LayoutCachedTop =2777
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =3077
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =9137
                    Top =3111
                    Width =917
                    Height =343
                    TabIndex =5
                    Name ="cmdChangeLetterCase"
                    Caption ="Commit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    HorizontalAnchor =1

                    LayoutCachedLeft =9137
                    LayoutCachedTop =3111
                    LayoutCachedWidth =10054
                    LayoutCachedHeight =3454
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5555
                    Top =4370
                    Width =4550
                    Height =223
                    FontSize =8
                    Name ="lblVersionInfo"
                    HorizontalAnchor =2
                    LayoutCachedLeft =5555
                    LayoutCachedTop =4370
                    LayoutCachedWidth =10105
                    LayoutCachedHeight =4593
                End
            End
        End
    End
End
CodeBehindForm
' See "DeclarationDictApiDialog.cls"
