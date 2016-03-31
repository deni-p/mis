VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmProdEkAutoClave 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ekstraksi Di AutoClave"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProdEkAutoClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9270
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   9270
      TabIndex        =   11
      Top             =   0
      Width           =   9270
      Begin VB.TextBox lblEkstraksi 
         Appearance      =   0  'Flat
         DataField       =   "noekstraksi"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Tag             =   "EA"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         Height          =   675
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   8
         Tag             =   "ACID"
         Top             =   6410
         Width           =   6645
      End
      Begin TrueOleDBGrid80.TDBGrid tgAutoClave 
         Height          =   4365
         Left            =   120
         TabIndex        =   9
         Top             =   1995
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7699
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Prosedur"
         Columns(0).DataField=   "Prosedur"
         Columns(0).Group=   -1  'True
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Analisa"
         Columns(1).DataField=   "Analysis"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Kebutuhan"
         Columns(2).DataField=   "kebutuhan"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "ProsesID"
         Columns(3).DataField=   "ProsesID"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "ID"
         Columns(4).DataField=   "ID"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).SizeMode=   2
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=6165"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6085"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Merge=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=5450"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5371"
         Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=3466"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3387"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(24)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(30)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(1)._UserFlags=   0
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   688
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1)._GSX_SAVERECORDSELECTORS=   65562
         Splits(1).DividerColor=   14215660
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=5"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=6165"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=6085"
         Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(7)=   "Column(1).Width=5450"
         Splits(1)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(9)=   "Column(1)._WidthInPix=5371"
         Splits(1)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(11)=   "Column(2).Width=2699"
         Splits(1)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(13)=   "Column(2)._WidthInPix=2619"
         Splits(1)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(1)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(1)._ColumnProps(18)=   "Column(3).Visible=0"
         Splits(1)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(20)=   "Column(4).Width=2725"
         Splits(1)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
         Splits(1)._ColumnProps(23)=   "Column(4).Visible=0"
         Splits(1)._ColumnProps(24)=   "Column(4).Order=5"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DataView        =   2
         GroupByCaption  =   ""
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=67,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=69,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=71,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=82,.parent=67,.valignment=2"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=67"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=68"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=69"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=71"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=90,.parent=67"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=68"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=69"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=71"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=94,.parent=67"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=68"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=69"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=71"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=98,.parent=67"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=68"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=69"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=71"
         _StyleDefs(52)  =   "Splits(1).Style:id=13,.parent=1"
         _StyleDefs(53)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(54)  =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(55)  =   "Splits(1).FooterStyle:id=15,.parent=3"
         _StyleDefs(56)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(57)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(58)  =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(59)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(60)  =   "Splits(1).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(61)  =   "Splits(1).OddRowStyle:id=21,.parent=10"
         _StyleDefs(62)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(63)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(64)  =   "Splits(1).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(65)  =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(66)  =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(67)  =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(68)  =   "Splits(1).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(69)  =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(70)  =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(71)  =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(72)  =   "Splits(1).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(73)  =   "Splits(1).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(74)  =   "Splits(1).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(75)  =   "Splits(1).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(76)  =   "Splits(1).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(77)  =   "Splits(1).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(78)  =   "Splits(1).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(79)  =   "Splits(1).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(80)  =   "Splits(1).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(81)  =   "Splits(1).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(82)  =   "Splits(1).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(83)  =   "Splits(1).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(84)  =   "Named:id=33:Normal"
         _StyleDefs(85)  =   ":id=33,.parent=0"
         _StyleDefs(86)  =   "Named:id=34:Heading"
         _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   ":id=34,.wraptext=-1"
         _StyleDefs(89)  =   "Named:id=35:Footing"
         _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=36:Selected"
         _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=37:Caption"
         _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(95)  =   "Named:id=38:HighlightRow"
         _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=39:EvenRow"
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(99)  =   "Named:id=40:OddRow"
         _StyleDefs(100) =   ":id=40,.parent=33"
         _StyleDefs(101) =   "Named:id=41:RecordSelector"
         _StyleDefs(102) =   ":id=41,.parent=34"
         _StyleDefs(103) =   "Named:id=42:FilterBar"
         _StyleDefs(104) =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cmdRefLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8385
         Picture         =   "FrmProdEkAutoClave.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "BAHAN"
         Top             =   495
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtReaktor 
         Appearance      =   0  'Flat
         DataField       =   "reaktor"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   6420
         TabIndex        =   4
         Tag             =   "EA"
         Top             =   120
         Width           =   2340
      End
      Begin VB.TextBox txtTanki 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   10800
         TabIndex        =   15
         Tag             =   "ALKALI"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "group"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Tag             =   "EA"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtBerat 
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   11160
         TabIndex        =   14
         Tag             =   "ALKALI"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton OptBersih 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Bersih"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12000
         TabIndex        =   13
         Top             =   1965
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptKotor 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Kotor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12960
         TabIndex        =   12
         Top             =   1965
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Tag             =   "EA"
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   65273859
         CurrentDate     =   39634
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   6420
         TabIndex        =   6
         Tag             =   "EA"
         Top             =   840
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy    hh:mm"
         Format          =   65273859
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   6420
         TabIndex        =   7
         Tag             =   "EA"
         Top             =   1200
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy    hh:mm"
         Format          =   65273859
         CurrentDate     =   39419
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   6750
         Width           =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   2420
         X2              =   120
         Y1              =   7070
         Y2              =   7070
      End
      Begin VB.Label lblRekomendasi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rekomendasi"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   900
         Width           =   1065
      End
      Begin VB.Label lblRekom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "no_rekom"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1320
         TabIndex        =   33
         Tag             =   "EA"
         Top             =   840
         Width           =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   1360
         X2              =   120
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Selesai Ekstraksi"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   32
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Mulai Ekstraksi"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   31
         Top             =   870
         Width           =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   6480
         X2              =   4440
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   6480
         X2              =   4440
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4440
         TabIndex        =   30
         Top             =   525
         Width           =   1545
      End
      Begin VB.Label lblMO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "refid"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6420
         TabIndex        =   10
         Tag             =   "EA"
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label lblReaktor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reaktor"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   4440
         TabIndex        =   28
         Top             =   180
         Width           =   630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   6465
         X2              =   4440
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1360
         X2              =   120
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   14640
         TabIndex        =   27
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1360
         X2              =   120
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label lblTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   1260
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1360
         X2              =   120
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label lblTanki 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanki"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   9600
         TabIndex        =   24
         Top             =   195
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   6510
         X2              =   4440
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblNoEkstraksi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   10840
         X2              =   9600
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lblNoRL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataSource      =   "DataTrans"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   11280
         TabIndex        =   22
         Tag             =   "ALKALI"
         Top             =   840
         Width           =   2190
      End
      Begin VB.Label lblNoStock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Stock RL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10080
         TabIndex        =   21
         Top             =   900
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   11320
         X2              =   10080
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblBeratRL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat RL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   9960
         TabIndex        =   20
         Top             =   1260
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   11200
         X2              =   9960
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lblDokumentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dok Number"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   165
         Width           =   1020
      End
      Begin VB.Label lblDokNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "dokno"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Tag             =   "EA"
         Top             =   120
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   1360
         X2              =   120
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   11420
         X2              =   9360
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label lblTempatAlkali 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Alkali Treatment"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   9360
         TabIndex        =   18
         Top             =   1605
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   6160
         X2              =   4920
         Y1              =   3495
         Y2              =   3495
      End
      Begin VB.Label lblKondisi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kondisi"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   9720
         TabIndex        =   17
         Top             =   1980
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   11780
         X2              =   9720
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label lblTempat 
         BackColor       =   &H0080FFFF&
         DataField       =   "CompanyName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   11640
         TabIndex        =   16
         Tag             =   "RN"
         Top             =   1605
         Visible         =   0   'False
         Width           =   1605
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1005
      BindFormTAG     =   "EA"
      ActiveLanguage  =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1240
      X2              =   0
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Label lblKeterangan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "FrmProdEkAutoClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Xval As String
Dim MEdit As Boolean
Dim RcDetail As DBQuick
Dim GridAltColor As String
Dim Changingsel As Byte

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RcProduksi As New DBQuick

Private Sub cmdEkstraksi_Click()
    OpenPartner 0
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
    Set mCall = New frmCaller
    
    Select Case Index

        Case 0
            RcProduksi.DBOpen "SELECT LabRekomEkstraksi.RLNO,LabRekomEkstraksi.SplNo From LabRekomEkstraksi", CNN, lckLockReadOnly

        Case 1
            RcProduksi.DBOpen "SELECT [Manufacture Order].OrderID From [Manufacture Order] Order By  [Manufacture Order].OrderID ", CNN, lckLockReadOnly
    End Select
    
    If RcProduksi.Recordcount <> 0 Then

        Select Case Index

            Case 0
                mCall.FromTagActive = "ACID"

            Case 1
                mCall.FromTagActive = "REFERENCE"
        End Select

        Set mCall.FormData = RcProduksi.DBRecordset
        mCall.LookUp Me
    Else

        MessageBox "Konfigurasi EKSTRAKSI DI AUTOCLAVE masih kosong", "Peringatan", msgOkOnly, msgInfo
        OpenPartner = True
    End If

End Function

Private Sub cmdRefLink_Click()
    OpenPartner 1
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm

        Case "ACID"
            lblEkstraksi = mCall.GetFieldByName("RlNo")

        Case "REFERENCE"

            lblMO.Caption = mCall.GetFieldByName("OrderID")
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    ScanKey KeyCode, Shift, MyDDE

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    HiasFormManTell Picture2, Me

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "EA"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From LabProsesProduksi_Header where type_proses='EA'"
        .SetPermissions = aksess.MayDo("Extraction AutoClave")
    End With

End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
    'RcDetail.DBOpen "SELECT LabProsesProduksi_Header.DokNo, labprosesproduksi_line.ProsesID, labprosesproduksi_line.ID,labprosesproduksi_line.Reaktor,labprosesproduksi_line.Bak_Luar1, labprosesproduksi_line.Bak_Luar2, labprosesproduksi_line.Kebutuhan, LabProses.Prosedur, LabAnalysis.Analysis From labprosesproduksi_line INNER JOIN LabProsesProduksi_Header ON (labprosesproduksi_line.DokNo = LabProsesProduksi_Header.DokNo) INNER JOIN LabProses ON (labprosesproduksi_line.ProsesID = LabProses.ProsesID) INNER JOIN LabAnalysis ON (labprosesproduksi_line.ID = LabAnalysis.ID) Where LabProsesProduksi_Header.DokNo = '" & ParameterString & "'", CNN
    Set RcDetail = New DBQuick
    Dim I, ncount As Integer

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
 
    RcDetail.DBOpen "SELECT LabProsesProduksi_Header.DokNo,LabProsesProduksi_Header.rlno,LabProsesProduksi_Header.[refid],LabProsesProduksi_Header.[keterangan],LabProsesProduksi_Header.[tanki],LabProsesProduksi_Header.[group],LabProsesProduksi_Header.[berat], LabProsesProduksi_Header.rlno, labprosesproduksi_line.ProsesID, labprosesproduksi_line.ID,labprosesproduksi_line.Reaktor,labprosesproduksi_line.Bak_Luar1, labprosesproduksi_line.Bak_Luar2, labprosesproduksi_line.Kebutuhan, LabProses.Prosedur, LabAnalysis.Analysis From labprosesproduksi_line INNER JOIN LabProsesProduksi_Header ON (labprosesproduksi_line.DokNo = LabProsesProduksi_Header.DokNo) INNER JOIN LabProses ON (labprosesproduksi_line.ProsesID = LabProses.ProsesID) INNER JOIN LabAnalysis ON (labprosesproduksi_line.ID = LabAnalysis.ID) Where LabProsesProduksi_Header.DokNo = '" & ParameterString & "' and labprosesproduksi_line.Kebutuhan<>''", CNN
    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set tgAutoClave.DataSource = MyDDE.ChildRecordset
    RcDetail.CloseDB

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)

    If (MyDDE.ActiveRecordset.BOF = False) And (MyDDE.ActiveRecordset.EOF = False) Then OpenDetail MyDDE.ActiveRecordset.Fields("DokNo")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case adAddNew

            MEdit = True
            Me.Tag = "baru"
            DcTanggal.Value = Format(Now, "dd mmmm yyyy")
            tgl(0).Value = Format(Now, "dd mmmm yyyy hh:mm")
            tgl(1).Value = Format(Now, "dd mmmm yyyy hh:mm")
        Case tmbSave
            isiText

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                PrepareQuery
            Else
                MyDDE.IsChildMemberReady = False
            End If

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Public Function CheckControls(frmForm As Form) As Boolean

    Dim ctlControl As Variant
    On Error Resume Next

    For Each ctlControl In frmForm

        If Trim(ctlControl.Text) = "" Then
            ctlControl.Text = "-"
            CheckControls = True
            DoEvents
            End If
        Next
        
    End Function
    
Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            CheckControls Me
            SimpanDetail
            SaveToMO

        Case tmbAddNew
            MEdit = True
            Me.Tag = "baru"
            txtBerat.Enabled = True
            DcTanggal.SetFocus
            
            lblMO = frmProduksi.txtBox(5)
            lblNoStock.Caption = frmProduksi.lblSplNo.Caption
            
            lblRekom.Caption = frmProduksi.txtBox(0)
            lblDokNo.Caption = IDGen.GetID("EA")
            lblNoRL.Caption = frmProduksi.lblNoRL
            
            BindDataToGrid lblRekom.Caption
            
            tgAutoClave.AllowUpdate = True
            tgAutoClave.Columns(0).Locked = True
            tgAutoClave.Columns(1).Locked = True
            tgAutoClave.Columns(2).Locked = False

        Case tmbCancel
 
            tgAutoClave.Columns(2).Locked = True

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub BindDataToGrid(ByVal ParameterString As String)
    Dim ncount As Integer
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
  
    RcDetail.DBOpen _
       "SELECT LabRekomEkstraksi.tempatalkali, labprosesproduksi_line.kebutuhan, LabRekomEkstraksi.SplNo , LabRekomEkstraksi.RLNO, LabRekomEkstraksi_Line.FORMID, LabRekomEkstraksi_Line.FormName, LabProses.ProsesID, LabProses.Prosedur, LabAnalysis.ID, " & _
       " LabAnalysis.Analysis, LabSetupRekom_Line.minvalue, LabSetupRekom_Line.maxvalue From  LabRekomEkstraksi_Line INNER JOIN LabRekomEkstraksi ON (LabRekomEkstraksi_Line.SplNo = LabRekomEkstraksi.SplNo)  INNER JOIN LabSetupRekom_Header ON (LabRekomEkstraksi_Line.FORMID = LabSetupRekom_Header.FormID) INNER JOIN LabSetupRekom_Line ON (LabSetupRekom_Header.DocID = LabSetupRekom_Line.DocID)  AND (LabSetupRekom_Header.FormID = LabSetupRekom_Line.FormID) " & _
       "INNER JOIN LabAnalysis ON (LabSetupRekom_Line.ID_ANALYSIS = LabAnalysis.ID) INNER JOIN LabProses ON (LabSetupRekom_Line.ProsesID = LabProses.ProsesID)  left outer join  labprosesproduksi_line ON labrekomekstraksi.splno = labprosesproduksi_line.dokno " & _
       " Where  LabRekomEkstraksi.SplNo = '" & ParameterString & "'  and LabRekomEkstraksi_Line.FORMNAME = 'EXTRACTION AUTOCLAVE' and LabProses.kolom = '0' Order By  LabSetupRekom_Line.ProsesID ", CNN, lckLockBatch
    Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)

    If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveFirst
    Set tgAutoClave.DataSource = MyDDE.ChildRecordset
    RcDetail.CloseDB
   
End Sub

Private Sub SaveToMO()
    On Error GoTo Masjid
    Dim dStart As Date
    Dim dFinish As Date
    Dim ActualTime As Double
    Dim rsCek As New DBQuick
    Dim sWCID As String
    
    dStart = tgl(0).Value
    dFinish = tgl(1).Value
    ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 36", CNN

    If rsCek.DBRecordset.Recordcount > 0 Then
        sWCID = rsCek.DBRecordset.Fields(0)
        SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & lblMO.Caption & "' and WCID='" & sWCID & "'"
    End If

    rsCek.CloseDB
    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

Private Sub isiText()
    txtTanki.Text = "1"
    txtBerat.Text = "1"
End Sub

Private Sub PrepareQuery()
    On Error GoTo Masjid
    Dim ket As Byte

    If OptBersih.Value = True Then
        ket = 1
    Else
        ket = 0
    End If

    With MyDDE

        .PrepareAppend = "INSERT INTO LabProsesProduksi_Header(DokNo, NoEkstraksi, RLNo, Berat, Tanggal,[Group], Tanki, Keterangan, Kondisi,refid,type_proses,no_rekom) VALUES ('" & lblDokNo.Caption & "','" & lblEkstraksi & "','" & lblNoRL & "','" & txtBerat.Text & "','" & DcTanggal & "','" & txtGroup.Text & "','" & txtTanki.Text & "','" & txtKeterangan.Text & "','" & ket & "','" & lblMO.Caption & "','" & "EA" & "','" & lblRekom.Caption & "')"
        .PrepareUpdate = "UPDATE LabProsesProduksi_Header SET  DokNo = '" & lblDokNo.Caption & "', no_rekom='" & lblRekom.Caption & "',refid='" & lblMO.Caption & "',NoEkstraksi ='" & lblEkstraksi & "',RLNo ='" & lblNoRL & "',Berat ='" & txtBerat.Text & "',Tanggal ='" & DcTanggal & "',[Group] ='" & txtGroup.Text & "', Tanki ='" & txtTanki.Text & "', Keterangan ='" & txtKeterangan.Text & "', Kondisi = '" & txtKeterangan.Text & "'"
        .PrepareDelete = "DELETE From LabProsesProduksi_Header Where LabProsesProduksi_Header.DokNo = '" & lblDokNo.Caption & "'"
    End With

    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

Private Sub SimpanDetail()
    On Error GoTo Masjid
    Dim I As Integer

    tgAutoClave.row = 0
    MyDDE.ChildRecordset.MoveFirst

    For I = 0 To MyDDE.ChildRecordset.Recordcount - 1
        SendDataToServer "INSERT INTO labprosesproduksi_line(DokNo,ProsesID,ID,kebutuhan) VALUES('" & lblDokNo.Caption & "','" & tgAutoClave.Columns(3).Value & "','" & tgAutoClave.Columns(4).Value & "','" & IIf(tgAutoClave.Columns("kebutuhan").Value = "", ".", tgAutoClave.Columns("kebutuhan").Value) & "')"
        MyDDE.ChildRecordset.MoveNext
    Next
    SendDataToServer "UPDATE StatusProduksi SET Rekomendasi = '" & lblRekom.Caption & "',Posisi = '" & "EXTRACTION AUTOCLAVE" & "', status = '1',tanggal = convert(datetime,'" & Format(Now, "dd/mm/yy") & "',3) Where  StatusProduksi.NoEkstraksi = '" & lblEkstraksi & "'"
    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

