VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmProdAlkali 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alkali Treatment"
   ClientHeight    =   8160
   ClientLeft      =   -1005
   ClientTop       =   -150
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProdAlkali2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9285
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      DataSource      =   "MyDDE"
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
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   9285
      TabIndex        =   19
      Top             =   0
      Width           =   9285
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   17
         Tag             =   "ACID"
         Top             =   6900
         Width           =   6645
      End
      Begin TrueOleDBGrid80.TDBGrid tgAlkali 
         Height          =   4095
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
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
         Columns(2).Caption=   "Nilai Minimum"
         Columns(2).DataField=   "minvalue"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nilai Maksimum"
         Columns(3).DataField=   "maxvalue"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Kebutuhan"
         Columns(4).DataField=   "kebutuhan"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ProsesID"
         Columns(5).DataField=   "ProsesID"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ID"
         Columns(6).DataField=   "ID"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).SizeMode=   2
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5741"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5662"
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
         Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=3466"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3387"
         Splits(0)._ColumnProps(25)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(31)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(1)._UserFlags=   0
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   688
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1)._GSX_SAVERECORDSELECTORS=   65562
         Splits(1).DividerColor=   14215660
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=7"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=6165"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=6085"
         Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(7)=   "Column(1).Width=4075"
         Splits(1)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(9)=   "Column(1)._WidthInPix=3995"
         Splits(1)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(11)=   "Column(2).Width=1852"
         Splits(1)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(13)=   "Column(2)._WidthInPix=1773"
         Splits(1)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(15)=   "Column(3).Width=2011"
         Splits(1)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(17)=   "Column(3)._WidthInPix=1931"
         Splits(1)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(19)=   "Column(4).Width=1588"
         Splits(1)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(21)=   "Column(4)._WidthInPix=1508"
         Splits(1)._ColumnProps(22)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(23)=   "Column(5).Width=1296"
         Splits(1)._ColumnProps(24)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(25)=   "Column(5)._WidthInPix=1217"
         Splits(1)._ColumnProps(26)=   "Column(5).Visible=0"
         Splits(1)._ColumnProps(27)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(28)=   "Column(6).Width=1746"
         Splits(1)._ColumnProps(29)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(30)=   "Column(6)._WidthInPix=1667"
         Splits(1)._ColumnProps(31)=   "Column(6).Visible=0"
         Splits(1)._ColumnProps(32)=   "Column(6).Order=7"
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=66,.parent=67"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=68"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=69"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=71"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=58,.parent=67"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=68"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=69"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=71"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=90,.parent=67"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=87,.parent=68"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=88,.parent=69"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=89,.parent=71"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=94,.parent=67"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=91,.parent=68"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=92,.parent=69"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=93,.parent=71"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=98,.parent=67"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=95,.parent=68"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=96,.parent=69"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=97,.parent=71"
         _StyleDefs(60)  =   "Splits(1).Style:id=13,.parent=1"
         _StyleDefs(61)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(62)  =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(63)  =   "Splits(1).FooterStyle:id=15,.parent=3"
         _StyleDefs(64)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(65)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(66)  =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(67)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(68)  =   "Splits(1).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(69)  =   "Splits(1).OddRowStyle:id=21,.parent=10"
         _StyleDefs(70)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(71)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(72)  =   "Splits(1).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(73)  =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(74)  =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(75)  =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(76)  =   "Splits(1).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(77)  =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(78)  =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(79)  =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(80)  =   "Splits(1).Columns(2).Style:id=102,.parent=13"
         _StyleDefs(81)  =   "Splits(1).Columns(2).HeadingStyle:id=99,.parent=14"
         _StyleDefs(82)  =   "Splits(1).Columns(2).FooterStyle:id=100,.parent=15"
         _StyleDefs(83)  =   "Splits(1).Columns(2).EditorStyle:id=101,.parent=17"
         _StyleDefs(84)  =   "Splits(1).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(85)  =   "Splits(1).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(86)  =   "Splits(1).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(87)  =   "Splits(1).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(88)  =   "Splits(1).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(89)  =   "Splits(1).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(90)  =   "Splits(1).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(91)  =   "Splits(1).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(92)  =   "Splits(1).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(93)  =   "Splits(1).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(94)  =   "Splits(1).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(95)  =   "Splits(1).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(96)  =   "Splits(1).Columns(6).Style:id=54,.parent=13"
         _StyleDefs(97)  =   "Splits(1).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(98)  =   "Splits(1).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(99)  =   "Splits(1).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(100) =   "Named:id=33:Normal"
         _StyleDefs(101) =   ":id=33,.parent=0"
         _StyleDefs(102) =   "Named:id=34:Heading"
         _StyleDefs(103) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(104) =   ":id=34,.wraptext=-1"
         _StyleDefs(105) =   "Named:id=35:Footing"
         _StyleDefs(106) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(107) =   "Named:id=36:Selected"
         _StyleDefs(108) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(109) =   "Named:id=37:Caption"
         _StyleDefs(110) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(111) =   "Named:id=38:HighlightRow"
         _StyleDefs(112) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(113) =   "Named:id=39:EvenRow"
         _StyleDefs(114) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(115) =   "Named:id=40:OddRow"
         _StyleDefs(116) =   ":id=40,.parent=33"
         _StyleDefs(117) =   "Named:id=41:RecordSelector"
         _StyleDefs(118) =   ":id=41,.parent=34"
         _StyleDefs(119) =   "Named:id=42:FilterBar"
         _StyleDefs(120) =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cmdRefLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         Picture         =   "FrmProdAlkali2.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "BAHAN"
         Top             =   840
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtReaktor 
         Appearance      =   0  'Flat
         DataField       =   "reaktor"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Tag             =   "ALKALI"
         Top             =   1920
         Width           =   2340
      End
      Begin VB.OptionButton optKotor 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Kotor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   14
         Top             =   1590
         Width           =   855
      End
      Begin VB.OptionButton optBersih 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Bersih"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6690
         TabIndex        =   13
         Top             =   1590
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtBerat 
         Appearance      =   0  'Flat
         DataField       =   "berat"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6690
         TabIndex        =   8
         Tag             =   "ALKALI"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "group"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6690
         TabIndex        =   9
         Tag             =   "ALKALI"
         Top             =   480
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Tag             =   "ALKALI"
         Top             =   1560
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
         Format          =   65536003
         CurrentDate     =   39634
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   6690
         TabIndex        =   15
         Tag             =   "ALKALI"
         Top             =   1920
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
         Format          =   65536003
         CurrentDate     =   39584
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   6690
         TabIndex        =   16
         Tag             =   "ALKALI"
         Top             =   2280
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
         Format          =   65536003
         CurrentDate     =   39584
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   7290
         Width           =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   2420
         X2              =   120
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   17
         X1              =   1360
         X2              =   120
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label lblMetode1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Metode"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   1252
         Width           =   735
      End
      Begin VB.Label lblMetode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "methode"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Tag             =   "ALKALI"
         Top             =   1200
         Width           =   2190
      End
      Begin VB.Label labell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rekomendasi"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   892
         Width           =   1065
      End
      Begin VB.Label lblMO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "refid"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6690
         TabIndex        =   11
         Tag             =   "ALKALI"
         Top             =   840
         Width           =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   6705
         X2              =   4440
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblNoRL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "rlno"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Tag             =   "ALKALI"
         Top             =   2280
         Width           =   2190
      End
      Begin VB.Label lblDokNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "dokno"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Tag             =   "ALKALI"
         Top             =   120
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   6810
         X2              =   4440
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   6810
         X2              =   4440
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Mulai Alkali"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   33
         Top             =   1950
         Width           =   2250
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Selesai Alkali"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   32
         Top             =   2310
         Width           =   2055
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4440
         TabIndex        =   31
         Top             =   885
         Width           =   1545
      End
      Begin VB.Label lblRekomendasi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "no_rekom"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Tag             =   "ALKALI"
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label lblBeratRL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reaktor"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   1965
         Width           =   630
      End
      Begin VB.Label lblTempat 
         BackColor       =   &H0080FFFF&
         DataField       =   "tempatalkali"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6690
         TabIndex        =   12
         Tag             =   "ALKALI"
         Top             =   1230
         Width           =   1965
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   6810
         X2              =   4440
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label lblKondisi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kondisi"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   4440
         TabIndex        =   29
         Top             =   1605
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   6810
         X2              =   4440
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label lblTempatAlkali 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Alkali Treatment"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   4440
         TabIndex        =   28
         Top             =   1245
         Width           =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   6700
         X2              =   4440
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   1360
         X2              =   120
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblDokumentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dok Number"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   165
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   1360
         X2              =   120
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lblBeratRL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat RL"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4440
         TabIndex        =   26
         Top             =   172
         Width           =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   1360
         X2              =   120
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblNoStock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Stock RL"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label lblEkstraksi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "noekstraksi"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Tag             =   "ALKALI"
         Top             =   480
         Width           =   2190
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
         Index           =   5
         X1              =   6705
         X2              =   4440
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4440
         TabIndex        =   22
         Top             =   532
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1360
         X2              =   120
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label lblTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1620
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1360
         X2              =   120
         Y1              =   765
         Y2              =   765
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
         TabIndex        =   20
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1360
         X2              =   120
         Y1              =   1860
         Y2              =   1860
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7590
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   1005
      BindFormTAG     =   "ALKALI"
      ActiveLanguage  =   1
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
      TabIndex        =   24
      Top             =   0
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1240
      X2              =   0
      Y1              =   255
      Y2              =   255
   End
End
Attribute VB_Name = "FrmProdAlkali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MEdit As Boolean
Dim RcDetail As DBQuick
Dim Changingsel As Byte

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RcProduksi As New DBQuick

Private Sub cmdEkstraksi_Click()
    OpenPartner 0
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

        MessageBox "Konfigurasi ALKALI TREATMENT masih kosong", "Peringatan", msgOkOnly, msgInfo
        OpenPartner = True
    End If

End Function

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm

        Case "ACID"
            lblEkstraksi.Caption = mCall.GetFieldByName("RlNo")

        Case "REFERENCE"

            lblRekomendasi.Caption = mCall.GetFieldByName("OrderID")
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    ScanKey KeyCode, Shift, MYDDE

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    HiasFormManTell Picture2, Me

    With MYDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "ALKALI"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From LabProsesProduksi_Header where type_proses='ALKALI'"
        .SetPermissions = aksess.MayDo("Alkali Treatment")
    End With

End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
    Set RcDetail = New DBQuick
    Dim I, ncount As Integer

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
 
    RcDetail.DBOpen "SELECT LabProsesProduksi_Header.DokNo,LabProsesProduksi_Header.[refid],LabProsesProduksi_Header.[keterangan],LabProsesProduksi_Header.[tanki],LabProsesProduksi_Header.[group],LabProsesProduksi_Header.[berat], LabProsesProduksi_Header.rlno, labprosesproduksi_line.ProsesID, labprosesproduksi_line.ID,labprosesproduksi_line.Reaktor,labprosesproduksi_line.Bak_Luar1, labprosesproduksi_line.Bak_Luar2, labprosesproduksi_line.Kebutuhan, LabProses.Prosedur, LabAnalysis.Analysis From labprosesproduksi_line INNER JOIN LabProsesProduksi_Header ON (labprosesproduksi_line.DokNo = LabProsesProduksi_Header.DokNo) INNER JOIN LabProses ON (labprosesproduksi_line.ProsesID = LabProses.ProsesID) INNER JOIN LabAnalysis ON (labprosesproduksi_line.ID = LabAnalysis.ID) Where LabProsesProduksi_Header.DokNo = '" & ParameterString & "' and labprosesproduksi_line.Kebutuhan<>''", CNN
    Set MYDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
    Set tgAlkali.DataSource = MYDDE.ChildRecordset
    RcDetail.CloseDB
End Sub

Private Sub BindDataToGrid(ByVal ParameterString As String)
    Dim ncount As Integer
    Set RcDetail = New DBQuick

    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
    RcDetail.DBOpen "SELECT  Case labrekomekstraksi.tempatalkali WHEN 0 THEN 'Reaktor' WHEN 1 THEN 'Bak Luar' WHEN 2 THEN 'AutoClave' END As [Tempat Ekstraksi],labrekomekstraksi.splno,labrekomekstraksi.rlno,labrekomekstraksi_line.formid,labrekomekstraksi_line.formname,labproses.ProsesID,labproses.Prosedur," & _
       "labanalysis.ID,labanalysis.Analysis,labsetuprekom_line.minvalue,labsetuprekom_line.maxvalue,labprosesproduksi_line.kebutuhan From labrekomekstraksi_line INNER JOIN labrekomekstraksi ON labrekomekstraksi_line.splno =labrekomekstraksi.splno " & _
       " INNER JOIN labsetuprekom_header ON labrekomekstraksi_line.formid =labsetuprekom_header.FormID INNER JOIN labsetuprekom_line ON labsetuprekom_header.DocID =labsetuprekom_line.DocID AND labsetuprekom_header.FormID =labsetuprekom_line.FormID " & _
       " INNER JOIN labanalysis ON labsetuprekom_line.ID_ANALYSIS = labanalysis.ID INNER JOIN labproses ON labsetuprekom_line.ProsesID = labproses.ProsesID LEFT OUTER JOIN labprosesproduksi_line ON labprosesproduksi_line.dokno =labrekomekstraksi.splno " & _
       " WHERE (labrekomekstraksi.splno = '" & ParameterString & "') AND (labrekomekstraksi_line.formname = 'ALKALI TREATMENT') AND (labproses.kolom = '0') ORDER BY labsetuprekom_line.ProsesID", CNN, lckLockBatch
       
       
     ' Debug.Print RcDetail.DBRecordset.Source
    Set MYDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)

    If Not MYDDE.ChildRecordset.EOF Then MYDDE.ChildRecordset.MoveFirst
    Set tgAlkali.DataSource = MYDDE.ChildRecordset
    lblTempat.Caption = RcDetail.DBRecordset.Fields("Tempat Ekstraksi")
    tgAlkali.Columns(4).Caption = RcDetail.DBRecordset.Fields("Tempat Ekstraksi")
    RcDetail.CloseDB
   
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbAddNew

            MEdit = True
            Me.Tag = "baru"
            txtBerat.Enabled = True
            DcTanggal.SetFocus
            
            lblMO = frmProduksi.txtBox(5)
            lblNoRL.Caption = frmProduksi.lblSplNo.Caption
            lblEkstraksi = frmProduksi.txtBox(1)
            lblRekomendasi.Caption = frmProduksi.txtBox(0)
            lblDokNo.Caption = IDGen.GetID("ALKALI")
            lblNoRL.Caption = frmProduksi.lblNoRL
            lblMetode.Caption = frmProduksi.lblMetode
            BindDataToGrid lblRekomendasi
            
            tgAlkali.AllowUpdate = True
            tgAlkali.Columns(0).Locked = True
            tgAlkali.Columns(1).Locked = True
            tgAlkali.Columns(2).Locked = False
            txtReaktor.Text = ""

        Case tmbCancel

            tgAlkali.Columns(2).Locked = True

        Case tmbSave
            SimpanDetail
            SaveToMO

        Case tmbDelete
            PrepareQuery
    End Select

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
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 37", CNN

    If rsCek.DBRecordset.Recordcount > 0 Then
        sWCID = rsCek.DBRecordset.Fields(0)
        SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & lblRekomendasi.Caption & "' and WCID='" & sWCID & "'"
    End If

    rsCek.CloseDB
    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)
                               
    If (MYDDE.ActiveRecordset.BOF = False) And (MYDDE.ActiveRecordset.EOF = False) Then OpenDetail MYDDE.ActiveRecordset.Fields("DokNo")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            CheckControls Me
            If MYDDE.CheckEmptyControl = False Then
                MYDDE.IsChildMemberReady = True
                PrepareQuery
                
            Else
                MYDDE.IsChildMemberReady = False
            End If
            
        Case tmbAddNew
            txtBerat.Enabled = True
            txtBerat.SetFocus
            DcTanggal.Value = Format(Now, "dd mmmm yyyy")
            tgl(0).Value = Format(Now, "dd mmmm yyyy hh:mm")
            tgl(1).Value = Format(Now, "dd mmmm yyyy hh:mm")
        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub PrepareQuery()
    On Error GoTo Masjid
    Dim ket As Byte

    If OptBersih.Value = True Then
        ket = 1
    Else
        ket = 0
    End If

    With MYDDE
        .PrepareAppend = "INSERT INTO LabProsesProduksi_Header(DokNo, NoEkstraksi, RLNo, Berat, Tanggal,[Group],  Keterangan, Kondisi, refid,type_proses, tempatalkali,no_rekom,waktu_mulai,waktu_selesai,methode) VALUES ('" & lblDokNo.Caption & "','" & lblEkstraksi.Caption & "','" & lblNoRL & "','" & txtBerat.Text & "',convert(datetime,'" & DcTanggal & "',3),'" & txtGroup.Text & "','" & TxtKeterangan.Text & "','" & ket & "','" & lblMO.Caption & "','" & "ALKALI" & "','" & lblTempat.Caption & "','" & lblRekomendasi.Caption & "',convert(datetime,'" & tgl(0).Value & "',3), convert(datetime,'" & tgl(1).Value & "',3),'" & lblMetode.Caption & "')"
        .PrepareUpdate = "UPDATE LabProsesProduksi_Header SET  DokNo = '" & lblDokNo.Caption & "',waktu_mulai=convert(datetime,'" & tgl(0).Value & "',3),waktu_selesai= convert(datetime,'" & tgl(1).Value & "',3), no_rekom='" & lblRekomendasi.Caption & "',NoEkstraksi ='" & lblEkstraksi.Caption & "',refid='" & lblMO.Caption & "',RLNo ='" & lblNoRL & "',Berat ='" & txtBerat.Text & "',Tanggal = convert(datetime,'" & DcTanggal & "',3),[Group] ='" & txtGroup.Text & "', Keterangan ='" & TxtKeterangan.Text & "', Kondisi = '" & TxtKeterangan.Text & "'"
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

    tgAlkali.row = 0
    MYDDE.ChildRecordset.MoveFirst

    For I = 0 To MYDDE.ChildRecordset.Recordcount - 1
        SendDataToServer "INSERT INTO labprosesproduksi_line(DokNo,ProsesID,ID,kebutuhan) VALUES('" & lblDokNo.Caption & "','" & tgAlkali.Columns("ProsesID").Value & "','" & tgAlkali.Columns("ID").Value & "','" & IIf(tgAlkali.Columns("kebutuhan").Value = "", ".", tgAlkali.Columns("kebutuhan").Value) & "')"
        MYDDE.ChildRecordset.MoveNext
    Next
    SendDataToServer "UPDATE StatusProduksi SET Rekomendasi = '" & lblRekomendasi.Caption & "',Posisi = '" & "ALKALI TREATMENT" & "', status = '1',tanggal = convert(datetime,'" & Format(Now, "dd/mm/yyyy") & "',3) Where  StatusProduksi.NoEkstraksi = '" & lblEkstraksi.Caption & "'"
    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

Private Function isNum(Keya As Integer)

    If Not (Keya >= Asc("0") And Keya <= Asc("9") Or Keya = vbKeyBack Or Keya = Asc(".")) Then
        Beep
        Keya = 0
    End If
  
End Function

Private Sub txtBerat_KeyPress(KeyAscii As Integer)
    isNum KeyAscii
End Sub
