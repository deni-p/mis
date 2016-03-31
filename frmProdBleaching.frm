VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmProdBleaching 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bleaching Treatment"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdBleaching.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11505
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
      Height          =   6705
      Left            =   0
      ScaleHeight     =   6705
      ScaleWidth      =   11505
      TabIndex        =   1
      Top             =   0
      Width           =   11505
      Begin VB.TextBox lblEkstraksi 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Tag             =   "BLEACHING"
         Top             =   180
         Width           =   2160
      End
      Begin VB.TextBox txtTanki 
         Appearance      =   0  'Flat
         DataField       =   "tanki"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Tag             =   "BLEACHING"
         Top             =   1230
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Tag             =   "BLEACHING"
         Top             =   525
         Width           =   2190
         _ExtentX        =   3863
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
         Format          =   64880643
         CurrentDate     =   39634
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         Height          =   675
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "BLEACHING"
         Top             =   5985
         Width           =   11280
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Tag             =   "BLEACHING"
         Top             =   885
         Width           =   1230
      End
      Begin VB.TextBox txtPh 
         Appearance      =   0  'Flat
         DataField       =   "ph_akhir"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   9990
         TabIndex        =   2
         Tag             =   "BLEACHING"
         Top             =   5340
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   2235
         TabIndex        =   6
         Tag             =   "BLEACHING"
         Top             =   1830
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
         CustomFormat    =   "dd MMM yyyy    HH:mm"
         Format          =   64880643
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   2265
         TabIndex        =   7
         Tag             =   "BLEACHING"
         Top             =   5340
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
         CustomFormat    =   "dd MMM yyyy    HH:mm"
         Format          =   64880643
         CurrentDate     =   39419
      End
      Begin TrueOleDBGrid80.TDBGrid gridDetail 
         Height          =   3075
         Left            =   120
         TabIndex        =   8
         Top             =   2220
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   5424
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tahapan"
         Columns(0).DataField=   "tahap"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Proses"
         Columns(1).DataField=   "proses"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Hasil"
         Columns(2).DataField=   "nilai"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3122"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3043"
         Splits(0)._ColumnProps(4)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Merge=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3651"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3572"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2064"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1984"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.wraptext=-1"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HE0E0E0&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSComCtl2.DTPicker dateGrid 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   7815
         TabIndex        =   9
         Tag             =   "ALKALI"
         Top             =   2685
         Visible         =   0   'False
         Width           =   1740
         _ExtentX        =   3069
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
         CustomFormat    =   "HH:mm"
         Format          =   64880643
         CurrentDate     =   39584
      End
      Begin TrueOleDBGrid80.TDBGrid gridPencucian 
         Height          =   3075
         Left            =   5700
         TabIndex        =   10
         Top             =   2220
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   5424
         _LayoutType     =   4
         _RowHeight      =   18
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pencucian"
         Columns(0).DataField=   "pencucian"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Jml Air"
         Columns(1).DataField=   "jml_air"
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Waktu Mulai"
         Columns(2).DataField=   "waktu_mulai"
         Columns(2).NumberFormat=   "HH:mm"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Waktu Selesai"
         Columns(3).DataField=   "waktu_selesai"
         Columns(3).NumberFormat=   "HH:mm"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1111"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1032"
         Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=3043"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2963"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=3043"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2963"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HEAEAEA&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9285
         TabIndex        =   25
         Top             =   915
         Width           =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   9945
         X2              =   7725
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         Height          =   195
         Left            =   7725
         TabIndex        =   26
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblRekomendasi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "rekomendasi"
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
         Left            =   9285
         TabIndex        =   12
         Tag             =   "BLEACHING"
         Top             =   555
         Width           =   2070
      End
      Begin VB.Label lblMO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "manufacture_order"
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
         Left            =   9285
         TabIndex        =   13
         Tag             =   "BLEACHING"
         Top             =   195
         Width           =   2070
      End
      Begin VB.Label lblKeterangan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   23
         Top             =   5730
         Width           =   840
      End
      Begin VB.Label lblNoEkstraksi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rekomendasi"
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
         Height          =   195
         Left            =   7725
         TabIndex        =   22
         Top             =   600
         Width           =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   9900
         X2              =   7695
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   930
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1360
         X2              =   120
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Label lblTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   585
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1360
         X2              =   120
         Y1              =   480
         Y2              =   480
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
         TabIndex        =   19
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1345
         X2              =   105
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Order"
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
         Height          =   195
         Left            =   7710
         TabIndex        =   18
         Top             =   240
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   2245
         X2              =   165
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   2215
         X2              =   135
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu mulai"
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
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   1860
         Width           =   1890
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && waktu selesai"
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
         Index           =   4
         Left            =   165
         TabIndex        =   16
         Top             =   5370
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   9945
         X2              =   7725
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lbleksno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ekstraksi"
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
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblMetod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reaktor"
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
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   1260
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   17
         X1              =   2355
         X2              =   120
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   120
         X2              =   11280
         Y1              =   1665
         Y2              =   1650
      End
      Begin VB.Label lblBeratRL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pH Akhir Bleachig Teratment"
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
         Height          =   195
         Index           =   2
         Left            =   7620
         TabIndex        =   11
         Top             =   5355
         Width           =   2025
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   11085
         X2              =   7605
         Y1              =   5625
         Y2              =   5625
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1005
      BindFormTAG     =   "BLEACH"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmProdBleaching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MEdit As Boolean
Private RsDetail As New DBQuick
Private RsPencucian As New DBQuick



Private Sub dateGrid_Change()
   If 3 >= gridPencucian.col <= 2 Then
      gridPencucian.Columns(gridPencucian.col) = dateGrid.Value
   End If
End Sub

Private Sub gridPencucian_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   dateGrid.Visible = False
   If (gridPencucian.col = 2) Or (gridPencucian.col = 3) Then
      If Not IsNull(gridPencucian.Columns(gridPencucian.col)) Then
         dateGrid.Value = IIf(gridPencucian.Columns(gridPencucian.col) = "", "00:00", gridPencucian.Columns(gridPencucian.col))
      Else
         dateGrid.Value = Now
      End If
      
      dateGrid.Visible = True
      dateGrid.Move gridPencucian.Left + gridPencucian.Columns(gridPencucian.col).Left, _
                    gridPencucian.Top + gridPencucian.RowTop(gridPencucian.row), _
                    gridPencucian.Columns(gridPencucian.col).width, _
                    gridPencucian.RowHeight
   End If
End Sub



Private Sub Form_Load()
    HiasFormManTell Picture2, Me

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "BLEACHING"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From bleaching"
        .SetPermissions = aksess.MayDo("Bleaching Treatment")
    End With

End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
      RsDetail.DBOpen "select * from bleaching_detail where no_ekstraksi='" & ParameterString & "' ", CNN
      Set gridDetail.DataSource = RsDetail.DBRecordset
      
      RsPencucian.DBOpen "select * from bleaching_pencucian where no_ekstraksi='" & ParameterString & "' ", CNN
      Set gridPencucian.DataSource = RsPencucian.DBRecordset
End Sub


Private Sub lblEkstraksi_LostFocus()
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & lblEkstraksi.Text & "'", CNN, lckLockBatch
   If rsCek.DBRecordset.Recordcount > 0 Then
      rsCek.DBOpen "select * from bleaching where no_Ekstraksi='" & lblEkstraksi.Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
         lblEkstraksi.Text = ""
      End If
   Else
      MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
      lblEkstraksi.Text = ""
   End If
   rsCek.CloseDB

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Dim x As Integer
   
    Select Case AdReasonActiveDb

        Case tmbAddNew

            MEdit = True
            Me.Tag = "baru"
            DcTanggal.SetFocus
            
            lblMO = frmProduksi.txtBox(5)
            'lblEkstraksi = frmProduksi.txtBox(1)
            lblRekomendasi.Caption = frmProduksi.txtBox(0)
            txtKeterangan.Text = "-"
            
            For x = 0 To 1
               tgl(x).Value = Now
            Next
            
            AddDetail
            AddPencucian
            

        Case tmbCancel


        Case tmbSave
            SimpanDetail
            SaveToMO
            SaveToStatus

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub SaveToStatus()
   Dim rsStatus As New DBQuick
   rsStatus.DBOpen "select * from statusProduksi where noEkstraksi='" & lblEkstraksi & "'", CNN, lckLockBatch
   If rsStatus.DBRecordset.Recordcount > 0 Then
      SendDataToServer "update statusproduksi set status=1,posisi='ACID' where noEkstraksi='" & lblEkstraksi & "'"
   Else
      SendDataToServer "INSERT INTO [statusproduksi] " & _
                              "([NoEkstraksi]" & _
                              ",[Rekomendasi]" & _
                              ",[Posisi]" & _
                              ",[status]" & _
                              ",[tanggal]) " & _
                       "Values " & _
                              "('" & lblEkstraksi & "'" & _
                              ",'" & lblMO.Caption & "'" & _
                              ",'ACID',1,'" & Format(Now, "yyyy-MM-dd") & "')"
   End If
End Sub

Private Sub AddDetail()
   With RsDetail.DBRecordset
      '*** 1
      .AddNew
      .Fields("tahap") = "Air bersih untuk penambahan proses Bleaching"
      .Fields("id_tahapan") = 1
      .Fields("proses") = "Jumlah (liter)"
      
      '*** 2
      .AddNew
      .Fields("tahap") = "Bleaching Agent"
      .Fields("id_tahapan") = 2
      .Fields("proses") = "Type Bleaching Agent"
   
      '*** 3
      .AddNew
      .Fields("tahap") = "Bleaching Agent"
      .Fields("id_tahapan") = 2
      .Fields("proses") = "Jumlah (Kg)"
      
      '*** 4
      .AddNew
      .Fields("tahap") = "Bleaching Agent"
      .Fields("id_tahapan") = 2
      .Fields("proses") = "Konsentrasi (%)"
   
      '*** 5
      .AddNew
      .Fields("tahap") = "Bleaching Agent"
      .Fields("id_tahapan") = 2
      .Fields("proses") = "Waktu"
      
      '*** 6
      .AddNew
      .Fields("tahap") = "Larutan Bleaching Akhir"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Jumlah (Liter)"
   
      '*** 7
      .AddNew
      .Fields("tahap") = "Larutan Bleaching Akhir"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Konsentrasi (%)"
      
      '*** 8
      .AddNew
      .Fields("tahap") = "Larutan Bleaching Akhir"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Suhu"
   
      '*** 9
      .AddNew
      .Fields("tahap") = "Waktu Bleaching Treatmnet"
      .Fields("id_tahapan") = 4
      .Fields("proses") = "Suhu pada 20 Menit"
            
      .MoveFirst
   End With
   Set gridDetail.DataSource = RsDetail.DBRecordset
End Sub


Private Sub AddPencucian()
   Dim x As Integer
   With RsPencucian.DBRecordset
      For x = 1 To 4
         '*** 1
         .AddNew
         .Fields("pencucian") = "Pencucian " & x
         .Fields("jml_air") = 0
         .Fields("waktu_mulai") = Now
         .Fields("waktu_selesai") = Now
      Next
      .MoveFirst
   End With
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
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 39", CNN

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
                               
    If MyDDE.ActiveRecordset.Recordcount > 0 Then OpenDetail MyDDE.GetFieldByName("no_ekstraksi")
    Label1.Caption = IIf(IsNull(MyDDE.GetFieldByName("Approved_by")), "", MyDDE.GetFieldByName("Approved_by"))
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                PrepareQuery
                
            Else
                MyDDE.IsChildMemberReady = False
            End If
            
        Case tmbAddNew
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


    With MyDDE
        .PrepareAppend = "INSERT INTO [bleaching] " & _
                           "([No_ekstraksi],[tanggal],[Grup],[tanki]" & _
                           ",[Manufacture_Order],[rekomendasi],[tanggal_mulai],[tanggal_selesai],[ph_akhir]" & _
                           ",[keterangan],issued_by) " & _
                        " Values " & _
                           "('" & lblEkstraksi & "','" & Format(DcTanggal.Value, "yyyy-MM-dd") & "','" & txtGroup & "'" & _
                           ",'" & txtTanki & "','" & lblMO.Caption & "','" & lblRekomendasi & "' " & _
                           ",'" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "','" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                           "," & FQty(txtPh) & ",'" & txtKeterangan & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
                           

        .PrepareUpdate = "UPDATE [bleaching] SET " & _
                             "[tanggal] = '" & Format(DcTanggal.Value, "yyyy-MM-dd") & "'" & _
                             ",[Grup] ='" & txtGroup & "'" & _
                             ",[tanki] = '" & txtTanki & "'" & _
                             ",[MAnufacture_order] = '" & lblMO.Caption & "'" & _
                             ",[rekomendasi] = '" & lblRekomendasi.Caption & "'" & _
                             ",[Tanggal_mulai] = '" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[Tanggal_selesai] = '" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[keterangan] = '" & txtKeterangan & "'" & _
                             ",[ph_akhir] = " & FQty(txtPh) & _
                        " WHERE no_ekstraksi='" & lblEkstraksi & "'"
        
        .PrepareDelete = "DELETE From bleaching Where no_ekstraksi='" & lblEkstraksi & "'"
    End With

    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

Private Sub SimpanDetail()
On Error GoTo xErr
   '*** Update Data on Alkali Detail
   SendDataToServer "delete from bleaching_detail where no_ekstraksi='" & lblEkstraksi & "'"
   With RsDetail.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO [bleaching_detail] ([no_ekstraksi]" & _
                                 ",[tahap]" & _
                                 ",[proses]" & _
                                 ",[nilai]" & _
                                 ",[rekomendasi]" & _
                                 ",id_tahapan)" & _
                           "Values ('" & lblEkstraksi & "'" & _
                                 ",'" & .Fields("tahap") & "'" & _
                                 ",'" & .Fields("proses") & "'" & _
                                 ",'" & .Fields("nilai") & "'" & _
                                 ",'" & .Fields("rekomendasi") & "'" & _
                                 ", " & .Fields("id_tahapan") & ")"
         .MoveNext
      Wend
   End With
   
   '*** Update Data on Alkali Pencucian
   SendDataToServer "delete from bleaching_pencucian where no_ekstraksi='" & lblEkstraksi & "'"
   
   With RsPencucian.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO [bleaching_pencucian] ([no_ekstraksi] " & _
                              ",[pencucian]" & _
                              ",[jml_air]" & _
                              ",[waktu_mulai]" & _
                              ",[waktu_selesai]) " & _
                          " Values ('" & lblEkstraksi & "'" & _
                              ",'" & .Fields("pencucian") & "'" & _
                              ",'" & FQty(.Fields("jml_air")) & "'" & _
                              ",'" & Format(.Fields("waktu_mulai"), "yyyy-MM-dd hh:mm:ss") & "'" & _
                              ",'" & Format(.Fields("waktu_selesai"), "yyyy-MM-dd hh:mm:ss") & "')"
         .MoveNext
      Wend
   End With
Exit Sub
xErr:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

