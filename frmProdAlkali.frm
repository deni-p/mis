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
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdAlkali.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   13215
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
      ScaleWidth      =   13215
      TabIndex        =   14
      Top             =   0
      Width           =   13215
      Begin VB.TextBox lblEkstraksi 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1290
         TabIndex        =   44
         Tag             =   "ALKALI"
         Top             =   225
         Width           =   2220
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Bak Luar"
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
         Left            =   3015
         TabIndex        =   41
         Top             =   2670
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAAF6F&
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
         Height          =   315
         Left            =   2115
         TabIndex        =   40
         Top             =   2670
         Width           =   1155
      End
      Begin VB.TextBox txtTempatAlkali 
         Appearance      =   0  'Flat
         DataField       =   "tempat_alkali"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   4095
         TabIndex        =   39
         Tag             =   "ALKALI"
         Top             =   2670
         Width           =   1335
      End
      Begin VB.TextBox txtPh 
         Appearance      =   0  'Flat
         DataField       =   "ph_akhir"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   11730
         TabIndex        =   36
         Tag             =   "ALKALI"
         Top             =   6360
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid gridDetail 
         Height          =   3720
         Left            =   90
         TabIndex        =   30
         Top             =   3420
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   6562
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tahapan"
         Columns(0).DataField=   "tahapan"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Proses"
         Columns(1).DataField=   "proses"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Reaktor"
         Columns(2).DataField=   "reaktor"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Bak Luar"
         Columns(3).DataField=   "bak_luar1"
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
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1296"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1217"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
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
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HE0E0E0&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         Height          =   675
         Left            =   7440
         MultiLine       =   -1  'True
         TabIndex        =   13
         Tag             =   "ALKALI"
         Top             =   6840
         Width           =   5625
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
         Height          =   330
         Left            =   8715
         Picture         =   "frmProdAlkali.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "BAHAN"
         Top             =   585
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtReaktor 
         Appearance      =   0  'Flat
         DataField       =   "reaktor"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1275
         TabIndex        =   3
         Tag             =   "ALKALI"
         Top             =   1245
         Width           =   2220
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
         Left            =   2985
         TabIndex        =   10
         Top             =   3045
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
         Left            =   1995
         TabIndex        =   9
         Top             =   3045
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtBerat 
         Appearance      =   0  'Flat
         DataField       =   "berat_rl"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1995
         TabIndex        =   5
         Tag             =   "ALKALI"
         Top             =   2220
         Width           =   1335
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1275
         TabIndex        =   6
         Tag             =   "ALKALI"
         Top             =   915
         Width           =   2220
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Tag             =   "ALKALI"
         Top             =   570
         Width           =   2205
         _ExtentX        =   3889
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
         Format          =   64946179
         CurrentDate     =   39634
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   11
         Tag             =   "ALKALI"
         Top             =   7185
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
         Format          =   64946179
         CurrentDate     =   39584
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   5130
         TabIndex        =   12
         Tag             =   "ALKALI"
         Top             =   7185
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
         Format          =   64946179
         CurrentDate     =   39584
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   2
         Left            =   10920
         TabIndex        =   32
         Tag             =   "ALKALI"
         Top             =   1785
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
         Format          =   64946179
         CurrentDate     =   39584
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   3
         Left            =   10935
         TabIndex        =   33
         Tag             =   "ALKALI"
         Top             =   2130
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
         Format          =   64946179
         CurrentDate     =   39584
      End
      Begin MSComCtl2.DTPicker dateGrid 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   9540
         TabIndex        =   38
         Tag             =   "ALKALI"
         Top             =   3195
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
         Format          =   64946179
         CurrentDate     =   39584
      End
      Begin TrueOleDBGrid80.TDBGrid gridPencucian 
         Height          =   3780
         Left            =   7455
         TabIndex        =   31
         Top             =   2520
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   6668
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
         Caption         =   "Pemindahan Rumput Laut dari Bak Luar Ke Reaktor"
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10935
         TabIndex        =   43
         Top             =   255
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   11235
         X2              =   9315
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label labell 
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
         Index           =   1
         Left            =   9345
         TabIndex        =   42
         Top             =   300
         Width           =   930
      End
      Begin VB.Label lblBeratRL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pH Akhir Alkali"
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
         Left            =   10545
         TabIndex        =   37
         Top             =   6375
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   12810
         X2              =   10545
         Y1              =   6645
         Y2              =   6645
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Selesai Pemidahan RL"
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
         Index           =   1
         Left            =   8610
         TabIndex        =   35
         Top             =   2190
         Width           =   2055
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Mulai Pemindahan RL"
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
         Index           =   0
         Left            =   8610
         TabIndex        =   34
         Top             =   1815
         Width           =   2250
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   10980
         X2              =   8610
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   11820
         X2              =   8595
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   105
         X2              =   13065
         Y1              =   1710
         Y2              =   1710
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
         Left            =   7500
         TabIndex        =   29
         Top             =   6630
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   17
         X1              =   2370
         X2              =   135
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Label labell 
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
         Index           =   0
         Left            =   5040
         TabIndex        =   28
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lblMO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "manufactureorder"
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
         Left            =   6630
         TabIndex        =   8
         Tag             =   "ALKALI"
         Top             =   585
         Width           =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   2400
         X2              =   135
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Label lblNoRL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "no_stock"
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
         Left            =   1995
         TabIndex        =   4
         Tag             =   "ALKALI"
         Top             =   1890
         Width           =   2190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   6030
         X2              =   3660
         Y1              =   7485
         Y2              =   7485
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   2430
         X2              =   60
         Y1              =   7485
         Y2              =   7485
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Mulai Alkali"
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
         Left            =   60
         TabIndex        =   27
         Top             =   7215
         Width           =   2250
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Selesai Alkali"
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
         Left            =   3660
         TabIndex        =   26
         Top             =   7245
         Width           =   2055
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
         Left            =   5025
         TabIndex        =   25
         Top             =   630
         Width           =   1380
      End
      Begin VB.Label lblRekomendasi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "rekomno"
         DataSource      =   "MyDDE"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6630
         TabIndex        =   1
         Tag             =   "ALKALI"
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label lblBeratRL 
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
         Left            =   120
         TabIndex        =   24
         Top             =   1290
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   2460
         X2              =   90
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Label lblKondisi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kondisi"
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
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   23
         Top             =   3060
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   2460
         X2              =   90
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Label lblTempatAlkali 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Alkali Treatment"
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
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   2730
         Width           =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   7255
         X2              =   4995
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label lblBeratRL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat RL"
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
         Left            =   135
         TabIndex        =   21
         Top             =   2250
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   6930
         X2              =   5010
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label lblNoStock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Stock RL"
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
         Left            =   135
         TabIndex        =   20
         Top             =   1935
         Width           =   915
      End
      Begin VB.Label lblNoEkstraksi 
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
         Left            =   105
         TabIndex        =   18
         Top             =   300
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   2400
         X2              =   135
         Y1              =   1215
         Y2              =   1215
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
         Left            =   135
         TabIndex        =   17
         Top             =   990
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1345
         X2              =   105
         Y1              =   1545
         Y2              =   1545
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
         Left            =   105
         TabIndex        =   16
         Top             =   660
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1345
         X2              =   105
         Y1              =   525
         Y2              =   525
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
         TabIndex        =   15
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1345
         X2              =   105
         Y1              =   870
         Y2              =   870
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7590
      Width           =   13215
      _ExtentX        =   23310
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
      TabIndex        =   19
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
Attribute VB_Name = "frmProdAlkali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private MEdit As Boolean
Private RsDetail As New DBQuick
Private RsPencucian As New DBQuick

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1


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
End Function

Private Sub Check1_Click()
   If Check1.Value = 0 Then
      Check2.Value = 1
   Else
      txtTempatAlkali.Text = "Reaktor"
      Check2.Value = 0
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 0 Then
      Check1.Value = 1
   Else
      txtTempatAlkali.Text = "Bak Luar"
      Check1.Value = 0
   End If

End Sub

Private Sub dateGrid_Change()
   If 3 >= gridPencucian.col <= 2 Then
      gridPencucian.Columns(gridPencucian.col) = dateGrid.Value
   End If
End Sub

Private Sub gridPencucian_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   dateGrid.Visible = False
   If (gridPencucian.col = 2) Or (gridPencucian.col = 3) Then
      If Not IsNull(gridPencucian.Columns(gridPencucian.col)) Then
         On Error Resume Next
         dateGrid.Value = gridPencucian.Columns(gridPencucian.col)
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

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm

        Case "ACID"
            lblEkstraksi = mCall.GetFieldByName("RlNo")

        Case "REFERENCE"

            lblRekomendasi.Caption = mCall.GetFieldByName("OrderID")
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
        .BindFormTAG = "ALKALI"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From alkali_treatment"
        .SetPermissions = aksess.MayDo("Alkali Treatment")
    End With

End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
      RsDetail.DBOpen "select * from alkali_detail where no_ekstraksi='" & ParameterString & "' order by ID", CNN
      Set gridDetail.DataSource = RsDetail.DBRecordset
      
      RsPencucian.DBOpen "select * from alkali_pencucian where no_ekstraksi='" & ParameterString & "' order by ID", CNN
      Set gridPencucian.DataSource = RsPencucian.DBRecordset
End Sub


Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Dim IDGen As New IDGenerator
   Dim x As Integer
   
    Select Case AdReasonActiveDb

        Case tmbAddNew

            MEdit = True
            Me.Tag = "baru"
            txtBerat.Enabled = True
            DcTanggal.SetFocus
            
            lblMO = frmProduksi.txtBox(5)
            lblNoRL.Caption = frmProduksi.lblSplNo.Caption
            lblEkstraksi = IDGen.GetID("EKS")
            lblRekomendasi.Caption = frmProduksi.txtBox(0)
            
            For x = 0 To 3
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
   Set IDGen = Nothing
End Sub


Private Sub SaveToStatus()
   Dim rsStatus As New DBQuick
   rsStatus.DBOpen "select * from statusProduksi where noEkstraksi='" & lblEkstraksi & "'", CNN, lckLockBatch
   If rsStatus.DBRecordset.Recordcount > 0 Then
      SendDataToServer "update statusproduksi set status=1 where noEkstraksi='" & lblEkstraksi & "'"
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
                              ",'ALKALI',1,'" & Format(Now, "yyyy-MM-dd") & "')"
   End If
End Sub

Private Sub AddDetail()
   With RsDetail.DBRecordset
      '*** 1
      .AddNew
      .Fields("tahapan") = "Larutan Alkali Bekas"
      .Fields("id_tahapan") = 1
      .Fields("proses") = "Jumlah (liter)"
      
      '*** 2
      .AddNew
      .Fields("tahapan") = "Larutan Alkali Bekas"
      .Fields("id_tahapan") = 1
      .Fields("proses") = "Konsentrasi (%)"
   
      '*** 3
      .AddNew
      .Fields("tahapan") = "Larutan Alkali Bekas"
      .Fields("id_tahapan") = 1
      .Fields("proses") = "Waktu memasukkan alkali bekas"
      
      '*** 4
      .AddNew
      .Fields("tahapan") = "Air bersih untuk penambahan proses alkali"
      .Fields("id_tahapan") = 2
      .Fields("proses") = "Jumlah (liter)"
   
      '*** 5
      .AddNew
      .Fields("tahapan") = "Alkali Baru"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Type Alkali"
      
      '*** 6
      .AddNew
      .Fields("tahapan") = "Alkali Baru"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Jumlah (Kg)"
   
      '*** 7
      .AddNew
      .Fields("tahapan") = "Alkali Baru"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Konsentrasi (%)"
      
      '*** 8
      .AddNew
      .Fields("tahapan") = "Alkali Baru"
      .Fields("id_tahapan") = 3
      .Fields("proses") = "Waktu Memasukkan alkali baru"
   
      
      '*** 10
      .AddNew
      .Fields("tahapan") = "Larutan Alkali Akhir"
      .Fields("id_tahapan") = 4
      .Fields("proses") = "Jumlah (liter)"
   
      '*** 11
      .AddNew
      .Fields("tahapan") = "Larutan Alkali Akhir"
      .Fields("id_tahapan") = 4
      .Fields("proses") = "Konsentrasi (%)"
      
      '*** 12
      .AddNew
      .Fields("tahapan") = "Memasukkan Rumput Laut"
      .Fields("id_tahapan") = 5
      .Fields("proses") = "Suhu Larutan Alkali"
   
      '*** 13
      .AddNew
      .Fields("tahapan") = "Memasukkan Rumput Laut"
      .Fields("id_tahapan") = 5
      .Fields("proses") = "Waktu Memasukkan Rumput laut"
      
      '*** 14
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 2 Jam"
   
      '*** 15
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 2 Jam"
      
      '*** 16
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 3 Jam"
   
      '*** 17
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 4 Jam"
      
      '*** 18
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 5 Jam"
   
      '*** 19
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 6 Jam"
      
      '*** 20
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 7 Jam"
   
      '*** 21
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 8 Jam"
      
      '*** 22
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 9 Jam"
   
      '*** 23
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 10 Jam"
      
      '*** 24
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 11 Jam"
   
      '*** 25
      .AddNew
      .Fields("tahapan") = "Alkali Treatment"
      .Fields("id_tahapan") = 6
      .Fields("proses") = "Suhu Setelah 12 Jam"
      
      
      .MoveFirst
   End With
   Set gridDetail.DataSource = RsDetail.DBRecordset
End Sub


Private Sub AddPencucian()
   Dim x As Integer
   
   With RsPencucian.DBRecordset
      For x = 1 To 6
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
                               
    If MyDDE.ActiveRecordset.Recordcount > 0 Then OpenDetail MyDDE.GetFieldByName("no_ekstraksi")
    Label1.Caption = IIf(IsNull(MyDDE.GetFieldByName("Approved_by")), "", MyDDE.GetFieldByName("Approved_by"))
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            CheckControls Me
            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                PrepareQuery
                
            Else
                MyDDE.IsChildMemberReady = False
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

    If optBersih.Value = True Then
        ket = 1
    Else
        ket = 0
    End If

    With MyDDE
        .PrepareAppend = "INSERT INTO [alkali_treatment] " & _
                           "([No_ekstraksi],[tanggal],[Grup],[Reaktor],[No_Stock],[rekomNo]" & _
                           ",[ManufactureOrder],[Berat_rl],[berat_rl_rekom],[tempat_alkali],[no_rl],[kondisi]" & _
                           ",[waktu_mulai],[waktu_selesai],[keterangan],[waktu_mulai_pemindahan],[waktu_selesai_pemindahan],[ph_akhir]" & _
                           ",[ph_akhir_rekom],issued_by)" & _
                        " Values " & _
                           "('" & lblEkstraksi & "','" & Format(DcTanggal.Value, "yyyy-MM-dd") & "','" & txtGroup & "'" & _
                           ",'" & txtReaktor & "','" & lblNoRL & "','" & lblRekomendasi & "' " & _
                           ",'" & lblMO.Caption & "'," & FQty(txtBerat) & "," & 0 & _
                           ",'" & txtTempatAlkali & "','" & lblNoRL.Caption & "','" & ket & "' " & _
                           ",'" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "','" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "','" & txtKeterangan & "'" & _
                           ",'" & Format(tgl(2).Value, "yyyy-MM-dd hh:mm:ss") & "','" & Format(tgl(3).Value, "yyyy-MM-dd hh:mm:ss") & "'," & FQty(txtPh) & ",0,'" & MainMenu.StatusBar1.Panels(1).Text & "')"
                           

        .PrepareUpdate = "UPDATE [alkali_treatment] SET " & _
                             "[tanggal] = '" & Format(DcTanggal.Value, "yyyy-MM-dd") & "'" & _
                             ",[Grup] ='" & txtGroup & "'" & _
                             ",[Reaktor] = '" & txtReaktor & "'" & _
                             ",[No_Stock] = '" & lblNoRL & "'" & _
                             ",[rekomNo] = '" & lblRekomendasi.Caption & "'" & _
                             ",[ManufactureOrder] = '" & lblMO.Caption & "'" & _
                             ",[Berat_rl] = " & FQty(txtBerat) & _
                             ",[berat_rl_rekom] = " & .GetFieldByName("berat_rl_rekom") & _
                             ",[tempat_alkali] = '" & txtTempatAlkali & "'" & _
                             ",[no_rl] = '" & lblNoRL & "'" & _
                             ",[kondisi] = '" & ket & "'" & _
                             ",[waktu_mulai] = '" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[waktu_selesai] = '" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[keterangan] = '" & txtKeterangan & "'" & _
                             ",[waktu_mulai_pemindahan] = '" & Format(tgl(2).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[waktu_selesai_pemindahan] = '" & Format(tgl(3).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[ph_akhir] = " & FQty(txtPh) & _
                             ",[ph_akhir_rekom] = " & .GetFieldByName("ph_akhir_rekom") & _
                        " WHERE no_ekstraksi='" & lblEkstraksi & "'"
        
        .PrepareDelete = "DELETE From alkali_treatment Where no_ekstraksi='" & lblEkstraksi & "'"
    End With

    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    Err.Clear
End Sub

Private Sub SimpanDetail()
On Error GoTo xErr
   '*** Update Data on Alkali Detail
   SendDataToServer "delete from alkali_detail where no_ekstraksi='" & lblEkstraksi & "'"
   With RsDetail.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO [alkali_detail] ([no_ekstraksi]" & _
                                 ",[tahapan]" & _
                                 ",[proses]" & _
                                 ",[reaktor]" & _
                                 ",[bak_luar1]" & _
                                 ",[reaktor_rekom]" & _
                                 ",[bak_luar1_rekom]" & _
                                 ",id_tahapan )" & _
                           "Values ('" & lblEkstraksi & "'" & _
                                 ",'" & .Fields("tahapan") & "'" & _
                                 ",'" & .Fields("proses") & "'" & _
                                 ",'" & .Fields("reaktor") & "'" & _
                                 ",'" & .Fields("bak_luar1") & "'" & _
                                 ",'" & .Fields("reaktor_rekom") & "'" & _
                                 ",'" & .Fields("bak_luar1_rekom") & "'" & _
                                 ", " & .Fields("id_tahapan") & ")"
         .MoveNext
      Wend
   End With
   
   '*** Update Data on Alkali Pencucian
   SendDataToServer "delete from alkali_pencucian where no_ekstraksi='" & lblEkstraksi & "'"
   
   With RsPencucian.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO [alkali_pencucian] ([no_ekstraksi] " & _
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

Private Function isNum(Keya As Integer)

    If Not (Keya >= Asc("0") And Keya <= Asc("9") Or Keya = vbKeyBack Or Keya = Asc(".")) Then
        Beep
        Keya = 0
    End If
  
End Function

Private Sub txtBerat_KeyPress(KeyAscii As Integer)
    isNum KeyAscii
End Sub

