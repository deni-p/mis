VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmProdEksReaktor 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ekstraksi Di Reaktor"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProdEksReaktor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   13185
   Begin VB.TextBox txtJmlAgarKotor 
      Appearance      =   0  'Flat
      DataField       =   "jml_agar_kotor"
      DataSource      =   "MyDDE"
      Height          =   315
      Left            =   10920
      TabIndex        =   48
      Tag             =   "EX"
      Top             =   3615
      Width           =   1710
   End
   Begin VB.TextBox txtAir2 
      Appearance      =   0  'Flat
      DataField       =   "jml_air_2"
      DataSource      =   "MyDDE"
      Height          =   315
      Left            =   10050
      TabIndex        =   47
      Tag             =   "EX"
      Top             =   2055
      Width           =   1140
   End
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
      Height          =   7065
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   13185
      TabIndex        =   7
      Top             =   0
      Width           =   13185
      Begin VB.TextBox lblEkstraksi 
         Appearance      =   0  'Flat
         DataField       =   "no_ekstraksi"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   57
         Tag             =   "EX"
         Top             =   195
         Width           =   2055
      End
      Begin VB.TextBox txtpHTransfer 
         Appearance      =   0  'Flat
         DataField       =   "ph_transfer"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   10920
         TabIndex        =   50
         Tag             =   "EX"
         Top             =   4665
         Width           =   1710
      End
      Begin VB.TextBox txtSuhuTransfer 
         Appearance      =   0  'Flat
         DataField       =   "suhu_transfer"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   10920
         TabIndex        =   49
         Tag             =   "EX"
         Top             =   4320
         Width           =   1710
      End
      Begin VB.TextBox txtAir1 
         Appearance      =   0  'Flat
         DataField       =   "jml_air_1"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   46
         Tag             =   "EX"
         Top             =   2055
         Width           =   870
      End
      Begin VB.ComboBox cmdDat1 
         DataField       =   "kondisi_DAT1"
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
         ItemData        =   "FrmProdEksReaktor.frx":6852
         Left            =   10950
         List            =   "FrmProdEksReaktor.frx":685C
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Tag             =   "EX"
         Top             =   6450
         Width           =   1635
      End
      Begin VB.ComboBox cmdAutoclave 
         DataField       =   "kondisi_autoclave"
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
         ItemData        =   "FrmProdEksReaktor.frx":686D
         Left            =   10950
         List            =   "FrmProdEksReaktor.frx":6877
         TabIndex        =   44
         Tag             =   "EX"
         Text            =   "cmdAutoclave"
         Top             =   6105
         Width           =   1635
      End
      Begin VB.ComboBox cmbReaktor 
         DataField       =   "kondisi_reaktor"
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
         ItemData        =   "FrmProdEksReaktor.frx":6888
         Left            =   10950
         List            =   "FrmProdEksReaktor.frx":6892
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Tag             =   "EX"
         Top             =   5760
         Width           =   1635
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3600
         Left            =   120
         TabIndex        =   20
         Top             =   2670
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   6350
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BackColor       =   15380335
         TabCaption(0)   =   "pH"
         TabPicture(0)   =   "FrmProdEksReaktor.frx":68A3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line1(6)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Line1(7)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label5(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblTanggal(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Line1(8)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "DtpMendidih"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "gridPh"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtPhAwal"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtPhAkhir"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Kondisi Rumput Laut"
         TabPicture(1)   =   "FrmProdEksReaktor.frx":68BF
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gridRL"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Penambahan Bahan Penunjang"
         TabPicture(2)   =   "FrmProdEksReaktor.frx":68DB
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "gridPenambahan"
         Tab(2).Control(1)=   "DTPWaktu"
         Tab(2).ControlCount=   2
         Begin TrueOleDBGrid80.TDBGrid gridRL 
            Height          =   3135
            Left            =   -74940
            TabIndex        =   54
            Top             =   375
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   5530
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Waktu"
            Columns(0).DataField=   "waktu"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   1
            Columns(1)._MaxComboItems=   5
            Columns(1).ValueItems(0)._DefaultItem=   0
            Columns(1).ValueItems(0).Value=   "Keras"
            Columns(1).ValueItems(0).Value.vt=   8
            Columns(1).ValueItems(0).DisplayValue=   "Keras"
            Columns(1).ValueItems(0).DisplayValue.vt=   8
            Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(1).ValueItems(1)._DefaultItem=   0
            Columns(1).ValueItems(1).Value=   "Lunak"
            Columns(1).ValueItems(1).Value.vt=   8
            Columns(1).ValueItems(1).DisplayValue=   "Lunak"
            Columns(1).ValueItems(1).DisplayValue.vt=   8
            Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(1).ValueItems(2)._DefaultItem=   0
            Columns(1).ValueItems(2).Value=   "Hancur"
            Columns(1).ValueItems(2).Value.vt=   8
            Columns(1).ValueItems(2).DisplayValue=   "Hancur"
            Columns(1).ValueItems(2).DisplayValue.vt=   8
            Columns(1).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
            Columns(1).ValueItems(3)._DefaultItem=   0
            Columns(1).ValueItems(3).Value=   "Tidak Hancur"
            Columns(1).ValueItems(3).Value.vt=   8
            Columns(1).ValueItems(3).DisplayValue=   "Tidak Hancur"
            Columns(1).ValueItems(3).DisplayValue.vt=   8
            Columns(1).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
            Columns(1).ValueItems.Count=   4
            Columns(1).Caption=   "Kondisi"
            Columns(1).DataField=   "kondisi"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0)._GSX_SAVERECORDSELECTORS=   0
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=6641"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6562"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=4736"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4657"
            Splits(0)._ColumnProps(8)=   "Column(1).Button=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin MSComCtl2.DTPicker DTPWaktu 
            Height          =   345
            Left            =   -70590
            TabIndex        =   53
            Top             =   1605
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   609
            _Version        =   393216
            Format          =   56164354
            CurrentDate     =   39651
         End
         Begin VB.TextBox txtPhAkhir 
            Appearance      =   0  'Flat
            DataField       =   "ph_akhir"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   1590
            TabIndex        =   52
            Tag             =   "EX"
            Top             =   3150
            Width           =   1170
         End
         Begin VB.TextBox txtPhAwal 
            Appearance      =   0  'Flat
            DataField       =   "ph_awal"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   1410
            TabIndex        =   51
            Tag             =   "EX"
            Top             =   510
            Width           =   1230
         End
         Begin MSDataGridLib.DataGrid gridPenambahan 
            Height          =   3105
            Left            =   -74940
            TabIndex        =   22
            Top             =   405
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   5477
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "nama_bahan"
               Caption         =   "Nama Bahan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "qty"
               Caption         =   "Qty"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "satuan"
               Caption         =   "Satuan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "waktu"
               Caption         =   "waktu"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   4
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gridPh 
            Height          =   2085
            Left            =   135
            TabIndex        =   21
            Top             =   960
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3678
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "kondisi"
               Caption         =   "Kondisi"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "nilai"
               Caption         =   "Nilai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DtpMendidih 
            DataField       =   "waktu_didih"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   6345
            TabIndex        =   27
            Tag             =   "EX"
            Top             =   510
            Width           =   1755
            _ExtentX        =   3096
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
            Format          =   56164354
            CurrentDate     =   39634
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   8
            X1              =   6420
            X2              =   4365
            Y1              =   810
            Y2              =   810
         End
         Begin VB.Label lblTanggal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Masakan Mendidih"
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
            Left            =   4350
            TabIndex        =   28
            Top             =   540
            Width           =   1815
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pH Awal Masak"
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
            Left            =   165
            TabIndex        =   26
            Top             =   555
            Width           =   1080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   7
            X1              =   1755
            X2              =   165
            Y1              =   810
            Y2              =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pH Akhir Ekstraksi"
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
            Left            =   180
            TabIndex        =   25
            Top             =   3210
            Width           =   1275
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   6
            X1              =   1770
            X2              =   180
            Y1              =   3465
            Y2              =   3465
         End
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         Height          =   675
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "EX"
         Top             =   6360
         Width           =   8205
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Tag             =   "EX"
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtReaktor 
         Appearance      =   0  'Flat
         DataField       =   "reaktor"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Tag             =   "EX"
         Top             =   1245
         Width           =   1710
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Tag             =   "EX"
         Top             =   540
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
         Format          =   56164355
         CurrentDate     =   39634
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_mulai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   6300
         TabIndex        =   4
         Tag             =   "EX"
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
         Format          =   56164355
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_selesai"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   6300
         TabIndex        =   5
         Tag             =   "EX"
         Top             =   2190
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
         Format          =   56164355
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "waktu_tambah_air"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   2
         Left            =   10050
         TabIndex        =   31
         Tag             =   "EX"
         Top             =   2400
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
         Format          =   56164355
         CurrentDate     =   39419
      End
      Begin MSComCtl2.DTPicker tgl 
         DataField       =   "tanggal_transfer"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   3
         Left            =   10920
         TabIndex        =   35
         Tag             =   "EX"
         Top             =   3960
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
         Format          =   56164355
         CurrentDate     =   39419
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
         Left            =   7425
         TabIndex        =   55
         Top             =   915
         Width           =   1980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   22
         X1              =   7485
         X2              =   5295
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   5295
         TabIndex        =   56
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DAT1"
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
         Left            =   8865
         TabIndex        =   42
         Top             =   6480
         Width           =   390
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   21
         X1              =   10965
         X2              =   8865
         Y1              =   6735
         Y2              =   6735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autoclave"
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
         Left            =   8865
         TabIndex        =   41
         Top             =   6135
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   20
         X1              =   10965
         X2              =   8865
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Label Label16 
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
         Left            =   8865
         TabIndex        =   40
         Top             =   5805
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   19
         X1              =   10965
         X2              =   8865
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KONDISI SCREEN SETELAH TRANSFER HASIL EKSTRAKSI"
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
         Left            =   8865
         TabIndex        =   39
         Top             =   5445
         Width           =   4110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pH Pada Saat Transfer"
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
         Left            =   8865
         TabIndex        =   38
         Top             =   4710
         Width           =   1635
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   18
         X1              =   10965
         X2              =   8865
         Y1              =   4965
         Y2              =   4965
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suhu Pada Saat Transfer"
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
         Left            =   8865
         TabIndex        =   37
         Top             =   4365
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   17
         X1              =   10965
         X2              =   8865
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal && Waktu Transfer"
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
         Left            =   8865
         TabIndex        =   36
         Top             =   4005
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   11235
         X2              =   8865
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Agar Kotor (Liter)"
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
         Left            =   8865
         TabIndex        =   34
         Top             =   3660
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   10965
         X2              =   8865
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSFER KE DAT1"
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
         Left            =   8865
         TabIndex        =   33
         Top             =   3360
         Width           =   1440
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
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
         Left            =   8865
         TabIndex        =   32
         Top             =   2430
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   11235
         X2              =   8865
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah (Liter)"
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
         Left            =   8865
         TabIndex        =   30
         Top             =   2100
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   10105
         X2              =   8865
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JML AIR BERSIH UNTUK PENAMBAHAN PROSES EKSTRAKSI"
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
         Left            =   8865
         TabIndex        =   29
         Top             =   1860
         Width           =   4260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah (Liter)"
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
         TabIndex        =   24
         Top             =   2085
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   1375
         X2              =   135
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Air Bersih untuk proses Ekstraksi"
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
         TabIndex        =   23
         Top             =   1845
         Width           =   2595
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   105
         X2              =   13065
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Label lblRekom 
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
         Left            =   7425
         TabIndex        =   19
         Tag             =   "EX"
         Top             =   195
         Width           =   1980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   7635
         X2              =   5280
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label lblRekomendasi 
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
         Left            =   5295
         TabIndex        =   18
         Top             =   240
         Width           =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   6510
         X2              =   4140
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   6510
         X2              =   4140
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Mulai Masak"
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
         Left            =   4140
         TabIndex        =   17
         Top             =   1860
         Width           =   1890
      End
      Begin VB.Label lblTanggalWaktu 
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Selesai Masak"
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
         Left            =   4140
         TabIndex        =   16
         Top             =   2220
         Width           =   2055
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
         Left            =   7425
         TabIndex        =   15
         Tag             =   "EX"
         Top             =   555
         Width           =   1980
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
         Left            =   5295
         TabIndex        =   14
         Top             =   600
         Width           =   1380
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
         Left            =   120
         TabIndex        =   12
         Top             =   255
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   7485
         X2              =   5295
         Y1              =   855
         Y2              =   855
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
         TabIndex        =   11
         Top             =   960
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   1360
         X2              =   120
         Y1              =   1200
         Y2              =   1200
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
         TabIndex        =   10
         Top             =   600
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
         TabIndex        =   9
         Top             =   7560
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1360
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   2370
         X2              =   120
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Label lblReaktor 
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
         TabIndex        =   8
         Top             =   1305
         Width           =   570
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7065
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   1005
      BindFormTAG     =   "EX"
      InitControlSet  =   1
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
      TabIndex        =   13
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
Attribute VB_Name = "FrmProdEksReaktor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Xval As String
Dim MEdit As Boolean

Dim RsPH As New DBQuick
Dim RsRL As New DBQuick
Dim RsPenambahan As New DBQuick

Dim GridAltColor As String
Dim Changingsel As Byte

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RcProduksi As New DBQuick
Dim RcPartner As New DBQuick


Private Sub OPenPartner(index As Integer)
On Error GoTo Hell:
Select Case index
       Case 0:
            RcPartner.DBOpen " SELECT noItem,InternalName as [Nama Bahan], UOM as Satuan from inventory where substring(noItem,4,2)='CH'", CNN, lckLockReadOnly
            
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case index
          Case 0:
            mCall.FromTagActive = "Bahan"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Bahan Tidak Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
    Err.Clear
End Sub


Private Sub DTPWaktu_Change()
   gridPenambahan.Columns(3).Value = DTPWaktu.Value
End Sub

Private Sub Form_Load()
    HiasFormManTell Picture2, Me
    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "EX"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From extraction "
        .SetPermissions = aksess.MayDo("Ekstraksi di Reaktor")
    End With
   
   gridPh.HeadLines = 2
   gridRL.HeadLines = 2
   gridPenambahan.HeadLines = 2
   gridRL.RowHeight = 300
   gridPenambahan.RowHeight = 300
   
   Set mCall = New frmCaller
End Sub


Private Sub OpenDetail(ByVal ParameterString As String)

   If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
 
   '*** Load Data Ph saat masak
   Set RsPH = New DBQuick
   RsPH.DBOpen "select * from extraction_ph where no_ekstraksi='" & ParameterString & "'", CNN, lckLockBatch
   Set gridPh.DataSource = RsPH.DBRecordset
   
   
   '*** Load Data Kondisi RL
   Set RsRL = New DBQuick
   RsRL.DBOpen "select * from extraction_kondisi_rl where no_ekstraksi='" & ParameterString & "'", CNN, lckLockBatch
   Set gridRL.DataSource = RsRL.DBRecordset
   
   
   '*** Load Penambahan RL
   Set RsPenambahan = New DBQuick
   RsPenambahan.DBOpen "select * from extraction_penambahan where  no_ekstraksi='" & ParameterString & "'", CNN, lckLockBatch
   Set MyDDE.ChildRecordset = RsPenambahan.DBRecordset
   Set gridPenambahan.DataSource = MyDDE.ChildRecordset
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mCall = Nothing
End Sub

Private Sub gridPenambahan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   DTPWaktu.Visible = False
   If gridPenambahan.col = 3 Then
      On Error Resume Next
      DTPWaktu.Move gridPenambahan.Left + gridPenambahan.Columns(3).Left, _
                    gridPenambahan.Top + gridPenambahan.RowTop(gridPenambahan.row), _
                    gridPenambahan.Columns(3).width, _
                    gridPenambahan.RowHeight
      DTPWaktu.Value = gridPenambahan.Columns(3)
      DTPWaktu.Visible = True
   End If
End Sub


Private Sub lblEkstraksi_LostFocus()
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & lblEkstraksi.Text & "'", CNN, lckLockBatch
   If rsCek.DBRecordset.Recordcount > 0 Then
      rsCek.DBOpen "select * from extraction where no_Ekstraksi='" & lblEkstraksi.Text & "'", CNN, lckLockBatch
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

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   RsPenambahan.DBRecordset.Fields("noItem") = mCall.GetFieldByName("NOiTEM")
   RsPenambahan.DBRecordset.Fields("nama_bahan") = mCall.GetFieldByName("nama Bahan")
   RsPenambahan.DBRecordset.Fields("Satuan") = mCall.GetFieldByName("satuan")
   RsPenambahan.DBRecordset.Fields("waktu") = Now
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)

    If (MyDDE.ActiveRecordset.BOF = False) And (MyDDE.ActiveRecordset.EOF = False) Then OpenDetail MyDDE.ActiveRecordset.Fields("no_ekstraksi")
   Label1.Caption = IIf(IsNull(MyDDE.GetFieldByName("Approved_by")), "", MyDDE.GetFieldByName("Approved_by"))
End Sub

    
Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            SimpanDetail
            SaveToMO

        Case tmbAddNew
            MEdit = True
            DcTanggal.SetFocus
            
            lblMO = frmProduksi.txtBox(5)
            'lblEkstraksi = frmProduksi.txtBox(1)
            lblRekom.Caption = frmProduksi.txtBox(0)
            txtKeterangan.Text = "-"
            
            DcTanggal.Value = Now
            tgl(0).Value = Now
            tgl(1).Value = Now
            tgl(2).Value = Now
            tgl(3).Value = Now
            DtpMendidih.Value = Now
            
            LoadPh
            LoadKondisiRL
            
        Case tmbDelete
            PrepareQuery
            
        Case tmbDetail
            If SSTab1.Tab = 2 Then
               OPenPartner 0
            End If
    End Select

End Sub

Private Sub LoadPh()
   With RsPH.DBRecordset
      .AddNew
      .Fields("kondisi") = "pH pada 50°C"
      
      .AddNew
      .Fields("kondisi") = "pH pada 80°C"
      
      .AddNew
      .Fields("kondisi") = "pH pada 100°C"
   
      .AddNew
      .Fields("kondisi") = "pH pada 30 Menit"
   
      .AddNew
      .Fields("kondisi") = "pH pada 60 Menit"
   
      .AddNew
      .Fields("kondisi") = "pH pada 90 Menit"
   
      .AddNew
      .Fields("kondisi") = "pH pada 120 Menit"
   
      .AddNew
      .Fields("kondisi") = "pH pada 150 Menit"
   
      .AddNew
      .Fields("kondisi") = "pH pada 180 Menit"
      
      .MoveFirst
   End With
End Sub

Private Sub LoadKondisiRL()
   With RsRL.DBRecordset
      .AddNew
      .Fields("waktu") = "Kondisi Rumput Laut pada 30 Menit"
   
      .AddNew
      .Fields("waktu") = "Kondisi Rumput Laut pada 60 Menit"
   
      .AddNew
      .Fields("waktu") = "Kondisi Rumput Laut pada 90 Menit"
   
      .AddNew
      .Fields("waktu") = "Kondisi Rumput Laut pada 120 Menit"
   
      .AddNew
      .Fields("waktu") = "Kondisi Rumput Laut pada 150 Menit"
   
      .AddNew
      .Fields("waktu") = "Kondisi Rumput Laut pada 180 Menit"
      
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
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 40", CNN

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

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave
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


Private Sub PrepareQuery()
    On Error GoTo Masjid
    Dim ket As Byte


    With MyDDE
        .PrepareAppend = "INSERT INTO [extraction] ([no_ekstraksi],[tanggal],[grup],[reaktor] ,[rekomendasi] " & _
                              ",[manufacture_order],[jml_air_1],[waktu_mulai],[waktu_selesai],[jml_air_2] " & _
                              ",[waktu_tambah_air],[jml_agar_kotor],[tanggal_transfer],[suhu_transfer],[ph_awal] " & _
                              ",[ph_transfer],[ph_akhir],[waktu_didih],[kondisi_reaktor],[kondisi_autoclave] " & _
                              ",[kondisi_DAT1],[keterangan],issued_by) " & _
                         "Values ('" & lblEkstraksi & "'" & _
                              ",'" & Format(DcTanggal.Value, "yyyy-MM-dd") & "'" & _
                              ",'" & txtGroup & "','" & txtReaktor & "'" & _
                              ",'" & lblRekom.Caption & "','" & lblMO.Caption & "'" & _
                              ", " & FQty(txtAir1) & _
                              ",'" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                              ",'" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                              ", " & FQty(txtAir2) & _
                              ",'" & Format(tgl(2).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                              ", " & FQty(txtJmlAgarKotor) & _
                              ",'" & Format(tgl(3).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                              ", " & FQty(txtSuhuTransfer) & _
                              ", " & FQty(txtPhAwal) & _
                              ", " & FQty(txtpHTransfer) & _
                              ", " & FQty(txtPhAkhir) & _
                              ",'" & Format(DtpMendidih.Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                              ",'" & cmbReaktor.Text & "'" & _
                              ",'" & cmdAutoclave.Text & "'" & _
                              ",'" & cmdDat1.Text & "'" & _
                              ",'" & txtKeterangan & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
        
        .PrepareUpdate = "UPDATE [extraction] SET [tanggal] = '" & Format(DcTanggal.Value, "yyyy-MM-dd") & "'" & _
                             ",[grup] = '" & txtGroup & "'" & _
                             ",[reaktor] = '" & txtReaktor & "'" & _
                             ",[rekomendasi] = '" & lblRekom.Caption & "'" & _
                             ",[manufacture_order] = '" & lblMO.Caption & "'" & _
                             ",[jml_air_1] = " & FQty(txtAir1) & _
                             ",[waktu_mulai] = '" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[waktu_selesai] = '" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[jml_air_2] = " & FQty(txtAir2) & _
                             ",[waktu_tambah_air] = '" & Format(tgl(2).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[jml_agar_kotor] = " & FQty(txtJmlAgarKotor) & _
                             ",[tanggal_transfer] = '" & Format(tgl(3).Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[suhu_transfer] = " & FQty(txtSuhuTransfer) & _
                             ",[ph_awal] = " & FQty(txtPhAwal) & _
                             ",[ph_transfer] = " & FQty(txtpHTransfer) & _
                             ",[ph_akhir] = " & FQty(txtPhAkhir) & _
                             ",[waktu_didih] = '" & Format(DtpMendidih.Value, "yyyy-MM-dd hh:mm:ss") & "'" & _
                             ",[kondisi_reaktor] = '" & cmbReaktor.Text & "'" & _
                             ",[kondisi_autoclave] = '" & cmdAutoclave.Text & "'" & _
                             ",[kondisi_DAT1] = '" & cmdDat1.Text & "'" & _
                             ",[keterangan] = '" & txtKeterangan.Text & "' " & _
                        " WHERE no_ekstraksi='" & lblEkstraksi & "'"
        
        
        .PrepareDelete = "DELETE From extraction Where no_ekstraksi = '" & lblEkstraksi & "'"
    End With

    Exit Sub
Masjid:
    MessageBox Err.Description, "Peringatan", msgOkOnly, msgExclamation
    
    Err.Clear
End Sub

Private Sub SimpanDetail()
   '*** Saving to extraction ph
   If SendDataToServer("delete from extraction_ph where no_ekstraksi='" & lblEkstraksi & "'") Then
      With RsPH.DBRecordset
         .MoveFirst
         While Not .EOF
            SendDataToServer "INSERT INTO [extraction_ph] ([no_ekstraksi],[kondisi],[nilai]) " & _
                             " Values ('" & lblEkstraksi & "'" & _
                                      ",'" & .Fields("kondisi") & "'" & _
                                      ",'" & .Fields("nilai") & "')"
            .MoveNext
         Wend
      End With
   End If
   
   '*** SAving to Extraction KOndisi RL
   If SendDataToServer("delete from extraction_kondisi_rl where no_ekstraksi='" & lblEkstraksi & "'") Then
      With RsRL.DBRecordset
         .MoveFirst
         While Not .EOF
            SendDataToServer "INSERT INTO [extraction_kondisi_rl]([no_ekstraksi],[waktu] ,[kondisi]) " & _
                             " Values ('" & lblEkstraksi & "'" & _
                                       ",'" & .Fields("waktu") & "'" & _
                                       ",'" & .Fields("kondisi") & "')"
            .MoveNext
         Wend
      End With
   End If
   
   
   '*** Saving to extraction penambahan bahan penunjang
   If SendDataToServer("delete from extraction_penambahan where no_ekstraksi='" & lblEkstraksi & "'") Then
      With MyDDE.ChildRecordset
         .MoveFirst
         While Not .EOF
            SendDataToServer "INSERT [extraction_penambahan] " & _
                                    "([no_ekstraksi]" & _
                                    ",[nama_bahan]" & _
                                    ",[qty]" & _
                                    ",[satuan]" & _
                                    ",[waktu])" & _
                              "Values ('" & lblEkstraksi & "'" & _
                                    ",'" & .Fields("nama_bahan") & "'" & _
                                    ", " & FQty(.Fields("qty")) & _
                                    ",'" & .Fields("satuan") & "'" & _
                                    ",'" & Format(.Fields("waktu"), "yyyy-MM-dd hh:mm:ss") & "')"
            .MoveNext
         Wend
      End With
   End If
   
End Sub
