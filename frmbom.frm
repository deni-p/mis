VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmbom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BOM"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmbom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Tag             =   "Bill of Material"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1005
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5490
      Left            =   15
      ScaleHeight     =   5490
      ScaleWidth      =   9975
      TabIndex        =   10
      Top             =   -30
      Width           =   9975
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "BomReff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6075
         TabIndex        =   40
         Tag             =   "Partner"
         Text            =   "#Bom-00001"
         Top             =   120
         Width           =   1500
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Phantom"
         DataField       =   "Phantom"
         ForeColor       =   &H80000004&
         Height          =   225
         Left            =   6090
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   855
         Width           =   1290
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "BOM Id"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1395
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   120
         Width           =   3000
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "UOM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1395
         MaxLength       =   15
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   1185
         Width           =   3000
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   480
         Width           =   3000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Qty Availlable"
         Height          =   330
         Left            =   4695
         TabIndex        =   8
         Top             =   1185
         Width           =   1470
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataField       =   "Status"
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
         ItemData        =   "frmbom.frx":6852
         Left            =   6060
         List            =   "frmbom.frx":6862
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   495
         Width           =   1875
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   7590
         Picture         =   "frmbom.frx":688A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   350
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "MethodeID"
         Height          =   315
         Left            =   1395
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   840
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "MethodeID"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3600
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6350
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         BackColor       =   15380335
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "WorkCenter"
         TabPicture(0)   =   "frmbom.frx":6C14
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Material"
         TabPicture(1)   =   "frmbom.frx":6C30
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture5"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "BOM Structure"
         TabPicture(2)   =   "frmbom.frx":6C4C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TGBom"
         Tab(2).Control(1)=   "semeruBOM"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Routing"
         TabPicture(3)   =   "frmbom.frx":6C68
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture1"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Bill of Manufacture"
         TabPicture(4)   =   "frmbom.frx":6C84
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txt(2)"
         Tab(4).Control(1)=   "txt(5)"
         Tab(4).Control(2)=   "txt(3)"
         Tab(4).Control(3)=   "txt(0)"
         Tab(4).Control(4)=   "txt(4)"
         Tab(4).Control(5)=   "txt(1)"
         Tab(4).Control(6)=   "SemeruTree1"
         Tab(4).Control(7)=   "Line1(16)"
         Tab(4).Control(8)=   "Label1(15)"
         Tab(4).Control(9)=   "lblUOM"
         Tab(4).Control(10)=   "Label1(16)"
         Tab(4).Control(11)=   "Line1(17)"
         Tab(4).Control(12)=   "Label1(17)"
         Tab(4).Control(13)=   "Line1(18)"
         Tab(4).Control(14)=   "Label1(18)"
         Tab(4).Control(15)=   "Line1(19)"
         Tab(4).Control(16)=   "Label1(19)"
         Tab(4).Control(17)=   "Line1(20)"
         Tab(4).Control(18)=   "Label1(20)"
         Tab(4).Control(19)=   "Line1(21)"
         Tab(4).ControlCount=   20
         Begin TrueOleDBGrid80.TDBGrid TGBom 
            Height          =   3075
            Left            =   -72330
            TabIndex        =   45
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5424
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Item Barang"
            Columns(0).DataField=   "noItem"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Komponen"
            Columns(1).DataField=   "component"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Description"
            Columns(2).DataField=   "Description"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Terpakai"
            Columns(3).DataField=   "Usage Qty"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0)._GSX_SAVERECORDSELECTORS=   0
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(0).Merge=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2408"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2328"
            Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(10)=   "Column(2).Width=5371"
            Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=5292"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=1402"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1323"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
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
            HeadLines       =   1
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
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
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
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00EAAF6F&
            Height          =   3180
            Left            =   -74925
            ScaleHeight     =   3120
            ScaleWidth      =   9540
            TabIndex        =   41
            Top             =   360
            Width           =   9600
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frmbom.frx":6CA0
               Height          =   3135
               Index           =   2
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   9540
               _ExtentX        =   16828
               _ExtentY        =   5530
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
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
               ColumnCount     =   8
               BeginProperty Column00 
                  DataField       =   "noLine"
                  Caption         =   "SeqNo"
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
                  DataField       =   "SeqStageID"
                  Caption         =   "Work Center"
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
               BeginProperty Column02 
                  DataField       =   "Keterangan"
                  Caption         =   "Description"
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
                  DataField       =   "unit_run"
                  Caption         =   "Unit Run"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Setup_Time"
                  Caption         =   "Setup Time"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "Queue_Time"
                  Caption         =   "Queue Time"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "Wait_Time"
                  Caption         =   "Wait Time"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "Total_Run_Time"
                  Caption         =   "Total Run Time"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column06 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column07 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Catatan"
               Height          =   930
               Index           =   5
               Left            =   2685
               MaxLength       =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Text            =   "frmbom.frx":6CB5
               Top             =   3510
               Width           =   6360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
               Height          =   210
               Index           =   3
               Left            =   1365
               TabIndex        =   43
               Top             =   4170
               Width           =   630
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   1365
               X2              =   2790
               Y1              =   4425
               Y2              =   4425
            End
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   -68835
            TabIndex        =   31
            Text            =   " - Kategori -"
            Top             =   1545
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   -68835
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Text            =   " - Lead Time -"
            Top             =   2940
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   -68835
            TabIndex        =   29
            Text            =   " - Kelompok -"
            Top             =   2010
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   -68835
            TabIndex        =   28
            Text            =   " - Kode Barang -"
            Top             =   615
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   -68835
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Text            =   " - Jumlah -"
            Top             =   2475
            Width           =   2115
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   -68835
            TabIndex        =   26
            Text            =   " - Nama Barang -"
            Top             =   1080
            Width           =   3000
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00EAAF6F&
            Height          =   3180
            Left            =   -74925
            ScaleHeight     =   3120
            ScaleWidth      =   9495
            TabIndex        =   15
            Top             =   375
            Width           =   9555
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Keterangan"
               Height          =   930
               Index           =   4
               Left            =   2685
               MaxLength       =   15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Text            =   "frmbom.frx":6CBB
               Top             =   3510
               Width           =   6360
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   3135
               Index           =   1
               Left            =   0
               TabIndex        =   17
               Top             =   -15
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   5530
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
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
               ColumnCount     =   6
               BeginProperty Column00 
                  DataField       =   "WCID"
                  Caption         =   "Work Center"
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
                  DataField       =   "SeqStageID"
                  Caption         =   "Stage"
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
               BeginProperty Column02 
                  DataField       =   "Komponen ID"
                  Caption         =   "Kode Barang"
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
                  DataField       =   "Nama Komponen"
                  Caption         =   "Nama Barang"
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
               BeginProperty Column04 
                  DataField       =   "UOM"
                  Caption         =   "UOM"
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
               BeginProperty Column05 
                  DataField       =   "QTYUsage"
                  Caption         =   "QTY"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "#,##0;(#,##0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
               Height          =   210
               Index           =   5
               Left            =   1365
               TabIndex        =   18
               Top             =   4170
               Width           =   630
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   1365
               X2              =   2790
               Y1              =   4425
               Y2              =   4425
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00EAAF6F&
            Height          =   3180
            Left            =   75
            ScaleHeight     =   3120
            ScaleWidth      =   9540
            TabIndex        =   11
            Top             =   360
            Width           =   9600
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Catatan"
               Height          =   930
               Index           =   3
               Left            =   2685
               MaxLength       =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Text            =   "frmbom.frx":6CC1
               Top             =   3510
               Width           =   6360
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frmbom.frx":6CC7
               Height          =   3120
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   9540
               _ExtentX        =   16828
               _ExtentY        =   5503
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
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
                  DataField       =   "NoLine"
                  Caption         =   "Seq. No"
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
                  DataField       =   "SeqStageID"
                  Caption         =   "Work Center ID"
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
               BeginProperty Column02 
                  DataField       =   "Keterangan"
                  Caption         =   "Description"
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
                  DataField       =   "ResourcesID"
                  Caption         =   "Resources"
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
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
               EndProperty
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   1365
               X2              =   2790
               Y1              =   4425
               Y2              =   4425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
               Height          =   210
               Index           =   8
               Left            =   1365
               TabIndex        =   14
               Top             =   4170
               Width           =   630
            End
         End
         Begin SemeruDC.SemeruTree SemeruTree1 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   25
            Top             =   435
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   5318
            BackColorTree   =   7159830
            BackColorBackground=   -2147483648
         End
         Begin SemeruDC.SemeruTree semeruBOM 
            Height          =   3060
            Left            =   -74880
            TabIndex        =   39
            Top             =   360
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   5398
            BackColorTree   =   7159830
            BackColorBackground=   -2147483648
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   -70065
            X2              =   -68640
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori"
            Height          =   195
            Index           =   15
            Left            =   -70065
            TabIndex        =   38
            Top             =   1605
            Width           =   600
         End
         Begin VB.Label lblUOM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PCS"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   -66675
            TabIndex        =   37
            Top             =   2475
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Time"
            Height          =   195
            Index           =   16
            Left            =   -70065
            TabIndex        =   36
            Top             =   3000
            Width           =   720
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   -70065
            X2              =   -68640
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
            Height          =   195
            Index           =   17
            Left            =   -70065
            TabIndex        =   35
            Top             =   1140
            Width           =   960
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   -70065
            X2              =   -68640
            Y1              =   1380
            Y2              =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
            Height          =   195
            Index           =   18
            Left            =   -70065
            TabIndex        =   34
            Top             =   675
            Width           =   915
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   -70050
            X2              =   -68625
            Y1              =   915
            Y2              =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelompok"
            Height          =   195
            Index           =   19
            Left            =   -70065
            TabIndex        =   33
            Top             =   2070
            Width           =   675
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   -70065
            X2              =   -68640
            Y1              =   2310
            Y2              =   2310
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            Height          =   195
            Index           =   20
            Left            =   -70065
            TabIndex        =   32
            Top             =   2535
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   -70065
            X2              =   -68640
            Y1              =   2775
            Y2              =   2775
         End
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   4680
         X2              =   6795
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   24
         Top             =   548
         Width           =   960
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   165
         X2              =   1590
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   23
         Top             =   188
         Width           =   915
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   1605
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Metode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   22
         Top             =   900
         Width           =   540
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   165
         X2              =   1590
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   21
         Top             =   1260
         Width           =   510
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   165
         X2              =   1590
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BOM Referensi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   4695
         TabIndex        =   20
         Top             =   188
         Width           =   1065
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   4680
         X2              =   6930
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   4695
         TabIndex        =   19
         Top             =   555
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmbom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAdd As Boolean
Private RcPart As New DBQuick
Private Rc As New DBQuick
Private RcComponent As New DBQuick
Private RcBOMStruc As New DBQuick
Private RsTeeBOM As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mFirstCaller, mvarDelete, mVarCallTab As Boolean
Private mKeyNode As Double
Private vOldKey As String

Private Sub cmdLink_Click(Index As Integer)
On Error GoTo 1
Dim I As Integer
Select Case Index
       Case 0:
            I = MessageBox("Anda akan melakukan penambahan BOM Reference?", "BOM Referense", msgYesNo)
            If I = 1 Then
               MyDDE.GetFieldByName("BomReff") = IndexAuto
               OpenDetailComponent txtBox(0)
               OpenDetail txtBox(0)
               SSTab1.Tab = 0
               mVarCallTab = True
            Else
               OpenDetailPartner 5
               mVarCallTab = False
            End If
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbom:click" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Command1_Click()
FrmAvaillable.SetNoItem(txtBox(2)) = txtBox(0)
FrmAvaillable.Show vbModal
End Sub

Private Sub DataCombo1_GotFocus()
cmdLink(0).Enabled = DataCombo1.Enabled
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataGrid1_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
On Error GoTo 1
If mAdd = False Then Exit Sub
Select Case Index
       Case 0:
            If DataGrid1(Index).col = 1 Then
               OpenDetailPartner Index
            ElseIf DataGrid1(Index).col = 3 Then
               OpenDetailPartner 3
            End If
       Case 1:
            If DataGrid1(Index).col = 0 Then
               OpenDetailPartner 6
            ElseIf DataGrid1(Index).col = 1 Then
               OpenDetailPartner 0
            ElseIf DataGrid1(Index).col = 2 Then
               OpenDetailPartner 2
            ElseIf DataGrid1(Index).col = 3 Then
                FrmAvaillable.SetNoItem(DataGrid1(Index).Columns(2).Value) = DataGrid1(Index).Columns(1).Value
                FrmAvaillable.Show vbModal
            End If
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbom:buttonclick" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If DataCombo1.Enabled = True Then ScanKeyGrid KeyCode, Shift, MyDDE
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo 2
If mAdd = True Then
   DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
   Select Case Index
          Case 0:
               Select Case DataGrid1(Index).col
                      Case 0, 3:
                           DataGrid1(Index).AllowUpdate = mAdd
                      Case Else
                           
                           DataGrid1(Index).AllowUpdate = False
               End Select
          Case 1:
               Select Case DataGrid1(Index).col
                      Case 0, 1, 2, 3, 4:
                           DataGrid1(Index).AllowUpdate = False
                      Case 5:
                           DataGrid1(Index).AllowUpdate = True
               End Select
   End Select
   MoveCtrl
Else
   DataGrid1(0).Columns(1).Button = False
   DataGrid1(1).Columns(0).Button = False
   DataGrid1(1).Columns(1).Button = False
   DataGrid1(1).Columns(2).Button = False
   DataGrid1(Index).AllowUpdate = mAdd
End If
Exit Sub
2:
MessageBox Err.Description, "frmbom:rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
On Error GoTo 1
    'HiasForm Picture1, Me
    HiasFormManTell Picture2, Me
    GridLayout
    Check1.BackColor = &HEAAF6F
    SSTab1.Tab = 0
    OpenPartner

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = frmbom
        .BindFormTAG = "Partner"
        Set .ActiveConnection = CNN
        .PrepareQuery = "SELECT NoItem AS [BOM Id], ItemName AS Keterangan, UOM AS UOM, MethodeID, Phantom ,Status,BomReff " & " FROM Inventory WHERE (Manufacture = 1) ORDER BY NoItem"
    End With

    'SSTab1.TabEnabled = False
    Set mCall = New frmCaller
    LoadTreeBOM
    TreeBOMStructure
Exit Sub
1:
MessageBox Err.Description, "frmbom:form_load" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If MyDDE.CheckRecordPendinged = True Then
'   ScanKey vbKeyF5, 0, MyDDE
'   If MyDDE.IsSucces = True Then
'      Cancel = False
'      MyDDE.ClearRecordset
'      Set Frmbom = Nothing
'   Else
'      Cancel = True
'   End If
'Else
Set RcPart = Nothing
Set RcComponent = Nothing
Set RcPartner = Nothing
Set mCall = Nothing
Set Rc = Nothing
MyDDE.ClearRecordset
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmbom = Nothing
End Sub

Private Sub mCall_BeforeUnload()
On Error GoTo 1
Select Case mCall.FromTagActive
       Case "MASTER BOM":
'            If SSTab1.Tab = 0 Then
'                If Not IsNull(MyDDE.GetFieldByName("Bom ID")) = True Then
'                   If MyDDE.GetFieldByName("Bom ID") = "" Then MyDDE.CallButtonActive tmbCancel
'                End If
'                'mFirstCaller = False
'                If DataGrid1(0).Enabled = True Then DataGrid1(0).SetFocus
'            End If
                If FindOwnRecordset(MyDDE.ActiveRecordset, "[BOM Id] = '" & MyDDE.GetFieldByName("BOM Id") & "'") = True Then
                   MessageBox "Record -> " & MyDDE.GetFieldByName("BOM Id") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
                   MyDDE.CallButtonActive (tmbCancel)
'                   If MyDDE.ActiveRecordset.Recordcount <> 0 Then MyDDE.ActiveRecordset.MoveLast
                Else
                    If IsNull(MyDDE.GetFieldByName("BOM Id")) = True Or MyDDE.GetFieldByName("BOM Id") = "" Then
                       MyDDE.CallButtonActive (tmbCancel)
                    Else
                       MyDDE.GetFieldByName("BomReff") = IndexAuto
                    End If
                End If
       Case "MASTER STAGE":
            If SSTab1.Tab = 0 Then
                If FindOwnRecordset(MyDDE.ChildRecordset, "SeqStageID = '" & MyDDE.ChildRecordset.Fields("SeqStageID") & "'") = True Then
                   MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("SeqStageID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
                   MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                Else
                    If IsNull(MyDDE.ChildRecordset.Fields("SeqStageID")) = True Or MyDDE.ChildRecordset.Fields("SeqStageID") = "" Then
                       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                    End If
                End If
            Else
                If FindOwnRecordset(MyDDE.ChildRecordset, "SeqStageID = '" & MyDDE.ChildRecordset.Fields("SeqStageID") & "' and [Komponen ID] ='" & IIf(Not IsNull(MyDDE.ChildRecordset.Fields("Komponen ID")), MyDDE.ChildRecordset.Fields("Komponen ID"), "xxx") & "'") = True Then
                   MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("SeqStageID") & " dan " & MyDDE.ChildRecordset.Fields("Komponen ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
                   MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                Else
                    If IsNull(MyDDE.ChildRecordset.Fields("SeqStageID")) = True Or MyDDE.ChildRecordset.Fields("SeqStageID") = "" Then
                       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                    End If
                End If
                'mFirstCaller = False
            End If
            If DataGrid1(0).Enabled = True Then DataGrid1(0).SetFocus
       Case "MASTER BARANG":
                If FindOwnRecordset(MyDDE.ChildRecordset, "SeqStageID = '" & MyDDE.ChildRecordset.Fields("SeqStageID") & "' And [Komponen ID] = '" & MyDDE.ChildRecordset.Fields("Komponen ID") & "'") = True Then
                   MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("SeqStageID") & " Dan " & MyDDE.ChildRecordset.Fields("Komponen ID") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
                   MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                Else
                    If IsNull(MyDDE.ChildRecordset.Fields("SeqStageID")) = True Or MyDDE.ChildRecordset.Fields("SeqStageID") = "" Then
                       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                    End If
                End If
                'mFirstCaller = False
                'If DataGrid1(1).Enabled = True Then DataGrid1(1).SetFocus
       Case "BOM Referense":
        
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbom:mcall_beforeunload" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 2
Select Case mCall.FromTagActive
       Case "MASTER STAGE":
            If SSTab1.Tab = 0 Then
                With MyDDE.ChildRecordset
                     'messagebox .Source
                     .Fields("Noline") = .Recordcount
                     .Fields("Keterangan") = ""
                     .Fields("SeqStageID") = mCall.GetFieldByName(0)
                     .Fields("Keterangan") = mCall.GetFieldByName(1)
                     .Fields("Catatan") = "-"
                End With
            ElseIf SSTab1.Tab = 1 Then
                With MyDDE.ChildRecordset
                     .Fields("SeqStageID") = mCall.GetFieldByName(0)
                     '.Fields("NoLine") = mCall.GetFieldByName(0)
                End With
            End If
       Case "MASTER BARANG":
            With RcComponent.DBRecordset
                 .Fields("Komponen ID") = mCall.GetFieldByName(0)
                 .Fields("Nama Komponen") = mCall.GetFieldByName(1)
                 .Fields("UOM") = mCall.GetFieldByName(3)
                 .Fields("QTYUsage") = 1
                 .Fields("Keterangan") = "-"
            End With
       Case "MASTER RESOURCES":
            With MyDDE.ChildRecordset
                 .Fields("ResourcesID") = mCall.GetFieldByName(0)
                 .Fields("Resources") = "-"
            End With
       Case "MASTER BOM":
            With MyDDE
                 .GetFieldByName("BOM Id") = mCall.GetFieldByName(0)
                 .GetFieldByName("Keterangan") = mCall.GetFieldByName(1)
                 .GetFieldByName("BomReff") = IndexAuto
            End With
       Case "BOM Referense":
            If mVarCallTab = False Then MyDDE.GetFieldByName("BomReff") = mCall.GetFieldByName(0)
            OpenDetailComponent txtBox(0)
            OpenDetail txtBox(0)
            If SSTab1.Tab = 1 Then
               OpenDetailComponent txtBox(0)
            ElseIf SSTab1.Tab = 0 Then
               OpenDetail txtBox(0)
            End If
       Case "Work Center":
            With MyDDE.ChildRecordset
                 .Fields("WCID") = mCall.GetFieldByName(0)
                 '.Fields("Resources") = "-"
            End With
End Select
Exit Sub
2:
MessageBox Err.Description, "frmbom:mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim I As Integer
Dim j As Integer
Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
            txtBox(0).Enabled = False
            txtBox(2).Enabled = False
            DataCombo1.SetFocus
            Check1.Value = 0
            SSTab1.Tab = 0
            MyDDE.GetFieldByName("uom") = "PCS"
            
            OpenDetailPartner 4
       Case tmbEdit:
            mAdd = True
            txtBox(0).Enabled = False
            txtBox(2).Enabled = False
            DataCombo1.SetFocus
            Call SSTab1_Click(SSTab1.Tab)
       Case tmbCancel:
            mAdd = DataCombo1.Enabled
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
              ' 'mAdd = True
            Else
              ' mAdd = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               mvarDelete = False
                If MyDDE.ChildRecordset.Recordcount <> 0 Then
                   If SSTab1.Tab = 0 Then
                        With Rc.DBRecordset
                             If .Recordcount <> 0 Then
                                I = .AbsolutePosition
                                .MoveFirst
                                j = 0
                                Do
                                  j = j + 1
                                  If .EOF Then Exit Do
                                  .Fields("NoLine") = j
                                  .MoveNext
                                Loop
                                .AbsolutePosition = I
                             End If
                        End With
                   End If
                End If
            End If
       Case tmbDetail:
            
            If MyDDE.IsChildMemberReady = True Then
               If SSTab1.Tab = 0 Then
                  OpenDetailPartner 0
               Else
                  OpenDetailPartner 6
               End If
            End If
            
       Case tmbSave:
         On Error GoTo xErr
               'mAdd = True
                If MyDDE.IsChildMemberReady = True Then
                   CreateInventoryBatch 0
                   CreateInventoryBatch 1
                   If MyDDE.ActiveRecordset.Recordcount <> 0 Then
                      If SSTab1.Tab = 1 Then
                         OpenDetailComponent MyDDE.GetFieldByName("BOM Id")
                      Else
                         OpenDetail MyDDE.GetFieldByName("BOM Id")
                      End If
                   End If
                   Call SSTab1_Click(SSTab1.Tab)
                   mAdd = False
                 End If
       Case tmbPrint:
            CallRPTReport "BOM Report.rpt", "Select * from [BOM Report] Where [BOM ID] ='" & txtBox(0) & "'"
       Case Else: 'mVarDataDc = False
End Select
'mAdd = DataCombo1.Enabled
cmdLink(0).Enabled = DataCombo1.Enabled
MoveCtrl
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
OpenDetailComponent MyDDE.GetFieldByName("BOM Id")
OpenDetail MyDDE.GetFieldByName("BOM Id")
LoadTreeBOM

'SSTab1.Tab = 0
If SSTab1.Tab = 2 Then
   mKeyNode = 0
   CreateTree
'   OpenListData MyDDE.GetFieldByName("BOM Id")
End If
'Call SSTab1_Click(SSTab1.Tab)
mAdd = False
TreeBOMStructure
Exit Sub
1:
MessageBox Err.Description, "frmbom:mydde_movecomplete" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Dim I As Long
Dim j As Long
On Error GoTo xErr
Select Case AdReasonActiveDb
       Case tmbEdit:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.CancelTrans = CekHapus(0)
               If MyDDE.CancelTrans = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If CekHapus(SSTab1.Tab + 1) = False Then
                  'CopyRecordsetByQuery "Select * from Inventory Where Noitem='" & txtBox(0) & "'", "INVENTORY"
                  MyDDE.IsChildMemberReady = True
                  If SSTab1.Tab = 0 Then

                        With RcComponent.DBRecordset
                             If .Recordcount <> 0 Then
                                .Filter = "WCID ='" & MyDDE.ChildRecordset.Fields("SeqStageID") & "'"
                                If .Recordcount <> 0 Then
                                    .MoveFirst
                                        Do
                                          If .EOF Then Exit Do
                                          .Delete adAffectCurrent
                                          If Not .EOF Then .MoveNext
                                        Loop
                                    If .Recordcount <> 0 Or Not .EOF Then .MoveFirst
                                    .Filter = adFilterNone
                                    
                                End If
                             End If
                        End With
                  End If
                  If SSTab1.Tab = 0 Then
                     SendDataToServer ("Delete from [BOM Stage Detail] where (NoItem =N'" & txtBox(0) & "') and (BomReff=N'" & MyDDE.GetFieldByName("BomReff") & "')")
                  Else
                     SendDataToServer ("Delete from [BOM Component Detail] where (NoItem =N'" & txtBox(0) & "') and (BomReff=N'" & MyDDE.GetFieldByName("BomReff") & "')")
                  End If
                  PrepareQuery
                  mvarDelete = True
               Else
                  MyDDE.CancelTrans = True
                  'MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
               mvarDelete = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If Rc.Recordcount <> 0 Or RcComponent.Recordcount <> 0 Then
' Enda
'                  If CekDatakosong(0) = False And CekDatakosong(1) = False Then
                     MyDDE.IsChildMemberReady = True
                     PrepareQuery
'                  Else
'                     MyDDE.IsChildMemberReady = False
'                     MyDDE.CancelTrans = True
'                  End If
' end of Enda
               End If
            Else
               MessageBox "Detail transaksi BOM masih ada yang kosong. Silahkan dicek kembali", "Peringatan", msgOkOnly
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDetail:
'            If SSTab1.Tab = 0 Then
                MyDDE.CancelTrans = mFirstCaller
                If MyDDE.CancelTrans = True Then Exit Sub
                If MyDDE.ChildRecordset.Recordcount <> 0 Then
'                   If MyDDE.ChildRecordset.Fields(4) = 0 Then
'                      MyDDE.IsChildMemberReady = False
'                      MyDDE.CancelTrans = True
'                      MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly
'                   Else
                      MyDDE.IsChildMemberReady = True
                      MyDDE.CancelTrans = False
'                   End If
                Else
                   MyDDE.IsChildMemberReady = True
                   MyDDE.CancelTrans = False
                End If
'            Else
'               mAdd = True
'               MyDDE.CancelTrans = True
'               'RcComponent.DBRecordset.AddNew
'            End If
End Select
Set mDel = Nothing
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub SemeruTree1_NodeClick(ByVal Node As MSComctlLib.INode)
On Error GoTo 1
   If RcComponent.DBRecordset.Recordcount > 0 Then
      RcComponent.DBRecordset.MoveFirst
      RcComponent.DBRecordset.Find "[Komponen ID]='" & Node.Key & "'", , adSearchForward, 1
      If Not RcComponent.DBRecordset.EOF Then
            txt(0).Text = IIf(IsNull(RcComponent.DBRecordset.Fields("komponen ID")), "", RcComponent.DBRecordset.Fields("komponen ID"))
            txt(1).Text = IIf(IsNull(RcComponent.DBRecordset.Fields("NAma Komponen")), "", RcComponent.DBRecordset.Fields("Nama KOmponen"))
            txt(2).Text = IIf(IsNull(RcComponent.DBRecordset.Fields("Kategori")), "", RcComponent.DBRecordset.Fields("kategori"))
            txt(3).Text = IIf(IsNull(RcComponent.DBRecordset.Fields("kelompok")), "", RcComponent.DBRecordset.Fields("kelompok"))
            txt(4).Text = IIf(IsNull(RcComponent.DBRecordset.Fields("QtyUsage")), "", RcComponent.DBRecordset.Fields("QtyUsage"))
            txt(5).Text = IIf(IsNull(RcComponent.DBRecordset.Fields("LeadTime")), "", RcComponent.DBRecordset.Fields("LeadTime"))
            lblUOM.Caption = IIf(IsNull(RcComponent.DBRecordset.Fields("UOM")), "", RcComponent.DBRecordset.Fields("UOM"))
      End If
   End If
Exit Sub
1:
MessageBox Err.Description, "frmbom:semerutree1_nodeclick" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo 1
Select Case SSTab1.Tab
       Case 0:
            Set MyDDE.ChildRecordset = Rc.DBRecordset
'            Debug.Print Rc.DBRecordset.Source
            Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
'            Set txtBox(3).DataSource = MyDDE.ChildRecordset
       Case 1:
'            CreateInventoryBatch 0
            Set MyDDE.ChildRecordset = RcComponent.DBRecordset
            Set DataGrid1(1).DataSource = MyDDE.ChildRecordset
'            Set txtBox(4).DataSource = MyDDE.ChildRecordset
       Case 2:
            mKeyNode = 0
            CreateTree
'            OpenListData MyDDE.GetFieldByName("BOM Id")
'        Case 3:

End Select
Exit Sub
1:
MessageBox Err.Description, "frmbom:sstab1_click" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    '.PrepareAppend = " INSERT INTO [Inventory]" & _
                     " (NoItem, ItemName, UOM, MethodeID, Phantom,Manufacture)" & _
                     " VALUES  (N'" & txtBox(0) & "', N'" & txtBox(2) & "', N'" & txtBox(1) & "', N'" & DataCombo1.BoundText & "', " & Check1.Value & ",1)"
                     
    .PrepareAppend = " UPDATE [Inventory] Set Bomreff=N'" & Text1 & "', Status = N'" & Combo1 & "',[ItemName] = N'" & txtBox(2) & "',MethodeID=N'" & DataCombo1.BoundText & "',UOM=N'" & txtBox(1) & "',Phantom=" & Check1.Value & ",Manufacture =1  WHERE     ([NoItem] = N'" & txtBox(0) & "')"
    .PrepareUpdate = " UPDATE [Inventory] Set Bomreff=N'" & Text1 & "',Status = N'" & Combo1 & "',[ItemName] = N'" & txtBox(2) & "',MethodeID=N'" & DataCombo1.BoundText & "',UOM=N'" & txtBox(1) & "',Phantom=" & Check1.Value & "  WHERE     ([NoItem] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " UPDATE [Inventory] Set Manufacture=0 WHERE ([NoItem] = N'" & txtBox(0) & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub OpenDetail(ByVal Param As String)
'Rc.DBOpen "SELECT     [BOM Stage Detail].NoLine, [BOM Stage Detail].WCID As SeqStageID, [BOM Stage Detail].Description AS Keterangan, [BOM Stage Detail].ResourcesID,  [Resources Table].Description AS Resources, [BOM Stage Detail].StageNote AS Catatan FROM         [BOM Stage Detail] LEFT OUTER JOIN [Resources Table] ON [BOM Stage Detail].ResourcesID = [Resources Table].ResourcesID WHERE     ([BOM Stage Detail].NoItem = N'" & Param & "') AND ([BOM Stage Detail].BomReff = N'" & MyDDE.GetFieldByName("BomReff") & "') ORDER BY [BOM Stage Detail].NoLine", Cnn, lckLockBatch
Rc.DBOpen "SELECT [BOM Stage Detail].NoLine, [BOM Stage Detail].WCID AS SeqStageID, wcenter_header.Description AS Keterangan,wcenter_header.cycle_time / 60 as unit_run,wcenter_header.queue_time,wcenter_header.setup_time,wcenter_header.wait_time,wcenter_header.queue_time + wcenter_header.setup_time + wcenter_header.wait_time as total_run_time, [BOM Stage Detail].ResourcesID, [Resources Table].Description AS Resources, [BOM Stage Detail].StageNote AS Catatan FROM         [BOM Stage Detail] INNER JOIN wcenter_header ON [BOM Stage Detail].WCID = wcenter_header.WCID LEFT OUTER JOIN [Resources Table] ON [BOM Stage Detail].ResourcesID = [Resources Table].ResourcesID WHERE     ([BOM Stage Detail].NoItem = N'" & Param & "') AND ([BOM Stage Detail].BomReff = N'" & MyDDE.GetFieldByName("BomReff") & "') ORDER BY [BOM Stage Detail].NoLine", CNN
'Debug.Print Rc.DBRecordset.Source
Set MyDDE.ChildRecordset = Rc.DBRecordset '.Clone(adLockBatchOptimistic)
Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
Set DataGrid1(2).DataSource = MyDDE.ChildRecordset
'Set txtBox(3).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub OpenDetailBOMStructure(ByVal Param As String)
On Error GoTo 7
  RcBOMStruc.DBOpen "SELECT [Ord Comp Detail].StageID,[Ord Comp Detail].NoItem,Inventory.InternalName,inventory_categories.description AS class, [Inventory Group].[Group Name],[Ord Comp Detail].[Quote Qty],Inventory.UOM,Inventory.LeadTimeDays From [Ord Comp Detail] INNER JOIN Inventory ON ([Ord Comp Detail].NoItem = Inventory.NoItem) " & _
        " LEFT OUTER JOIN inventory_categories ON (Inventory.categid = inventory_categories.categid) LEFT OUTER JOIN [Inventory Group] ON (Inventory.NoGroup = [Inventory Group].NoGroup) Where [Ord Comp Detail].OrderID = '" & MyDDE.GetFieldByName("BomReff") & "' Order By [Ord Comp Detail].StageID ", CNN, lckLockBatch
Debug.Print "SELECT [Ord Comp Detail].StageID,[Ord Comp Detail].NoItem,Inventory.InternalName,inventory_categories.description AS class, [Inventory Group].[Group Name],[Ord Comp Detail].[Quote Qty],Inventory.UOM,Inventory.LeadTimeDays From [Ord Comp Detail] INNER JOIN Inventory ON ([Ord Comp Detail].NoItem = Inventory.NoItem) " & _
        " LEFT OUTER JOIN inventory_categories ON (Inventory.categid = inventory_categories.categid) LEFT OUTER JOIN [Inventory Group] ON (Inventory.NoGroup = [Inventory Group].NoGroup) Where [Ord Comp Detail].OrderID = '" & MyDDE.GetFieldByName("BomReff") & "' Order By [Ord Comp Detail].StageID "
        
Set MyDDE.ChildRecordset = RcBOMStruc.DBRecordset '.Clone(adLockBatchOptimistic)
Set TGBom.DataSource = MyDDE.ChildRecordset
Exit Sub
7:
MessageBox Err.Description, "frmbom:opendetailbomstructure" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDetailComponent(ByVal Param As String)
'  RcComponent.DBOpen " SELECT [BOM Component Detail].WCID, [BOM Component Detail].SeqStageID AS SeqStageID, [BOM Component Detail].Description AS Keterangan,   [BOM Component Detail].Component AS [Komponen ID], Inventory.ItemName AS [Nama Komponen], [BOM Component Detail].UOM, " & _
'                      " [BOM Component Detail].QTYUsage as QtyUsage FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem INNER JOIN  [BOM Stage Detail] ON [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND  [BOM Component Detail].BomReff = [BOM Stage Detail].BomReff AND [BOM Component Detail].WCID = [BOM Stage Detail].WCID WHERE     ([BOM Component Detail].NoItem = N'" & Param & "') AND ([BOM Component Detail].BomReff = N'" & MyDDE.GetFieldByName("BomReff") & "') ORDER BY [BOM Stage Detail].NoLine", CNN, lckLockBatch
                      
   RcComponent.DBOpen "SELECT [BOM Component Detail].WCID, [BOM Component Detail].SeqStageID, " & _
                           "[BOM Component Detail].Description AS Keterangan, [BOM Component Detail].Component AS [Komponen ID], " & _
                           "Inventory.internalName AS [Nama Komponen], [BOM Component Detail].UOM, [BOM Component Detail].QtyUsage, " & _
                           "inventory_categories.description AS kategori, [Inventory Group].[Group Name] AS kelompok, Inventory.LeadTimeDays AS leadTime " & _
                      "FROM [BOM Component Detail] INNER JOIN " & _
                           "Inventory ON [BOM Component Detail].Component = Inventory.NoItem INNER JOIN " & _
                           "[BOM Stage Detail] ON [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND " & _
                           "[BOM Component Detail].BomReff = [BOM Stage Detail].BomReff AND " & _
                           "[BOM Component Detail].WCID = [BOM Stage Detail].WCID INNER JOIN " & _
                           "inventory_categories ON Inventory.categid = inventory_categories.categid AND " & _
                           "Inventory.NoGroup = inventory_categories.nogroup INNER JOIN " & _
                           "[Inventory Group] ON inventory_categories.nogroup = [Inventory Group].NoGroup " & _
                      "WHERE  ([BOM Component Detail].NoItem = N'" & Param & "') AND ([BOM Component Detail].BomReff = N'" & MyDDE.GetFieldByName("BomReff") & "') " & _
                      "ORDER BY [BOM Stage Detail].NoLine ", CNN, lckLockBatch
                      
   Set MyDDE.ChildRecordset = RcComponent.DBRecordset '.Clone(adLockBatchOptimistic)
   Set DataGrid1(1).DataSource = MyDDE.ChildRecordset

   LoadBillOfMAnufacture
End Sub

Private Sub LoadBillOfMAnufacture()
Dim sWC As String
   SemeruTree1.MenuTreeView.Nodes.Clear
   If RcComponent.DBRecordset.Recordcount > 0 Then
        With SemeruTree1
            
            Set .MenuTreeView.ImageList = MainMenu.ImageList1
            .BackColorTree = &H6D4016
            .NodeAdd , tvwChild, "Master", MyDDE.GetFieldByName("Keterangan"), "Master", , , True, , , True, , &HFCF1ED, &H6D4016
            RcComponent.DBRecordset.MoveFirst
            sWC = ""
            While Not RcComponent.DBRecordset.EOF
               If sWC <> RcComponent.DBRecordset.Fields("WCID") Then
                  sWC = RcComponent.DBRecordset.Fields("WCID")
                  .NodeAdd "Master", tvwChild, sWC, sWC, "biru", , , , , , True, , &HFCF1ED, &H6D4016
                  .NodeAdd sWC, tvwChild, RcComponent.DBRecordset.Fields("komponen ID"), RcComponent.DBRecordset.Fields("komponen ID"), "ijo", , , True, , , True, False, &HFCF1ED, &H6D4016
               Else
                  .NodeAdd sWC, tvwChild, RcComponent.DBRecordset.Fields("komponen ID"), RcComponent.DBRecordset.Fields("komponen ID"), "ijo", , , , , , True, , &HFCF1ED, &H6D4016
               End If
               RcComponent.DBRecordset.MoveNext
            Wend
        End With
    End If
End Sub

Private Sub OpenPartner()
RcPart.DBOpen "SELECT MethodeID, Description FROM [BOM Methode] ORDER BY Description", CNN, lckLockBatch
DataCombo1.ListField = "Description"
Set DataCombo1.RowSource = RcPart.DBRecordset
End Sub

Private Function IndexAuto() As String
On Error GoTo 4
Dim Rc As New DBQuick
Dim Inom As Long
Rc.DBOpen "SELECT     MAX(RIGHT(BomReff, 5)) AS MaxNom FROM         [BOM Stage Detail] WHERE     (NoItem = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "#BOM" & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "#BOM" & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "#BOM" & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "#BOM" & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "#BOM" & "-" & "0" & Trim(Str(Inom))
     End Select
End With
Exit Function
4:
MessageBox Err.Description, "frmbom:indexauto" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub OpenDetailPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            If SSTab1.Tab = 0 Then
               RcPartner.DBOpen "SELECT     WCID, Description AS Keterangan FROM         wcenter_header ORDER BY WCID", CNN, lckLockReadOnly
            ElseIf SSTab1.Tab = 1 Then
               'RcPartner.DBOpen "SELECT NoLine AS [No], SeqStageID AS [Stage ID], Description AS Keterangan FROM         [BOM Stage Detail] WHERE     (NoItem = N'" & MyDDE.GetFieldByName("Bom Id") & "') AND (BomReff=N'" & MyDDE.GetFieldByName("BomReff") & "')ORDER BY NoLine", Cnn, lckLockReadOnly
               RcPartner.DBOpen "SELECT [WC Stage].StageID AS [Stage ID], [Manufacture Stage].Description AS Keterangan FROM [WC Stage] INNER JOIN [Manufacture Stage] ON [WC Stage].StageID = [Manufacture Stage].StageID WHERE     ([WC Stage].WCID = N'" & MyDDE.ChildRecordset.Fields("WCID") & "') ORDER BY [WC Stage].[no]", CNN, lckLockReadOnly
               'messagebox RcPartner.DBRecordset.Source
            End If
       Case 2: RcPartner.DBOpen "SELECT Inventory.NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.Merk, Inventory.UOM FROM         Inventory LEFT OUTER JOIN  [BOM Component Detail] ON Inventory.NoItem = [BOM Component Detail].Component GROUP BY Inventory.NoItem, Inventory.ItemName, Inventory.Merk, Inventory.UOM HAVING      (Inventory.NoItem <> N'" & txtBox(0) & "') ORDER BY Inventory.NoItem", CNN, lckLockReadOnly
       Case 3: RcPartner.DBOpen "SELECT ResourcesID AS Resources, Description AS Keterangan, TypeID AS [Tipe Resources], Note AS Catatan FROM         [Resources Table] ORDER BY ResourcesID", CNN, lckLockReadOnly
       Case 4: RcPartner.DBOpen "SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], Merk, UOM FROM         Inventory order by NoItem ", CNN, lckLockReadOnly
       Case 5: RcPartner.DBOpen "SELECT [BOM Stage Detail].BomReff, [BOM Stage Detail].NoItem, Inventory.ItemName FROM         [BOM Stage Detail] INNER JOIN                       Inventory ON [BOM Stage Detail].NoItem = Inventory.NoItem GROUP BY [BOM Stage Detail].NoItem, [BOM Stage Detail].BomReff, Inventory.ItemName HAVING      ([BOM Stage Detail].NoItem = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
       Case 6: RcPartner.DBOpen "SELECT wcenter_header.WCID, wcenter_header.Description FROM  wcenter_header INNER JOIN [BOM Stage Detail] ON wcenter_header.WCID = [BOM Stage Detail].WCID GROUP BY wcenter_header.WCID, wcenter_header.Description, [BOM Stage Detail].NoItem HAVING      ([BOM Stage Detail].NoItem = N'" & txtBox(0) & "') ORDER BY [BOM Stage Detail].NoItem", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
               mCall.FromTagActive = "MASTER STAGE"
          Case 2:
               mCall.FromTagActive = "MASTER BARANG"
          Case 3:
               mCall.FromTagActive = "MASTER RESOURCES"
          Case 4:
               mCall.FromTagActive = "MASTER BOM"
          Case 5:
               mCall.FromTagActive = "BOM Referense"
          Case 6:
               mCall.FromTagActive = "Work Center"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
    MessageBox "Data Belum Ada. Silahkan dicek kembali", "Peringatan", msgOkOnly
    If IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields(0) = "" Then
        MyDDE.ChildRecordset.CancelBatch adAffectCurrent
        If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
    End If
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub MoveCtrl()
On Error GoTo 6
If mAdd = False Then Exit Sub
If SSTab1.Tab = 0 Then
   If MyDDE.ChildRecordset.Recordcount = 0 And mAdd = True Then Exit Sub
ElseIf SSTab1.Tab = 1 Then
   If RcComponent.DBRecordset.Recordcount = 0 And mAdd = True Then Exit Sub
End If
Select Case DataGrid1(SSTab1.Tab).Index
       Case 0:
            Select Case DataGrid1(SSTab1.Tab).col
                   Case 1, 3:
                       DataGrid1(SSTab1.Tab).Columns(DataGrid1(SSTab1.Tab).col).Button = True
                       DataGrid1(SSTab1.Tab).AllowUpdate = False
                   Case Else:
                       DataGrid1(SSTab1.Tab).Columns(DataGrid1(SSTab1.Tab).col).Button = False
            End Select
       Case 1:
            Select Case DataGrid1(SSTab1.Tab).col
                   Case 0, 1, 2:
                       DataGrid1(SSTab1.Tab).Columns(DataGrid1(SSTab1.Tab).col).Button = True
                       DataGrid1(SSTab1.Tab).AllowUpdate = False
                   Case Else:
                       DataGrid1(SSTab1.Tab).Columns(DataGrid1(SSTab1.Tab).col).Button = False
            End Select
End Select
Err.Clear
Exit Sub
6:
MessageBox Err.Description, "frmbom:movecontrol" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub CreateInventoryBatch(ByVal TabIndexNo As Integer)
On Error GoTo 2
Dim z As Integer
If mAdd = True Then
    If MyDDE.IsChildMemberReady = True Then
        If TabIndexNo = 0 Then
            If Rc.DBRecordset.Recordcount <> 0 Then
               With Rc.DBRecordset
                    .MoveFirst
                    If SendDataToServer("Delete From [BOM Stage Detail] WHERE  (NoItem = N'" & txtBox(0) & "') and (BomReff= N'" & Text1 & "')") = True Then
                       Do
                       If Rc.DBRecordset.EOF Then Exit Do
                          If IsNull(.Fields("ResourcesID")) = False Then
                             SendDataToServer " INSERT INTO [BOM Stage Detail]" & _
                                              " (BomReff,WCID, Description, NoItem, ResourcesID, StageNote, NoLine)" & _
                                              " VALUES (N'" & Text1 & "',N'" & .Fields("SeqStageID") & "', N'" & .Fields("Keterangan") & "', N'" & txtBox(0) & "', N'" & .Fields("ResourcesID") & "', N'" & .Fields("Catatan") & "', " & .Fields("NoLine") & ")"
                          Else
                             SendDataToServer " INSERT INTO [BOM Stage Detail]" & _
                                              " (BomReff,WCID, Description, NoItem, StageNote, NoLine)" & _
                                              " VALUES (N'" & Text1 & "',N'" & .Fields("SeqStageID") & "', N'" & .Fields("Keterangan") & "', N'" & txtBox(0) & "', N'" & .Fields("Catatan") & "', " & .Fields("NoLine") & ")"
                          End If
                          .MoveNext
                       Loop
                    End If
                    .MoveLast
               End With
            End If
        ElseIf TabIndexNo = 1 Then
            If RcComponent.DBRecordset.Recordcount <> 0 Then
               With RcComponent.DBRecordset
                    .MoveFirst
                    If SendDataToServer("Delete From [BOM Component Detail] WHERE  (NoItem = N'" & txtBox(0) & "') and (BomReff= N'" & Text1 & "')") = True Then
                       Do
                       If .EOF Then Exit Do
                          If Not IsNull(.Fields("SeqStageID")) Then
                          SendDataToServer " INSERT INTO [BOM Component Detail]" & _
                                           " (BomReff,WCID,SeqStageID, NoItem, Component, UOM, QTYUsage,  Description)" & _
                                           " VALUES  (N'" & Text1 & "',N'" & .Fields("WCID") & "',N'" & .Fields("SeqStageID") & "', N'" & txtBox(0) & "', N'" & .Fields("Komponen ID") & "', N'" & .Fields("UOM") & "', " & CDbl(.Fields("QTYUsage")) & ",  N'" & .Fields("Keterangan") & "')"
                          End If
                          .MoveNext
                       Loop
                    End If
                    .MoveLast
               End With
            End If
        End If
    End If
End If
Exit Sub
2:
MessageBox Err.Description, "frmbom:create inventorybatch" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Function CekDatakosong(PreviousTab As Integer) As Boolean
On Error GoTo 1
Dim Avdata As Variant
Dim RcCek As New Recordset
Dim I As Integer
Dim j As Integer
Dim Fld As Field
If mAdd = False Then Exit Function
Select Case PreviousTab + 1
       Case 0: CekDatakosong = MyDDE.CheckEmptyControl
       Case 1:
            If Rc.DBRecordset.Recordcount <> 0 Then
               Set RcCek = Rc.DBRecordset.Clone(adLockReadOnly)
               Avdata = RcCek.Getrows(RcCek.Recordcount, adBookmarkFirst)
               For I = 0 To UBound(Avdata, 2)
                   j = 0
                   For Each Fld In RcCek.Fields
                       CekDatakosong = IsNull(Avdata(j, I))
                       
                       
                       If CekDatakosong = True Then
                          If j = 3 Or j = 4 Then
                             CekDatakosong = False
                          Else
                             CekDatakosong = True
                             GoTo Hell
'                             messagebox Fld.Name
                          End If
                       Else
                          CekDatakosong = False
                       End If
                       j = j + 1
                   Next
               Next I
            End If
       Case 2:
            If RcComponent.DBRecordset.Recordcount <> 0 Then
               Set RcCek = RcComponent.DBRecordset.Clone(adLockReadOnly)
               Avdata = RcCek.Getrows(RcCek.Recordcount, adBookmarkFirst)
               For I = 0 To UBound(Avdata, 2)
                   j = 0
                   For Each Fld In RcCek.Fields
                       CekDatakosong = IsNull(Avdata(j, I))
                       j = j + 1
                       If CekDatakosong = True Then GoTo Hell
                   Next
               Next I
            End If
End Select
Hell:
If CekDatakosong = True Then
   If PreviousTab <> 0 Then
      MessageBox "Data masih ada yang belum lengkap." & vbCrLf & "Silahkan dilengkapi dulu", "Peringatan", msgOkOnly
   Else
   End If
   SSTab1.Tab = PreviousTab
Else
  
End If
Set RcCek = Nothing
Exit Function
1:
MessageBox Err.Description, "frmbom:cekdatakosong" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function CekHapus(ByVal JenisHapus As Integer) As Boolean
On Error GoTo Hell
Dim RcHps As New DBQuick
Select Case JenisHapus
       Case 0:
            RcHps.DBOpen "SELECT     NoItem FROM         [Manufacture Order] WHERE     (NoItem = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
       Case 1:
            RcHps.DBOpen "SELECT     [Order Output Detail].WCID, [Manufacture Order].NoItem FROM         [Order Output Detail] INNER JOIN [Manufacture Order] ON [Order Output Detail].OrderID = [Manufacture Order].OrderID GROUP BY [Order Output Detail].WCID, [Manufacture Order].NoItem HAVING      ([Manufacture Order].NoItem = N'" & txtBox(0) & "') AND ([Order Output Detail].WCID = N'" & Rc.DBRecordset.Fields("SeqStageID") & "')", CNN, lckLockReadOnly
       Case 2:
            RcHps.DBOpen "SELECT     [Ord Comp Detail].NoItem FROM         [Ord Comp Detail] INNER JOIN [Manufacture Order] ON [Ord Comp Detail].OrderID = [Manufacture Order].OrderID WHERE     ([Manufacture Order].NoItem = N'" & RcComponent.DBRecordset.Fields("Komponen ID") & "') GROUP BY [Ord Comp Detail].NoItem HAVING      ([Ord Comp Detail].NoItem = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
End Select
If RcHps.DBRecordset.Recordcount <> 0 Then
   MessageBox "Data BOM sedang digunakan pada Production Order", "Peringatan", msgOkOnly
   CekHapus = True
End If
Hell:
RcHps.CloseDB
Set RcHps = Nothing
End Function

Private Sub GridLayout()
   DataGrid1(0).Columns(0).width = 810.1418
   DataGrid1(0).Columns(1).width = 1814.74
   DataGrid1(0).Columns(2).width = 3915.213
   DataGrid1(0).Columns(3).width = 2039.811
   DataGrid1(1).Columns(0).width = 1365.165
   DataGrid1(1).Columns(1).width = 1590.236
   DataGrid1(1).Columns(2).width = 1755.213
   DataGrid1(1).Columns(3).width = 2280.189
   DataGrid1(1).Columns(4).width = 824.882
   DataGrid1(1).Columns(5).width = 780.0945
End Sub

'Private Sub OpenListData(ByVal vKode As String)
'Dim Rc As New DBQuick
'Dim I As Integer
'Dim mLoad As Variant
'Dim mKey As String
'Dim mCaption As String
'Rc.DBOpen "SELECT  [BOM Component Detail].Component AS NoItem, Inventory.ItemName AS Description,Inventory.Manufacture FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY [BOM Component Detail].Component", Cnn, lckLockReadOnly
'With Rc.DBRecordset
'     If .Recordcount <> 0 Then
'        mLoad = .Getrows(.Recordcount, adBookmarkFirst)
'        For I = 0 To UBound(mLoad, 2)
'            mKeyNode = mKeyNode + 1
'            If TreeView1.Nodes.Count = 1 Then
'               With TreeView1.Nodes.Add("1~" & MyDDE.GetFieldByName("Bom ID"), tvwChild, mKeyNode & "~" & mLoad(0, I), TotalCaption(mLoad(0, I)) & " " & mLoad(1, I))
'                    .Expanded = True
'               End With
'               mKey = "1~" & MyDDE.GetFieldByName("Bom ID")
'               If OpenDetailStruc(MyDDE.GetFieldByName("Bom ID"), mKey) = "" Then
'                  mKey = mKey
'               Else
'                  mKey = mKeyNode & "~" & mLoad(0, I)
'               End If
'            Else
'               With TreeView1.Nodes.Add(mKey, tvwChild, mKeyNode & "~" & mLoad(0, I), TotalCaption(mLoad(0, I)) & " " & mLoad(1, I))
'                    .Expanded = True
'               End With
'               If OpenDetailStruc(mLoad(0, I), mKey) = "" Then
'                  mKey = mKey
'               Else
'                  mKey = mKeyNode & "~" & mLoad(0, I)
'               End If
'            End If
'        Next I
'     End If
'End With
'End Sub

'Private Sub OpenListData(ByVal vKode As String)
'On Error Resume Next
'Dim strSQL As String
'Dim Rc As New DBQuick
'Dim I As Integer
'Dim mLoad As Variant
'Dim mKey As String
'Dim mCaption As String
''Rc.DBOpen "SELECT  [BOM Component Detail].Component AS NoItem, Inventory.ItemName AS Description,Inventory.Manufacture,[BOM Component Detail].NoItem AS KeyNote FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY [BOM Component Detail].Component", Cnn, lckLockReadOnly
''Rc.DBOpen "SELECT     [BOM Component Detail].Component, Inventory.ItemName AS Description, Inventory.Manufacture, [BOM Component Detail].NoItem, ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY FROM         [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE     ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component", Cnn, lckLockReadOnly
'Rc.DBOpen "SELECT     [BOM Component Detail].Component, Inventory.ItemName AS Description, Inventory.Manufacture, [BOM Component Detail].NoItem, " & _
'            " ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS [QTY Available], [BOM Component Detail].QTYUsage " & _
'            " FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN " & _
'            " [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
'            " WHERE ([BOM Component Detail].NoItem = N'" & vKode & "') " & _
'            " ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component ", CNN, lckLockReadOnly
'
''strSQL = "SHAPE {SELECT [BOM Component Detail].Component, Inventory.ItemName AS Description, Inventory.Manufacture, [BOM Component Detail].NoItem, " & _
'            " ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS [QTY Available], [BOM Component Detail].QTYUsage " & _
'            " FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN " & _
'            " [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
'            " WHERE ([BOM Component Detail].NoItem = N'" & vKode & "') " & _
'            " ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component } AS HEADER " & _
'            " APPEND ({SELECT [BOM Component Detail].Component, Inventory.ItemName AS Description, Inventory.Manufacture, " & _
'            " [BOM Component Detail].NoItem, ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY " & _
'            " FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem " & _
'            " LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
'            " WHERE ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, " & _
'            " [BOM Component Detail].Component} AS DETIL RELATE NoItem TO NoItem)"
''Debug.Print strSQL
'With Rc.DBRecordset
''   Debug.Print .Source
'     If .Recordcount <> 0 Then
'        mLoad = .Getrows(.Recordcount, adBookmarkFirst)
''        FlexGrid.RowExpanded = True
'        For I = 0 To UBound(mLoad, 2)
''            Debug.Print Len(TotalCaption(mLoad(1, I)))
'            With TreeView1.Nodes.Add(vOldKey, tvwChild, mKeyNode + 1 & "~" & _
'                mLoad(0, I), TotalCaption(mLoad(0, I)) & " " & _
'                TotalCaption(mLoad(1, I)) & " " & mLoad(5, I) & Space(10) & mLoad(4, I))   'TotalCaptionR(mLoad(4, I), Len(mLoad(1, I))))
'                     .Expanded = True
'
''                FlexGrid.AddItem mLoad(0, I)
''                FlexGrid.BandData(0) = mLoad(0, I)
'            End With
'
''            strSQL = "SELECT [BOM Component Detail].Component, Inventory.ItemName AS Description, Inventory.Manufacture, " & _
'            " [BOM Component Detail].NoItem, ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY " & _
'            " FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem " & _
'            " LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
'            " WHERE ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, " & _
'            " [BOM Component Detail].Component"
'
'            If CBool(mLoad(2, I)) = True Then
'               OpenDetailStruc mLoad(0, I), mKeyNode + 1 & "~" & mLoad(0, I)
'            Else
'               mKeyNode = mKeyNode + 1
'            End If
'        Next I
'     End If
'End With
'Err.Clear
'End Sub

Private Sub CreateTree()
On Error GoTo 3
Dim strSQL As String
Dim rsFlex As New ADODB.Recordset


'ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS [QTY Available],
'strSQL = "SHAPE { SELECT     TOP 100 PERCENT [BOM Component Detail].Component, Inventory.ItemName AS Description, " & _
'" [BOM Component Detail].NoItem, [BOM Component Detail].QTYUsage AS [Usage QTY]" & _
'" FROM  [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN " & _
'" [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
'" WHERE ([BOM Component Detail].NoItem = N'" & MyDDE.GetFieldByName("BOM Id") & "') " & _
'" ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component } AS HEADER " & _
'" APPEND({SELECT     TOP 100 PERCENT [BOM Component Detail].Component, Inventory.ItemName AS Description, " & _
'" [BOM Component Detail].NoItem,  [BOM Component Detail].QTYUsage AS [Usage QTY] " & _
'" FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem " & _
'" LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
'" ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component } AS DETIL " & _
'" RELATE COMPONENT TO NOITEM)"


strSQL = "SHAPE { SELECT  TOP (100) PERCENT [BOM Component Detail].Component, Inventory.ItemName AS Description, [BOM Component Detail].NoItem, " & _
               " [BOM Component Detail].QtyUsage AS [Usage QTY] FROM [BOM Component Detail] INNER JOIN " & _
               " Inventory ON [BOM Component Detail].Component = Inventory.NoItem " & _
         " WHERE ([BOM Component Detail].NoItem = N'" & MyDDE.GetFieldByName("BOM Id") & "') " & _
         " ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component } AS HEADER " & _
         " APPEND({SELECT     TOP 100 PERCENT [BOM Component Detail].Component, Inventory.ItemName AS Description, " & _
         " [BOM Component Detail].NoItem,  [BOM Component Detail].QTYUsage AS [Usage QTY] " & _
         " FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem " & _
         " LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem " & _
         " ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component } AS DETIL " & _
         " RELATE COMPONENT TO NOITEM)"


rsFlex.Open strSQL, CNN, adOpenStatic, adLockReadOnly
Set TGBom.DataSource = rsFlex
Exit Sub
3:
MessageBox Err.Description, "frmbom:createtree" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub TreeBOMStructure()
On Error GoTo 9
    Dim sWCID As String
    
                   
    If MyDDE.ActiveRecordset.Recordcount > 0 Then

        With semeruBOM

            .MenuTreeView.Nodes.Clear
            Set .MenuTreeView.ImageList = MainMenu.ImageList1
            .BackColorTree = &H6D4016
            .NodeAdd , tvwChild, "Master", "Item Barang", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
            sWCID = ""
            If MyDDE.ActiveRecordset.EOF Or MyDDE.ActiveRecordset.BOF Then Exit Sub
                    .NodeAdd "Master", tvwChild, MyDDE.ActiveRecordset.Fields("bom id"), IIf(IsNull(MyDDE.ActiveRecordset.Fields("Keterangan")), "", MyDDE.ActiveRecordset.Fields("Keterangan")), "biru", , , , , , True, , &HFCF1ED, &H6D4016
               ' .NodeAdd sWCID, tvwChild, MyDDE.ActiveRecordset.Fields("bom id"), MyDDE.ActiveRecordset.Fields("bom id"), "ijo", , , , , , True, , &HFCF1ED, &H6D4016
               ' RsTeeBOM.DBRecordset.MoveNext
           ' Wend
        End With

    End If
Exit Sub
9:
MessageBox Err.Description, "frmbom:treebomstructure" & Err.Number, msgOkOnly, msgExclamation
End Sub

'Private Function OpenDetailStruc(ByVal vKode As String, ByVal vKey As String) As String
'Dim Rc As New DBQuick
'Dim I As Integer
'Dim mLoad As Variant
'Dim mNodStr As String
'Rc.DBOpen "SELECT [BOM Component Detail].Component, Inventory.ItemName AS Description, Inventory.Manufacture, [BOM Component Detail].NoItem, ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY FROM         [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE     ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY Inventory.Manufacture, [BOM Component Detail].NoItem, [BOM Component Detail].Component", CNN, lckLockReadOnly
'With Rc.DBRecordset
'     If .Recordcount <> 0 Then
'        mLoad = .Getrows(.Recordcount, adBookmarkFirst)
'        For I = 0 To UBound(mLoad, 2)
'            mKeyNode = mKeyNode + 1
'            With TreeView1.Nodes.Add(vKey, tvwChild, mKeyNode + 1 & "~" & mLoad(0, I), TotalCaption(mLoad(0, I)) & " " & TotalCaption(mLoad(1, I)) & " " & mLoad(4, I))
'                 .Expanded = True
'            End With
'            mNodStr = mLoad(0, I)
'            If CBool(mLoad(2, I)) = True Then OpenDetailStruc = OpenDetailStruc(mNodStr, mKeyNode + 1 & "~" & mLoad(0, I))
'        Next I
'        OpenDetailStruc = OpenDetailStruc(mNodStr, vKey)
'     Else
'        OpenDetailStruc = ""
'        Exit Function
'     End If
'End With
'End Function

Private Function TotalCaption(ByVal vTextCaption As String, Optional ByVal vLen As Integer) As String
If vLen = 0 Then vLen = 35
'TotalCaption = vTextCaption + Space(vLen - LenB(vTextCaption)) , vLen + Len(vTextCaption))
TotalCaption = Left(vTextCaption + Space(vLen - Len(vTextCaption)), vLen + Len(vTextCaption))

'Debug.Print TotalCaption & Len(Space(vLen - Len(vTextCaption))) & Len(TotalCaption)
'Debug.Print Len(TotalCaption)
End Function

Private Function TotalCaptionR(ByVal vTextCaption As String, ByVal vLenPrevText As Integer, Optional ByVal vLen As Integer) As String
On Error GoTo 8
If vLen = 0 Then
   vLen = 35
   vLen = vLen - vLenPrevText
End If
TotalCaptionR = Format(Space(vLen - Len(vTextCaption)) + vTextCaption, "#,#0")
Exit Function
8:
MessageBox Err.Description, "frmbom:totalcaptionr" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub LoadTreeBOM()
On Error GoTo 5
   Dim sWCID As String
   RsTeeBOM.DBOpen "SELECT [BOM Component Detail].SeqStageID, [BOM Component Detail].NoItem, Inventory.InternalName, " & _
                           "inventory_categories.description AS class,[Inventory Group].[Group Name] , " & _
                           "[BOM Component Detail].[QtyUsage], Inventory.UOM, Inventory.LeadTimeDays " & _
                   "FROM  [BOM Component Detail] INNER JOIN " & _
                      " Inventory ON [BOM Component Detail].NoItem = Inventory.NoItem LEFT OUTER JOIN " & _
                      " inventory_categories ON Inventory.categid = inventory_categories.categid LEFT OUTER JOIN" & _
                      " [Inventory Group] ON Inventory.NoGroup = [Inventory Group].NoGroup " & _
                   "WHERE [BOM Component Detail].component = '" & txtBox(0).Text & "' order by [BOM Component Detail].SeqStageID ", CNN
   If RsTeeBOM.DBRecordset.Recordcount > 0 Then
      With SemeruTree1
      
         txt(0).Text = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("noItem")), "", RsTeeBOM.DBRecordset.Fields("noItem"))
         txt(1).Text = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("InternalName")), "", RsTeeBOM.DBRecordset.Fields("InternalName"))
         txt(2).Text = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("class")), "", RsTeeBOM.DBRecordset.Fields("class"))
         txt(3).Text = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("Group Name")), "", RsTeeBOM.DBRecordset.Fields("Group Name"))
         txt(4).Text = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("QtyUsage")), "", RsTeeBOM.DBRecordset.Fields("QtyUsage"))
         txt(5).Text = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("LeadTimeDays")), "", RsTeeBOM.DBRecordset.Fields("LeadTimeDays"))
         lblUOM.Caption = IIf(IsNull(RsTeeBOM.DBRecordset.Fields("UOM")), "", RsTeeBOM.DBRecordset.Fields("UOM"))

         .MenuTreeView.Nodes.Clear
         Set .MenuTreeView.ImageList = MainMenu.ImageList1
         .BackColorTree = &H6D4016
         .NodeAdd , tvwChild, "Master", txtBox(0).Text, "Master", , , True, , , True, , &HFCF1ED, &H6D4016
         sWCID = ""
         While Not RsTeeBOM.DBRecordset.EOF
            If sWCID <> RsTeeBOM.DBRecordset.Fields("SeqStageID") Then
               sWCID = RsTeeBOM.DBRecordset.Fields("SeqStageID")
               .NodeAdd "Master", tvwChild, sWCID, sWCID, "biru", , , , , , True, , &HFCF1ED, &H6D4016
            End If
            .NodeAdd sWCID, tvwChild, RsTeeBOM.DBRecordset.Fields("noItem"), RsTeeBOM.DBRecordset.Fields("InternalName"), "ijo", , , , , , True, , &HFCF1ED, &H6D4016
            RsTeeBOM.DBRecordset.MoveNext
         Wend
      End With
   End If
Exit Sub
5:
MessageBox Err.Description, "frmbom:loadfreebom" & Err.Number, msgOkOnly, msgExclamation
End Sub


