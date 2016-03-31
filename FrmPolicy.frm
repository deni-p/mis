VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmPolicy 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otorisasi User"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPolicy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11970
   Tag             =   "User Authentication Setting"
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11970
      TabIndex        =   25
      Top             =   7365
      Width           =   11970
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Simpan"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   1500
         Picture         =   "FrmPolicy.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Hapus"
         Enabled         =   0   'False
         Height          =   555
         Index           =   4
         Left            =   2940
         Picture         =   "FrmPolicy.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Batal"
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   2220
         Picture         =   "FrmPolicy.frx":138F6
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Tambah"
         Height          =   555
         Index           =   0
         Left            =   60
         Picture         =   "FrmPolicy.frx":1A148
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Edit"
         Height          =   555
         Index           =   1
         Left            =   780
         Picture         =   "FrmPolicy.frx":2099A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   60
         Width           =   720
      End
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   -30
         TabIndex        =   45
         Top             =   0
         Width           =   12030
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   555
         Left            =   3660
         Picture         =   "FrmPolicy.frx":271EC
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7425
      Left            =   0
      ScaleHeight     =   7425
      ScaleWidth      =   11970
      TabIndex        =   20
      Top             =   0
      Width           =   11970
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   810
         Top             =   825
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   7159830
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":28CE6
               Key             =   "Orang"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":298BA
               Key             =   "person1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":2A196
               Key             =   "person2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":2AA72
               Key             =   "TOP"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":2B8C6
               Key             =   "Dept"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":2C392
               Key             =   "abang"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":32BF4
               Key             =   "biru"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPolicy.frx":39456
               Key             =   "ijo"
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7215
         Left            =   3480
         TabIndex        =   1
         Top             =   75
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   12726
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "List User"
         TabPicture(0)   =   "FrmPolicy.frx":3FCB8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Akses Form"
         TabPicture(1)   =   "FrmPolicy.frx":3FCD4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Group"
         TabPicture(2)   =   "FrmPolicy.frx":3FCF0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame6"
         Tab(2).Control(1)=   "Frame_group"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Detil Group"
         TabPicture(3)   =   "FrmPolicy.frx":3FD0C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame7"
         Tab(3).Control(1)=   "Frame3"
         Tab(3).ControlCount=   2
         Begin VB.Frame Frame7 
            Caption         =   "List Group Form"
            Height          =   3030
            Left            =   -74895
            TabIndex        =   59
            Top             =   4110
            Width           =   8175
            Begin VB.TextBox txtcari1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1020
               TabIndex        =   61
               Top             =   2625
               Width           =   3360
            End
            Begin MSDataGridLib.DataGrid DGGrid 
               Height          =   2280
               Left            =   120
               TabIndex        =   60
               Top             =   270
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   4022
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
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
               ColumnCount     =   8
               BeginProperty Column00 
                  DataField       =   "ID"
                  Caption         =   "ID"
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
                  DataField       =   "Group Name"
                  Caption         =   "Group Name"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Group Form"
                  Caption         =   "Group Form"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Form Name"
                  Caption         =   "Form Name"
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
                  DataField       =   "Access"
                  Caption         =   "Access"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "True"
                     FalseValue      =   "False"
                     NullValue       =   ""
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "New"
                  Caption         =   "New"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "True"
                     FalseValue      =   "False"
                     NullValue       =   ""
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "edit"
                  Caption         =   "Edit"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "True"
                     FalseValue      =   "False"
                     NullValue       =   ""
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "del"
                  Caption         =   "Delete"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "True"
                     FalseValue      =   "False"
                     NullValue       =   ""
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  AllowRowSizing  =   0   'False
                  BeginProperty Column00 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   2
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   2
                  EndProperty
                  BeginProperty Column06 
                     Alignment       =   2
                  EndProperty
                  BeginProperty Column07 
                     Alignment       =   2
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Cari Kriteria"
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   62
               Top             =   2670
               Width           =   840
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "List Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5535
            Left            =   -74880
            TabIndex        =   57
            Top             =   1560
            Width           =   8175
            Begin MSDataGridLib.DataGrid DataGroup 
               Height          =   5055
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   8916
               _Version        =   393216
               AllowUpdate     =   0   'False
               HeadLines       =   1
               RowHeight       =   15
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
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   " Group Form "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5925
            Left            =   -74880
            TabIndex        =   48
            Top             =   360
            Width           =   8175
            Begin VB.ListBox List1 
               ForeColor       =   &H00400000&
               Height          =   5520
               ItemData        =   "FrmPolicy.frx":3FD28
               Left            =   75
               List            =   "FrmPolicy.frx":3FD2A
               TabIndex        =   49
               Top             =   240
               Width           =   1785
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
               Height          =   5535
               Left            =   1920
               TabIndex        =   50
               Top             =   240
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   9763
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "ID"
               Columns(0).DataField=   "ID"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Form Name"
               Columns(1).DataField=   "Form Name"
               Columns(1).DataWidth=   2500
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   68
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Access"
               Columns(2).DataField=   "Access"
               Columns(2).EditMaskUpdate=   -1  'True
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   4
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "New"
               Columns(3).DataField=   "new"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   4
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Edit"
               Columns(4).DataField=   "edit"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   4
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "Del"
               Columns(5).DataField=   "del"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   6
               Splits(0)._UserFlags=   0
               Splits(0).AllowRowSizing=   0   'False
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).DividerColor=   14215660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=6"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131329"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=4789"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4710"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131584"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=1270"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1191"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=1085"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1005"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131585"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=1164"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1085"
               Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=131585"
               Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(26)=   "Column(5).Width=1164"
               Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1085"
               Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=131585"
               Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   0
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   3
               BorderStyle     =   0
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               AllowArrows     =   0   'False
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000013&,.fgcolor=&H400000&"
               _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=2"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=-1,.fontsize=825"
               _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
               _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.alignment=3,.bgcolor=&H8000000F&"
               _StyleDefs(15)  =   ":id=5,.fgcolor=&H80000012&"
               _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(21)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(22)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(23)  =   "Splits(0).Style:id=47,.parent=1"
               _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=56,.parent=4"
               _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=48,.parent=2,.alignment=2"
               _StyleDefs(26)  =   "Splits(0).FooterStyle:id=49,.parent=3"
               _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
               _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=52,.parent=6"
               _StyleDefs(29)  =   "Splits(0).EditorStyle:id=51,.parent=7"
               _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=53,.parent=8"
               _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=54,.parent=9"
               _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=55,.parent=10"
               _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=57,.parent=11"
               _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=58,.parent=12"
               _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=98,.parent=47"
               _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=48,.alignment=0"
               _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=49"
               _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=51"
               _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=102,.parent=47,.alignment=0"
               _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=48,.alignment=2"
               _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=49"
               _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=51"
               _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=106,.parent=47,.alignment=3"
               _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=48,.alignment=2"
               _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=49"
               _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=51"
               _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=122,.parent=47"
               _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=48,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=49"
               _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=51"
               _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=126,.parent=47"
               _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=123,.parent=48,.alignment=2"
               _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=124,.parent=49"
               _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=125,.parent=51"
               _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=130,.parent=47"
               _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=127,.parent=48,.alignment=2"
               _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=128,.parent=49"
               _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=129,.parent=51"
               _StyleDefs(59)  =   "Named:id=33:Normal"
               _StyleDefs(60)  =   ":id=33,.parent=0,.alignment=2"
               _StyleDefs(61)  =   "Named:id=34:Heading"
               _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(63)  =   ":id=34,.wraptext=-1"
               _StyleDefs(64)  =   "Named:id=35:Footing"
               _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(66)  =   "Named:id=36:Selected"
               _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(68)  =   "Named:id=37:Caption"
               _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(70)  =   "Named:id=38:HighlightRow"
               _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(72)  =   "Named:id=39:EvenRow"
               _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(74)  =   "Named:id=40:OddRow"
               _StyleDefs(75)  =   ":id=40,.parent=33"
               _StyleDefs(76)  =   "Named:id=41:RecordSelector"
               _StyleDefs(77)  =   ":id=41,.parent=34"
               _StyleDefs(78)  =   "Named:id=42:FilterBar"
               _StyleDefs(79)  =   ":id=42,.parent=33"
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Group Form"
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
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Visible         =   0   'False
               Width           =   990
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " List User "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   8175
            Begin TrueOleDBGrid80.TDBGrid TDBListUser 
               Height          =   4215
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   7435
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "User ID"
               Columns(0).DataField=   "User ID"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "NIK"
               Columns(1).DataField=   "empID"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Nama User"
               Columns(2).DataField=   "User name"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Departemen"
               Columns(3).DataField=   "name_dept"
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=3519"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3440"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=5292"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5212"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=4048"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3969"
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
               AllowUpdate     =   0   'False
               Appearance      =   3
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=30,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000013&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2,.locked=0"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=2,.bold=-1"
               _StyleDefs(38)  =   ":id=25,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(39)  =   ":id=25,.fontname=Tahoma"
               _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
               _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
               _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
               _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2,.bold=-1"
               _StyleDefs(44)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(45)  =   ":id=29,.fontname=Tahoma"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.bold=-1,.fontsize=825"
               _StyleDefs(50)  =   ":id=43,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(51)  =   ":id=43,.fontname=Tahoma"
               _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
               _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
               _StyleDefs(54)  =   "Named:id=33:Normal"
               _StyleDefs(55)  =   ":id=33,.parent=0"
               _StyleDefs(56)  =   "Named:id=34:Heading"
               _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(58)  =   ":id=34,.wraptext=-1"
               _StyleDefs(59)  =   "Named:id=35:Footing"
               _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(61)  =   "Named:id=36:Selected"
               _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(63)  =   "Named:id=37:Caption"
               _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(65)  =   "Named:id=38:HighlightRow"
               _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(67)  =   "Named:id=39:EvenRow"
               _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(69)  =   "Named:id=40:OddRow"
               _StyleDefs(70)  =   ":id=40,.parent=33"
               _StyleDefs(71)  =   "Named:id=41:RecordSelector"
               _StyleDefs(72)  =   ":id=41,.parent=34"
               _StyleDefs(73)  =   "Named:id=42:FilterBar"
               _StyleDefs(74)  =   ":id=42,.parent=33"
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " Group Line "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3675
            Left            =   -74880
            TabIndex        =   32
            Top             =   375
            Width           =   8175
            Begin TrueOleDBGrid80.TDBGrid DGForm 
               Height          =   2265
               Left            =   75
               TabIndex        =   63
               Top             =   585
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   3995
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Group"
               Columns(0).DataField=   "Group Form"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Nama Form"
               Columns(1).DataField=   "Nama Form"
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=4948"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4868"
               Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(5)=   "Column(0).Merge=1"
               Splits(0)._ColumnProps(6)=   "Column(0).FilterButton=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=6429"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6350"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   0
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               MultiSelect     =   2
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
               _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
               _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
               _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(38)  =   "Named:id=33:Normal"
               _StyleDefs(39)  =   ":id=33,.parent=0"
               _StyleDefs(40)  =   "Named:id=34:Heading"
               _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(42)  =   ":id=34,.wraptext=-1"
               _StyleDefs(43)  =   "Named:id=35:Footing"
               _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(45)  =   "Named:id=36:Selected"
               _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(47)  =   "Named:id=37:Caption"
               _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(49)  =   "Named:id=38:HighlightRow"
               _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(51)  =   "Named:id=39:EvenRow"
               _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(53)  =   "Named:id=40:OddRow"
               _StyleDefs(54)  =   ":id=40,.parent=33"
               _StyleDefs(55)  =   "Named:id=41:RecordSelector"
               _StyleDefs(56)  =   ":id=41,.parent=34"
               _StyleDefs(57)  =   "Named:id=42:FilterBar"
               _StyleDefs(58)  =   ":id=42,.parent=33"
            End
            Begin VB.TextBox txttempgroupform 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5520
               TabIndex        =   7
               Top             =   720
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.Frame Frame4 
               Caption         =   " Control "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   60
               TabIndex        =   35
               Top             =   2880
               Width           =   8025
               Begin VB.ListBox Listdel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   270
                  ItemData        =   "FrmPolicy.frx":3FD2C
                  Left            =   6105
                  List            =   "FrmPolicy.frx":3FD2E
                  Style           =   1  'Checkbox
                  TabIndex        =   11
                  Top             =   165
                  Width           =   1065
               End
               Begin VB.ListBox ListEdit 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   270
                  ItemData        =   "FrmPolicy.frx":3FD30
                  Left            =   4305
                  List            =   "FrmPolicy.frx":3FD32
                  Style           =   1  'Checkbox
                  TabIndex        =   10
                  Top             =   165
                  Width           =   1065
               End
               Begin VB.ListBox ListNew 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   270
                  ItemData        =   "FrmPolicy.frx":3FD34
                  Left            =   2625
                  List            =   "FrmPolicy.frx":3FD36
                  Style           =   1  'Checkbox
                  TabIndex        =   9
                  Top             =   165
                  Width           =   1065
               End
               Begin VB.ListBox ListAccess 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   270
                  ItemData        =   "FrmPolicy.frx":3FD38
                  Left            =   900
                  List            =   "FrmPolicy.frx":3FD3A
                  Style           =   1  'Checkbox
                  TabIndex        =   8
                  Top             =   165
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Del"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   11
                  Left            =   5745
                  TabIndex        =   39
                  Top             =   285
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Edit"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   10
                  Left            =   3945
                  TabIndex        =   38
                  Top             =   285
                  Width           =   270
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "New"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   9
                  Left            =   2265
                  TabIndex        =   37
                  Top             =   285
                  Width           =   315
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Access"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   8
                  Left            =   285
                  TabIndex        =   36
                  Top             =   285
                  Width           =   495
               End
            End
            Begin VB.TextBox txtTempGroup 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3840
               TabIndex        =   3
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtTempform 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4995
               TabIndex        =   6
               Top             =   720
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CommandButton cmOK 
               Enabled         =   0   'False
               Height          =   330
               Index           =   11
               Left            =   4575
               Picture         =   "FrmPolicy.frx":3FD3C
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   720
               Width           =   330
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "Emp ID"
               Enabled         =   0   'False
               Height          =   330
               Index           =   6
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   4
               Tag             =   "ASM"
               Top             =   720
               Width           =   3480
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Height          =   330
               Index           =   1
               Left            =   1080
               TabIndex        =   2
               Tag             =   "ASM"
               Top             =   225
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               Style           =   2
               ListField       =   "description"
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Form Name"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   34
               Top             =   720
               Width           =   810
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Group Name"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   33
               Top             =   270
               Width           =   885
            End
         End
         Begin VB.CommandButton cmOK 
            Caption         =   "Cancel Group"
            Height          =   420
            Index           =   7
            Left            =   -73560
            TabIndex        =   31
            Top             =   1920
            Width           =   1320
         End
         Begin VB.CommandButton cmOK 
            Caption         =   "Save Group"
            Height          =   420
            Index           =   6
            Left            =   -74880
            TabIndex        =   30
            Top             =   1920
            Width           =   1320
         End
         Begin VB.Frame Frame_group 
            Caption         =   " Group "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   -74880
            TabIndex        =   27
            Top             =   360
            Width           =   8175
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "Nama Group"
               Height          =   330
               Index           =   5
               Left            =   1605
               MaxLength       =   100
               TabIndex        =   28
               Tag             =   "ASM"
               Top             =   480
               Width           =   4320
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Group"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   480
               TabIndex        =   29
               Top             =   540
               Width           =   885
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " Detil User "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2190
            Left            =   120
            TabIndex        =   21
            Top             =   4920
            Width           =   8175
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Index           =   0
               Left            =   1260
               TabIndex        =   18
               Tag             =   "ASM"
               Top             =   1410
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               Style           =   2
               ListField       =   "description"
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin VB.CommandButton cmOK 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   5610
               Picture         =   "FrmPolicy.frx":400C6
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   705
               Width           =   330
            End
            Begin VB.CommandButton cmOK 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               Left            =   5610
               Picture         =   "FrmPolicy.frx":40450
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   345
               Width           =   330
            End
            Begin VB.TextBox txtNameDept 
               Appearance      =   0  'Flat
               DataField       =   "name_dept"
               Height          =   315
               Left            =   6120
               TabIndex        =   41
               Tag             =   "ASM"
               Top             =   1440
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtkodeDept 
               Appearance      =   0  'Flat
               DataField       =   "dept"
               Height          =   315
               Left            =   6120
               TabIndex        =   40
               Tag             =   "ASM"
               Top             =   720
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtpassEnkrip 
               Appearance      =   0  'Flat
               DataField       =   "password"
               Height          =   315
               Left            =   6120
               TabIndex        =   17
               Top             =   1080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "group user"
               Height          =   315
               Index           =   7
               Left            =   6120
               TabIndex        =   14
               Tag             =   "ASM"
               Top             =   360
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "group user"
               Height          =   315
               Index           =   4
               Left            =   4920
               TabIndex        =   19
               Tag             =   "ASM"
               Top             =   1410
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "password"
               Enabled         =   0   'False
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   1260
               MaxLength       =   100
               PasswordChar    =   "*"
               TabIndex        =   16
               Top             =   1050
               Width           =   4800
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "Emp ID"
               Enabled         =   0   'False
               Height          =   330
               Index           =   0
               Left            =   1260
               MaxLength       =   15
               TabIndex        =   12
               Tag             =   "ASM"
               Top             =   330
               Width           =   1440
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "Sql User"
               Enabled         =   0   'False
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1260
               MaxLength       =   100
               TabIndex        =   15
               Tag             =   "ASM"
               Top             =   690
               Width           =   4350
            End
            Begin VB.TextBox txtbox 
               Appearance      =   0  'Flat
               DataField       =   "Full Name"
               Enabled         =   0   'False
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   2730
               MaxLength       =   100
               TabIndex        =   13
               Tag             =   "ASM"
               Top             =   330
               Width           =   2880
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   135
               X2              =   1470
               Y1              =   1710
               Y2              =   1710
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   165
               TabIndex        =   26
               Top             =   1470
               Width           =   435
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   165
               TabIndex        =   24
               Top             =   1125
               Width           =   690
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "User name"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   165
               TabIndex        =   23
               Top             =   765
               Width           =   765
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Employee"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   22
               Top             =   405
               Width           =   690
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   135
               X2              =   1470
               Y1              =   645
               Y2              =   645
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   135
               X2              =   1485
               Y1              =   1005
               Y2              =   1005
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   135
               X2              =   1470
               Y1              =   1365
               Y2              =   1365
            End
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   7215
         Left            =   60
         TabIndex        =   0
         Top             =   75
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   12726
         _Version        =   393217
         Indentation     =   617
         LabelEdit       =   1
         Style           =   3
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
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
   End
End
Attribute VB_Name = "FrmPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents RcLaporan As DBQuick
Attribute RcLaporan.VB_VarHelpID = -1
Private WithEvents RcLaporanUser As DBQuick
Attribute RcLaporanUser.VB_VarHelpID = -1
Private WithEvents RcGroup As DBQuick
Attribute RcGroup.VB_VarHelpID = -1
Private WithEvents RcViewGroup As Recordset
Attribute RcViewGroup.VB_VarHelpID = -1
Private WithEvents RcGroupview As Recordset
Attribute RcGroupview.VB_VarHelpID = -1
Private RccekGroupuser As Recordset

Private RcList As New DBQuick
Private RcDet As New DBQuick
Private RcDetReport As New DBQuick
Private RcPermit As New DBQuick
Private MyData As New DBQuick
Private RCcombo As New DBQuick
Private RCcomboGroup As New DBQuick
Private RcUser As New Recordset
Private RcListReport As New Recordset
Private RcCekUserReport As New Recordset
Private mVarnode As Nodes
Private mLoop, mLoop1 As Integer
'Private oSQLServer2 As New SQLServer2

Private WithEvents MyFrm As frmCaller
Attribute MyFrm.VB_VarHelpID = -1


Private Const TVM_SETBKCOLOR = 4381&
Private CurrentState As String

Private Type USER_INFO
        Name As String
        Comment As String
        UserComment As String
        FullName As String
End Type

Private Type USER_INFO_API
    Name As Long
    Comment As Long
    UserComment As Long
    FullName As Long
End Type

Dim CekUser, CekGroup As Boolean
Dim txtEnkrip, txtDenkrip, TempGroup As String
Dim ParamUser As String


Private Declare Function NetUserEnum Lib "netapi32" (lpServer As Any, ByVal Level As Long, ByVal Filter As Long, lpBuffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal pBuffer As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Const NERR_Success As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&
'Private Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Private Const FILTER_NORMAL_ACCOUNT As Long = &H2&
'Private Const FILTER_PROXY_ACCOUNT As Long = &H4&
'Private Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
'Private Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
'Private Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&
Dim SelectNode As Node
Dim yRow, xCol As Integer

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdTombol_Click(Index As Integer)
Dim I As Integer
Select Case Index
    Case 0  'TAMBAH
        Select Case SSTab1.Tab
            Case 0  'List User
                CurrentState = "ADD"
                EnabledControl True
                RcLaporan.DBRecordset.AddNew 0, MaxForm
                EnabledTombol True
                DataCombo1(0).Text = ""
                frmCallerBaru.Caption = "Find Karyawan"
                frmCallerBaru.Show 1
                TreeView1.Enabled = False
                CekUser = True
            Case 2  'Group User
                CurrentState = "NEW"
                RcGroupview.AddNew
                txtBox(5).SetFocus
                EnabledTombol True
            Case 3  'Detil Group
                CurrentState = "NEW"
                Clear
                aktifText
                EnabledTombol True
            Case Else
                Exit Sub
        End Select
    Case 1  'EDIT
        Select Case SSTab1.Tab
            Case 0  'List User
                If RcLaporan.DBRecordset.Recordcount <> 0 Then
                   TDBListUser.Enabled = False
                   CurrentState = "EDIT"
                   txtBox(1).Enabled = True
                   txtBox(3).Enabled = True
                   DataCombo1(0).Enabled = True
                   cmOK(4).Enabled = True
                   cmOK(5).Enabled = True
                   EnabledTombol True
                   TreeView1.Enabled = False
                Else
                   MessageBox "Data User Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
                   TreeView1.Enabled = True
                End If
                txtBox(1).SetFocus
            Case 2  'Group
                 CurrentState = "EDIT"
                 EnabledTombol True
            Case 3  'Detail Group
                If DataCombo2(1).Text = "" Then
                    MessageBox "Pilih Data Yang Akan Di Edit", vbInformation, msgOkOnly
                Else
                    CurrentState = "EDIT"
                    aktifText
                    EnabledTombol True
                End If
        End Select
    Case 2  'SIMPAN
        Select Case SSTab1.Tab
            Case 0  'List User
               I = MessageBox("Anda yakin untuk menyimpan data", "Konfirmasi", msgYesNo, msgQuestion)
               If I = 1 Then
                    If CurrentState = "ADD" Then
                      ' AddUser
                         I = MaxForm
                         'baru untuk input dept
                         txtEnkrip = EncodeStr64(PROBAencodeString(txtBox(3).Text, Key, True), 68)
                         SendDataToServer (" INSERT INTO [User Table] " & _
                                         " ([User ID], [User Name], EmpID, [group user], dept,employee,password, Name_Dept)" & _
                                         " VALUES     (" & MaxForm & ", N'" & txtBox(1) & "', N'" & txtBox(0) & "'," & txtBox(4) & ",'" & txtkodeDept.Text & "','" & Replace(txtBox(2).Text, "'", "''") & "','" & txtEnkrip & "','" & txtNameDept.Text & "')")
                                         
                    ElseIf CurrentState = "EDIT" Then
                        txtEnkrip = EncodeStr64(PROBAencodeString(txtBox(3).Text, Key, True), 68)
                        SendDataToServer (" UPDATE    [User Table]" & _
                                         " Set [User Name] = N'" & txtBox(1) & "', EmpID = N'" & txtBox(0) & "',[group user] = " & txtBox(4) & ",dept ='" & txtkodeDept.Text & "', employee ='" & Replace(txtBox(2).Text, "'", "''") & "', password ='" & txtEnkrip & "', name_dept='" & txtNameDept.Text & "'" & _
                                         " WHERE     ([User ID] = " & RcLaporan.DBRecordset.Fields("User ID") & ")")
                                            
                        TDBListUser.Enabled = True
                    End If
               End If
               CurrentState = ""
               EnabledControl False
               EnabledTombol False
               TreeView1.Enabled = True
               ViewerLogin
            Case 1  'Akses Form
            Case 2  'Group
                 'CurrentState = "GROUP"
                 I = MessageBox("Anda yakin untuk menyimpan data", "Konfirmasi", msgYesNo, msgQuestion)
                 If I = 1 Then
                    If CurrentState = "NEW" Then
                        SendDataToServer ("SET IDENTITY_INSERT user_table_group ON INSERT INTO user_table_group (id,[group Name]) VALUES(" & MaxGroup & ",'" & txtBox(5) & "')")
                        MessageBox "Data Group Tersimpan", "Group Setup", msgOkOnly, msgInfo
                    ElseIf CurrentState = "EDIT" Then
                        SendDataToServer (" update user_table_group set [group name]='" & txtBox(5) & "' where ID =" & DataGroup.Columns(0).Value)
                        MessageBox "Edit Data Group Tersimpan", "Group Setup", msgOkOnly, msgInfo
                    End If
                End If
                ViewGroup
                ViewerLogin
                EnabledTombol False
            Case 3 'Detil Group
                I = MessageBox("Anda Yakin Untuk Menyimpan Data", "Konformasi", msgYesNo, msgQuestion)
                If I = 1 Then
                    If CurrentState = "NEW" Then
                        Dim row As Variant
                        If DGForm.SelBookmarks.Count > 0 Then
                           For Each row In DGForm.SelBookmarks
                              DGForm.Bookmark = row
                              
                              'SendDataToServer ("SET IDENTITY_INSERT user_table_line ON INSERT INTO user_table_line (id,[group id],[form id],access, new, edit, del) VALUES(" & MaxGroupLine & "," & txtTempGroup.Text & "," & txtTempform.Text & ",'" & ListAccess.Text & "','" & ListNew.Text & "','" & ListEdit.Text & "','" & Listdel.Text & "')")
                              SendDataToServer ("SET IDENTITY_INSERT user_table_line ON INSERT INTO user_table_line (id,[group id],[form id],access, new, edit, del) VALUES(" & MaxGroupLine & "," & txtTempGroup.Text & "," & RcList.DBRecordset.Fields("ID") & ",'" & ListAccess.Text & "','" & ListNew.Text & "','" & ListEdit.Text & "','" & Listdel.Text & "')")
                           Next
                        End If
                        MessageBox "Data Group Line Tersimpan", vbInformation, msgOkOnly, msgInfo
                    ElseIf CurrentState = "EDIT" Then
                        SendDataToServer ("update user_table_line " & _
                                          " set [group id] = " & txtTempGroup.Text & ", [form id] = " & txtTempform.Text & ", access = '" & ListAccess.Text & "', new ='" & ListNew.Text & "', edit ='" & ListEdit.Text & "', del ='" & Listdel.Text & "'" & _
                                          " where (id = " & DGGrid.Columns(0).Value & ")")
                        MessageBox "Data Group Line Tersimpan", vbInformation, msgOkOnly, msgInfo
                    End If
                    NoaktifText
                    Clear
                    ViewGroupLine
                    ViewerLogin
                    EnabledTombol False
                End If
        End Select
    Case 3  'BATAL
        Select Case SSTab1.Tab
            Case 0  'List User
                 Form_Load
                 TreeView1.Enabled = True
                 EnabledControl False
                 EnabledTombol False
                 TDBListUser.Enabled = True
            Case 2  'Group User
                 RcGroupview.CancelBatch
                 EnabledTombol False
                 txtBox(5) = ""
            Case 3  'Detail Group
                 NoaktifText
                 Clear
                 EnabledTombol False
        End Select
    Case 4  'DELETE
        Select Case SSTab1.Tab
            Case 0  'List User
                 If RcLaporan.DBRecordset.Recordcount <> 0 Then
                    I = MessageBox("Anda yakin untuk menghapus User " & txtBox(0) & " - " & txtBox(1), "Peringatan", msgYesNo, msgCrtical)
                    If I = 1 Then
                        DeleteUser Trim(txtBox(1))
                        EnabledTombol False
                        SendDataToServer (" Delete from [User Table] where [User ID] =" & RcLaporan.DBRecordset.Fields(0))
                        TreeView1.Enabled = True
                        EnabledControl False
                        txtBox(3) = ""
                        ViewerLogin
                        CurrentState = ""
                    End If
                End If
            Case 2  'Group
                 If txtBox(5).Text = "" Then
                    MessageBox "Pilih Data Yang Akan Di Hapus", vbInformation, msgOkOnly
                 Else
                    I = MessageBox("Anda yakin untuk Hapus Group", "Konfirmasi", msgYesNo, msgQuestion)
                    If I = 1 Then
                       Dim sql As String
                       Set RccekGroupuser = New Recordset
                       'Cek Data untuk group yang ada relasnya
                       RccekGroupuser.CursorLocation = adUseClient
                       sql = "SELECT [group user] from [user table] where [group user]=" & DataGroup.Columns(0).Value & ""
                       RccekGroupuser.Open sql, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
                       If RccekGroupuser.Recordcount <> 0 Then
                          MessageBox "Tdk Bisa Di Hapus,Relasinya Di User", vbInformation, msgOkOnly
                          Exit Sub
                       Else
                          SendDataToServer (" Delete from user_table_group where ID =" & DataGroup.Columns(0).Value)    '
                          ViewGroup
                       End If
                       ViewerLogin
                       EnabledTombol False
                       EnabledControl False
                    End If
                 End If
            Case 3  'Detail Group
                If DataCombo2(1).Text = "" Then
                    MessageBox "Pilih Data Yang Akan Di Hapus", vbInformation, msgOkOnly
                Else
                    I = MessageBox("Anda yakin untuk Hapus data", "Konfirmasi", msgYesNo, msgQuestion)
                    If I = 1 Then
                        SendDataToServer (" Delete from user_table_line where ID =" & DGGrid.Columns(0).Value)
                        MessageBox "Data Terhapus", vbInformation, msgOkOnly
                        NoaktifText
                        Clear
                        ViewGroupLine
                        EnabledTombol False
                    End If
                End If
        End Select
    Case Else
        Exit Sub
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
Dim rcTemp As New DBQuick

CreateList SelectNode.Text

rcTemp.DBOpen "select id , [group name] from user_table_group where [group name]= '" & DataCombo1(0).Text & "'", CNN, lckLockReadOnly
If rcTemp.DBRecordset.Recordcount > 0 Then
    txtBox(4).Text = rcTemp.DBRecordset.Fields("id")

End If

RCcombo.DBOpen "SELECT [group name]as description from user_table_group order by [group name]", CNN, lckLockReadOnly
Set DataCombo1(0).RowSource = RCcombo.DBRecordset
End Sub

Private Sub DataCombo2_Change(Index As Integer)
Dim rcTempGroup As New DBQuick
rcTempGroup.DBOpen "select id , [group name] from user_table_group where [group name]= '" & DataCombo2(1).Text & "'", CNN, lckLockReadOnly
If rcTempGroup.DBRecordset.Recordcount > 0 Then
    txtTempGroup.Text = rcTempGroup.DBRecordset.Fields("id")
End If
End Sub

Private Sub DataCombo2_Click(Index As Integer, Area As Integer)
RCcomboGroup.DBOpen "SELECT [group name]as description from user_table_group order by [group name]", CNN, lckLockReadOnly
Set DataCombo2(1).RowSource = RCcomboGroup.DBRecordset
End Sub

Private Sub DataGroup_DblClick()
'txtbox(5).Text = DataGroup.Columns(1).Value
End Sub

Private Sub DGGrid_Click()
On Error Resume Next

    DataCombo2(1).Text = DGGrid.Columns(1).Text
    txttempgroupform.Text = DGGrid.Columns(2).Text
    txtBox(6).Text = DGGrid.Columns(3).Text
    
    
    If DGGrid.Columns(4).Text = "True" Then
       ListAccess.Selected(0) = True
    Else
        ListAccess.Selected(1) = True
    End If
    
    If DGGrid.Columns(5).Text = "True" Then
       ListNew.Selected(0) = True
    Else
        ListNew.Selected(1) = True
    End If
    
    If DGGrid.Columns(6).Text = "True" Then
       ListEdit.Selected(0) = True
    Else
        ListEdit.Selected(1) = True
    End If
    
    If DGGrid.Columns(7).Text = "True" Then
       Listdel.Selected(0) = True
    Else
        Listdel.Selected(1) = True
    End If


mLoop1 = xCol
Label1(13).Caption = "&Cari kriteria  berdasarkan " & UCase(DGGrid.Columns(mLoop1).DataField)
DGGrid.col = xCol
Err.Clear

End Sub

Private Sub DGGrid_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex <= 0 Then
   mLoop1 = 0
Else
   mLoop1 = ColIndex
End If
Label1(13) = "&Cari Kriteria Berdasarkan " & UCase(DGGrid.Columns(ColIndex).DataField)
DGGrid.col = ColIndex
Err.Clear
End Sub

Private Sub DGGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
xCol = DGGrid.ColContaining(x)
yRow = DGGrid.RowContaining(Y)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Or KeyCode = vbKeyF10 Then Unload Me
End Sub

Private Sub Form_Load()
Dim sUsers() As String
Set RcLaporan = New DBQuick
Set RcLaporanUser = New DBQuick
Set RcGroup = New DBQuick


SSTab1.Tab = 0
'SSTab1.TabVisible(1) = False
CurrentState = ""
EnabledControl False
'CreateList
Set MyFrm = New frmCaller

Call SendMessage(TreeView1.hwnd, TVM_SETBKCOLOR, 0, ByVal TranslateColor(&H6D4016))
'Call SendMessage(TreeView1.hwnd, TVM_SETBKCOLOR, 0, ByVal TranslateColor(&HC0FFFF))   '&H6D4016))
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
SSTab1.BackColor = Picture2.BackColor
List1.BackColor = txtBox(0).BackColor
'MSHFlexGrid1.BackColor = txtbox(0).BackColor
'DgLaporan.BackColor = txtbox(0).BackColor

ViewerLogin
ViewGroupLine     'Tampil Group Line
ViewListCntrol    'untuk cheklistbox
ViewGroup         'Tampil Group

'TDBGrid1.Columns(2).Locked = True 'access
'TDBGrid1.Columns(3).Locked = True 'insert
'TDBGrid1.Columns(4).Locked = True 'edit
'TDBGrid1.Columns(5).Locked = True 'del


RCcombo.DBOpen "SELECT [group name]as description from user_table_group order by [group name]", CNN, lckLockReadOnly
Set DataCombo1(0).RowSource = RCcombo.DBRecordset

RCcomboGroup.DBOpen "SELECT [group name]as description from user_table_group order by [group name]", CNN, lckLockReadOnly
Set DataCombo2(1).RowSource = RCcomboGroup.DBRecordset
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

Set RcLaporan = Nothing
CloseMenuAll
Set MyFrm = Nothing
Err.Clear
End Sub

Private Sub CmOK_Click(Index As Integer)
Dim I As Integer
Select Case Index
       Case 0:
            CurrentState = "ADD"
            EnabledControl True
            RcLaporan.DBRecordset.AddNew 0, MaxForm
            EnabledTombol True
            'OpenPartner 0
            DataCombo1(0).Text = ""
            frmCallerBaru.Caption = "Find Karyawan"
            frmCallerBaru.Show 1
            TreeView1.Enabled = False
            CekUser = True
       Case 1:  'EDIT

       Case 2:
            If RcLaporan.DBRecordset.Recordcount <> 0 Then
               I = MessageBox("Anda yakin untuk menghapus User " & txtBox(0) & " - " & txtBox(1), "Peringatan", msgYesNo, msgCrtical)
               If I = 1 Then
                  DeleteUser Trim(txtBox(1))
                  EnabledTombol False
                  SendDataToServer (" Delete from [User Table] where [User ID] =" & RcLaporan.DBRecordset.Fields(0))
                  TreeView1.Enabled = True
                  EnabledControl False
                  txtBox(3) = ""
                  ViewerLogin
                  CurrentState = ""
               End If
            End If
       Case 3:  'SAVE TAB 0
            If RcLaporan.DBRecordset.Recordcount <> 0 Then
              
               I = MessageBox("Anda yakin untuk menyimpan data", "Konfirmasi", msgYesNo, msgQuestion)
               If I = 1 Then
                    If CurrentState = "ADD" Then
                      ' AddUser
                       I = MaxForm
                                                           
                         'baru untuk input dept
                         txtEnkrip = EncodeStr64(PROBAencodeString(txtBox(3).Text, Key, True), 68)
                         SendDataToServer (" INSERT INTO [User Table] " & _
                                         " ([User ID], [User Name], EmpID, [group user], dept,employee,password, Name_Dept)" & _
                                         " VALUES     (" & MaxForm & ", N'" & txtBox(1) & "', N'" & txtBox(0) & "'," & txtBox(4) & ",'" & txtkodeDept.Text & "','" & Replace(txtBox(2).Text, "'", "''") & "','" & txtEnkrip & "','" & txtNameDept.Text & "')")

                       'CopyForm I
                       'CopyLaporan I
                    ElseIf CurrentState = "EDIT" Then
                       DeleteUser RcLaporan.DBRecordset.Fields("Sql User")
                       'AddUser
'                       SendDataToServer (" UPDATE    [User Table]" & _
'                                         " Set [User Name] = N'" & txtbox(1) & "', EmpID = N'" & txtbox(0) & "'" & _
'                                         " WHERE     ([User ID] = " & RcLaporan.DBRecordset.Fields("User ID") & ")")
                                         
                        
                        
'                        SendDataToServer (" UPDATE    [User Table]" & _
'                                         " Set [User Name] = N'" & txtBox(1) & "', EmpID = N'" & txtBox(0) & "',[group user] = " & txtBox(4) & "" & _
'                                         " WHERE     ([User ID] = " & RcLaporan.DBRecordset.Fields("User ID") & ")")
                        txtEnkrip = EncodeStr64(PROBAencodeString(txtBox(3).Text, Key, True), 68)
                                         
                        SendDataToServer (" UPDATE    [User Table]" & _
                                         " Set [User Name] = N'" & txtBox(1) & "', EmpID = N'" & txtBox(0) & "',[group user] = " & txtBox(4) & ",dept ='" & txtkodeDept.Text & "', employee ='" & Replace(txtBox(2).Text, "'", "''") & "', password ='" & txtEnkrip & "', name_dept='" & txtNameDept.Text & "'" & _
                                         " WHERE     ([User ID] = " & RcLaporan.DBRecordset.Fields("User ID") & ")")
                        
                    
                        TDBListUser.Enabled = True
                    End If
               End If
               
            End If
            CurrentState = ""
            EnabledControl False
            EnabledTombol False
            cmOK(0).SetFocus
            TreeView1.Enabled = True
            ViewerLogin
'            TDBGrid1.Columns(2).Locked = True 'access
'            TDBGrid1.Columns(3).Locked = True 'insert
'            TDBGrid1.Columns(4).Locked = True 'edit
'            TDBGrid1.Columns(5).Locked = True
       Case 4: LookupUser GetSetting(App.EXEName, "Lisence Profile", "Servername")
       Case 5: frmCallerBaru.Show 1 'OpenPartner 0
       Case 8:
            'CurrentState = "GROUP"
            I = MessageBox("Anda yakin untuk menyimpan data", "Konfirmasi", msgYesNo, msgQuestion)
            If I = 1 Then
                If SendDataToServer("SET IDENTITY_INSERT user_table_group ON INSERT INTO user_table_group (id,[group Name]) VALUES(" & MaxGroup & ",'" & txtBox(5) & "')") Then
                    MessageBox "Data Group Tersimpan", "Group Setup", msgOkOnly, msgInfo
'                    txtBox(5).Text = ""
                End If
            End If
            ViewGroup
       Case 9:
            I = MessageBox("Anda Yakin Untuk Menyimpan Data", "Konformasi", msgYesNo, msgQuestion)
            If I = 1 Then
                If CurrentState = "NEW" Then
                    SendDataToServer ("SET IDENTITY_INSERT user_table_line ON INSERT INTO user_table_line (id,[group id],[form id],access, new, edit, del) VALUES(" & MaxGroupLine & "," & txtTempGroup.Text & "," & txtTempform.Text & ",'" & ListAccess.Text & "','" & ListNew.Text & "','" & ListEdit.Text & "','" & Listdel.Text & "')")
                    MessageBox "Data Group Line Tersimpan", vbInformation, msgOkOnly, msgInfo
                ElseIf CurrentState = "EDIT" Then
                    SendDataToServer ("update user_table_line " & _
                                      " set [group id] = " & txtTempGroup.Text & ", [form id] = " & txtTempform.Text & ", access = '" & ListAccess.Text & "', new ='" & ListNew.Text & "', edit ='" & ListEdit.Text & "', del ='" & Listdel.Text & "'" & _
                                      " where (id = " & DGGrid.Columns(0).Value & ")")
                    MessageBox "Data Group Line Tersimpan", vbInformation, msgOkOnly, msgInfo
                End If
                NoaktifText
                Clear
                ViewGroupLine
                cmOK(14).Enabled = True
                cmOK(9).Enabled = False
                cmOK(10).Enabled = False
                cmOK(12).Enabled = True
                cmOK(13).Enabled = True
                
            End If
       Case 10:
            NoaktifText
            Clear
            cmOK(14).Enabled = True
            cmOK(9).Enabled = False
            cmOK(10).Enabled = False
            cmOK(12).Enabled = True
            cmOK(13).Enabled = True
                
       Case 11:
            OpenPartner 11
            CekGroup = True
       Case 12:
            If DataCombo2(1).Text = "" Then
                MessageBox "Pilih Data Yang Akan Di Hapus", vbInformation, msgOkOnly
            Else
                I = MessageBox("Anda yakin untuk Hapus data", "Konfirmasi", msgYesNo, msgQuestion)
                If I = 1 Then
                    SendDataToServer (" Delete from user_table_line where ID =" & DGGrid.Columns(0).Value)
                    MessageBox "Data Terhapus", vbInformation, msgOkOnly
                    ViewGroupLine
                End If
            End If
       Case 13:

       Case 14:

            
        Case 15:
            'tombol cancel
            Form_Load
            TreeView1.Enabled = True
            EnabledControl False
            EnabledTombol False
            TDBListUser.Enabled = True
        Case 16:
            If txtBox(5).Text = "" Then
                MessageBox "Pilih Data Yang Akan Di Hapus", vbInformation, msgOkOnly
            Else
                I = MessageBox("Anda yakin untuk Hapus Group", "Konfirmasi", msgYesNo, msgQuestion)
                If I = 1 Then
                   Dim sql As String
                   Set RccekGroupuser = New Recordset
                   'Cek Data untuk group yang ada relasnya
                   RccekGroupuser.CursorLocation = adUseClient
                   sql = "SELECT [group user] from [user table] where [group user]=" & DataGroup.Columns(0).Value & ""
                   RccekGroupuser.Open sql, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
                   If RccekGroupuser.Recordcount <> 0 Then
                      MessageBox "Tdk Bisa Di Hapus,Relasinya Di User", vbInformation, msgOkOnly
                      Exit Sub
                   Else
                      SendDataToServer (" Delete from user_table_group where ID =" & DataGroup.Columns(0).Value)
'                      MessageBox "Data Terhapus", vbInformation, msgOkOnly
                      ViewGroup
                    End If
                End If
            End If
        Case 17
'            DataGroup.AllowAddNew = True
            RcGroupview.AddNew
            txtBox(5).SetFocus
End Select
End Sub



Private Sub DgUser_DblClick()
'DgUser.Enabled = False
'DgUser.Visible = False
'If DgUser.Columns(0).Value <> "" Then txtbox(1) = DgUser.Columns(0)
txtBox(3).SetFocus
End Sub

Private Sub DgUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call DgUser_DblClick
End Sub

Private Sub DgUser_LostFocus()
'DgUser.Enabled = False
'DgUser.Visible = False
txtBox(3).SetFocus
End Sub

Private Sub DgUser_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If DgUser.Columns(0).Value <> "" Then txtBox(1) = DgUser.Columns(0)
End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    On Error Resume Next
    If OleTranslateColor(clr, hPal, TranslateColor) Then
       TranslateColor = -1
    End If
    Err.Clear
End Function

Private Sub Form_Unload(Cancel As Integer)
Set FrmPolicy = Nothing
RCcombo.CloseDB
Set RCcombo = Nothing

RcDet.CloseDB
Set RcDet = Nothing

RcLaporanUser.CloseDB
Set RcLaporanUser = Nothing
End Sub

Private Sub List1_Click()
'If CurrentState <> "EDIT" Then
   ' If RcLaporan.DBRecordset.Recordcount <> 0 Then
       'OpenDetail RcLaporan.DBRecordset.Fields(0), List1.Text
      ' If List1.Selected(List1.ListCount - 1) = True Then
        '  OpenDetail List1.Text
       ' End If
    'Else
      ' OpenDetail 0, List1.Text
      ' OpenDetail DataCombo1(0), List1.Text
       OpenDetailGrid TreeView1.SelectedItem.Text, List1.Text
    'End If
'End If
End Sub

Private Sub List1_DblClick()
If CurrentState = "EDIT" Then
   If RcDet.DBRecordset.Recordcount > 0 Then
'      List1.ItemData(List1.List) = OldData(RcDet.DBRecordset.Fields(0), List1.Text)
   End If
End If
End Sub

Private Sub List1_ItemCheck(Item As Integer)
Dim I As Integer
Dim Avdata As Variant
If CurrentState = "EDIT" Then
   With RcDet.DBRecordset
        If .Recordcount <> 0 Then
           Avdata = .Getrows(.Recordcount, adBookmarkFirst)
           For I = 0 To UBound(Avdata, 2)
               'MsgBox .Fields(5).Name
               SendDataToServer (" UPDATE [Detail User Table]" & _
                                 " Set [Laporan] = " & BoolToInt(List1.Selected(Item)) & " WHERE     ([User ID] = " & RcLaporan.DBRecordset.Fields("User ID") & ") AND ([idx] = '" & Avdata(5, I) & "')")
           Next I
           SaveSetting App.EXEName, "Lisence Profile", List1.Text, CBool(List1.Selected(Item))
        End If
   End With
Else
   If RcDet.DBRecordset.Recordcount > 0 Then
     ' List1.Selected(Item) = OldData(RcLaporan.DBRecordset.Fields(0), List1.Text)
   End If
End If
Set Avdata = Nothing
End Sub

'Private Sub List2_Click()
'Text1.Text = List2.ListIndex
'End Sub

Private Sub ListAccess_ItemCheck(Item As Integer)
On Error Resume Next
Select Case Item
    Case 0:
        ListAccess.Selected(0) = True
        ListAccess.Selected(1) = False
        Exit Sub
    Case 1:
        ListAccess.Selected(1) = True
        ListAccess.Selected(0) = False
        Exit Sub
End Select

       
End Sub

Private Sub Listdel_ItemCheck(Item As Integer)
On Error Resume Next
Select Case Item
    Case 0:
        Listdel.Selected(0) = True
        Listdel.Selected(1) = False
        Exit Sub
    Case 1:
        Listdel.Selected(1) = True
        Listdel.Selected(0) = False
        Exit Sub
End Select
End Sub

Private Sub ListEdit_ItemCheck(Item As Integer)
On Error Resume Next
Select Case Item
    Case 0:
        ListEdit.Selected(0) = True
        ListEdit.Selected(1) = False
        Exit Sub
    Case 1:
        ListEdit.Selected(1) = True
        ListEdit.Selected(0) = False
        Exit Sub
End Select
End Sub

Private Sub ListNew_ItemCheck(Item As Integer)
On Error Resume Next
Select Case Item
    Case 0:
        ListNew.Selected(0) = True
        ListNew.Selected(1) = False
        Exit Sub
    Case 1:
        ListNew.Selected(1) = True
        ListNew.Selected(0) = False
        Exit Sub
End Select
End Sub


Private Sub MyFrm_BeforeUnload()
Dim RcUsr As New DBQuick
If CekUser = True Then
    RcUsr.DBOpen "SELECT     [User Name] AS [Full Name] FROM         [User Table] GROUP BY [User Name] ORDER BY [User Name]", CNN, lckLockReadOnly
    With RcUsr.DBRecordset
         If .Recordcount <> 0 Then
             .Filter = "[Full Name] = '" & txtBox(2) & "'"
             If .Recordcount <> 0 Then
                MessageBox "Record -> " & txtBox(2) & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
                RcLaporan.DBRecordset.CancelBatch adAffectCurrent
                EnabledControl False
                EnabledTombol False
                RcUsr.CloseDB
                Set RcUsr = Nothing
                TreeView1.Enabled = True
                ViewerLogin
                cmOK(0).SetFocus
                Exit Sub
             End If
         End If
    End With
    
    If IsNull(RcLaporan.DBRecordset.Fields(2)) = True Or RcLaporan.DBRecordset.Fields(2) = "" Then
       RcLaporan.DBRecordset.CancelBatch adAffectCurrent
       EnabledControl False
       EnabledTombol False
       RcUsr.CloseDB
       Set RcUsr = Nothing
       TreeView1.Enabled = True
       ViewerLogin
       cmOK(0).SetFocus
       Exit Sub
    End If
'    cmOK(4).SetFocus
End If
End Sub

Private Sub MyFrm_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)


If CekUser = True Then
    RcLaporan.DBRecordset.Fields("Emp ID") = MyFrm.GetFieldByName(1)
    txtBox(3) = MyFrm.GetFieldByName(0)
    txtBox(1) = MyFrm.GetFieldByName(0)
    txtBox(2) = MyFrm.GetFieldByName(0)
ElseIf CekGroup = True Then
    RcGroup.DBOpen "SELECT     User_Table_group.ID, [User_Table_group].[group Name] FROM  User_table_group ", CNN, lckLockBatch
    RcGroup.DBRecordset.Fields("group name") = MyFrm.GetFieldByName(1)
    txttempgroupform.Text = MyFrm.GetFieldByName(1)
    txtBox(6) = MyFrm.GetFieldByName(0)
    
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub RcLaporan_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
With pRecordset
     If .Recordcount <> 0 Then
        'OpenDetail DataCombo1(0), List1.Text
        OpenDetailGrid DataCombo1(0), List1.Text
     Else
        'OpenDetail DataCombo1(0), List1.Text
        OpenDetailGrid DataCombo1(0), List1.Text
     End If
End With
Err.Clear
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
RcPermit.CloseDB
Select Case SSTab1.Tab
    Case 0:
    Case 1:
'            If RcLaporan.DBRecordset.Recordcount <> 0 Then
'               'OpenLaporan RcLaporan.DBRecordset.Fields("User ID")
'            End If
'             OpenDetailReport ParamUser
        
        If Not SelectNode Is Nothing Then
            CreateList SelectNode.Text
            OpenDetailGrid TreeView1.SelectedItem.Text, List1.Text
        End If
    Case 3:
        RcList.DBOpen "SELECT user_table_form.ID, user_table_form.[group form] as [Group Form], user_table_form.[form name] as [Nama Form]  FROM  user_table_form order by [group form]", CNN, lckLockReadOnly
        Set DGForm.DataSource = RcList.DBRecordset
End Select
End Sub

Private Sub TDBGrid1_ColEdit(ByVal ColIndex As Integer)
SendDataToServer ("update user_table_line " & _
                                      " set access = " & FQtyUser(TDBGrid1.Columns(2).Value) & ", new =" & FQtyUser(TDBGrid1.Columns(3).Value) & ", edit =" & FQtyUser(TDBGrid1.Columns(4).Value) & ", del =" & FQtyUser(TDBGrid1.Columns(5).Value) & "" & _
                                      " where (id = " & TDBGrid1.Columns(0).Value & ")")
End Sub

Private Sub TDBListUser_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If RcLaporanUser.DBRecordset.Recordcount <> 0 Then OpenDB Val(TDBListUser.Columns(0).Value)
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Set SelectNode = Node
If Not Node.Parent Is Nothing Then
'    Select Case SSTab1.Tab
'        Case 0:
'            If SSTab1.Tab = 0 Then
            Frame5.Caption = " List User " & Node 'TreeView1.SelectedItem.Text
            OpenDBListUser Val(Replace(Node.Key, "User", ""))
'            End If
'        Case 1:
            Frame2.Caption = " Group Form " & Node 'TreeView1.SelectedItem.Text
            CreateList Node.Text
            OpenDetailGrid Node.Text, List1.Text
'    End Select
End If
End Sub

'Private Sub ViewerLogin(oCollection As Object, Optional ByVal Deletelog As Byte)

Private Sub ViewerLogin()
Dim I As Integer
Dim mVarStr As String
Dim oObject As Variant
Dim oNode As Node
Dim Rc As New DBQuick
    
    TreeView1.Nodes.Clear
    Set mVarnode = TreeView1.Nodes
    With mVarnode.Add(, , "User0", "Grup User", "biru")
         .Expanded = True
         .Bold = True
         .BackColor = &H6D4016
         .ForeColor = vbWhite
    End With
    
   ' Rc.DBOpen " SELECT     [User ID], [User Name] FROM         [User Table] ORDER BY [User ID]", CNN, lckLockReadOnly
    Rc.DBOpen " SELECT     ID,  [Group Name] FROM         user_table_group ORDER BY [Group Name]", CNN, lckLockReadOnly
    With Rc.DBRecordset
         If .Recordcount <> 0 Then
            oObject = .Getrows(.Recordcount, adBookmarkFirst)
            For I = 0 To UBound(oObject, 2)
               With mVarnode.Add("User0", tvwChild, "User" & oObject(0, I), oObject(1, I), "ijo")
                .BackColor = &H6D4016
                .ForeColor = vbWhite
                End With
            Next
            Set oNode = TreeView1.Nodes(2)
'            TreeView1.SetFocus
            TreeView1.Nodes(2).Selected = True
            Call TreeView1_NodeClick(oNode)
         Else
            Set oNode = TreeView1.Nodes(1)
            TreeView1.Nodes(1).Selected = True
            Call TreeView1_NodeClick(oNode)
            'TreeView1.SelectedItem = oNode
         End If
    End With

OpenDB Val(Replace(oNode.Key, "User", ""))
Set oNode = Nothing
Set oObject = Nothing
Rc.CloseDB
Set Rc = Nothing
Exit Sub
    Err.Clear
End Sub

Private Function DeleteUser(ByVal pUserName As String) As Boolean
On Error GoTo Hell
Dim Cmm As New Command
Set Cmm.ActiveConnection = CNN
Cmm.CommandType = adCmdStoredProc
Cmm.Prepared = True
Cmm.CommandText = "DeleteUser"
Cmm.Parameters("@Username").Value = pUserName
Cmm.Execute
DeleteUser = True
Exit Function
Hell:
    'MessageBox Err.Description, "Drop Login Error", msgOkOnly
    Err.Clear
End Function

Private Sub LookupUser(Optional ByVal sServerName As String)
On Error GoTo Hell
Dim sUsers() As String
Dim lctr
Set RcUser = Nothing
'If DgUser.Visible = True Then
'   DgUser.Enabled = False
'   DgUser.Visible = False
'   Exit Sub
'End If
Set RcUser = New Recordset
With RcUser
    .Fields.Append "User Domain", adBSTR
    .Open

    If sServerName <> "" Then
       GetUsers sUsers, sServerName
    Else
       GetUsers sUsers
    End If
    For lctr = LBound(sUsers) To UBound(sUsers)
        .AddNew (0), sUsers(lctr)
    Next
End With
'Set DgUser.DataSource = RcUser
If RcUser.Recordcount <> 0 Then
   RcUser.MoveFirst
'   DgUser.Columns(0).width = 4545.071
'   DgUser.Columns(0).Locked = True
'   DgUser.Enabled = True
'   DgUser.Visible = True
'   DgUser.col = 0
'   DgUser.ZOrder (0)
'   DgUser.Refresh
'   DgUser.SetFocus
End If
Hell:
    'MsgBox Err.Description
    Err.Clear
End Sub

Private Sub LookUpLocal()
Dim Rc As New DBQuick
Rc.DBOpen "", CNN, lckLockReadOnly
With Rc.DBRecordset
     
End With
End Sub

Private Function GetUsers(UserNames() As String, Optional ServerName As String = "") As Boolean
    Dim lptrStrBuffer As Long
    Dim lRet As Long
    Dim lUsersRead As Long
    Dim lTotalUsers As Long
    Dim lHnd As Long
    Dim etUserInfo As USER_INFO_API
    Dim bytServerName() As Byte
    Dim lElement As Long
    Dim Users() As USER_INFO
    Dim I As Long
    ReDim Users(0) As USER_INFO
    ReDim UserNames(0) As String
    If Trim(ServerName) = "" Then
        bytServerName = vbNullString
    Else
        If InStr(ServerName, "\\") = 1 Then
            bytServerName = ServerName & vbNullChar
        Else
            bytServerName = "\\" & ServerName & vbNullChar
        End If
    End If
    lHnd = 0

 Do
         If Trim(ServerName) = "" Then
             lRet = NetUserEnum(vbNullString, 10, _
              FILTER_NORMAL_ACCOUNT, lptrStrBuffer, 1, _
               lUsersRead, lTotalUsers, lHnd)
         Else
             lRet = NetUserEnum(bytServerName(0), 10, _
              FILTER_NORMAL_ACCOUNT, lptrStrBuffer, 1, _
                lUsersRead, lTotalUsers, lHnd)
         End If
         For I = 0 To lUsersRead - 1
           CopyMem etUserInfo, ByVal lptrStrBuffer + Len(etUserInfo) * I, _
 Len(etUserInfo)
           If Users(0).Name = "" Then
               lElement = 0
           Else
               lElement = UBound(Users) + 1
           End If
           ReDim Preserve Users(lElement) As USER_INFO
           Users(lElement).Name = PtrToString(etUserInfo.Name)
           Users(lElement).Comment = PtrToString(etUserInfo.Comment)
           Users(lElement).UserComment = PtrToString(etUserInfo.UserComment)
           Users(lElement).FullName = PtrToString(etUserInfo.FullName)
            ReDim Preserve UserNames(lElement)
           UserNames(lElement) = Users(lElement).Name
         Next
         If lptrStrBuffer Then
             Call NetApiBufferFree(lptrStrBuffer)
         End If
         DoEvents
         If lRet = NERR_Success Then Exit Do
     Loop While lRet = ERROR_MORE_DATA
 GetUsers = True
    Exit Function
ErrHandler:
On Error Resume Next
Call NetApiBufferFree(lptrStrBuffer)
End Function

Private Function PtrToString(lpString As Long) As String
    Dim bytBuffer() As Byte
    Dim lLen As Long
    If lpString Then
        lLen = lstrlenW(lpString) * 2
        If lLen Then
            ReDim bytBuffer(0 To (lLen - 1)) As Byte
            CopyMem bytBuffer(0), ByVal lpString, lLen
            PtrToString = bytBuffer
        End If
    End If
End Function

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcList.DBOpen "SELECT Employees.FullName AS Employee, Employees.EmpID AS [Employee ID] FROM         [User Table] RIGHT OUTER JOIN Employees ON [User Table].EmpID = Employees.EmpID WHERE     ([User Table].[User ID] IS NULL)", CNN, lckLockReadOnly
       Case 11:
            RcList.DBOpen "SELECT user_table_form.[form name] as [Nama Form] ,user_table_form.[group form] as [Group Form] FROM  user_table_form order by [group form]", CNN, lckLockReadOnly
End Select
If RcList.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            MyFrm.FromTagActive = "Employee"
            Set MyFrm.FormData = RcList.DBRecordset
            MyFrm.LookUp Me
          Case 11:
            MyFrm.FromTagActive = "Group Form Name"
            Set MyFrm.FormData = RcList.DBRecordset
            MyFrm.LookUp Me
   End Select
   
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
'   If MyDDE.ChildRecordset.Recordcount <> 0 Then
'      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Function AddUser() As Boolean
On Error GoTo Hell
Dim Cmm As New Command
Dim NmAsli As String
Dim I As Integer
Set Cmm.ActiveConnection = CNN
Cmm.CommandType = adCmdStoredProc
Cmm.Prepared = True
Cmm.CommandText = "adduser"
'If CurrentState = "ADD" Then
   I = InStr(txtBox(1), "\")
   If I <> 0 Then
      NmAsli = Trim(Mid(txtBox(1), I + 1, Len(txtBox(1)) - I))
   Else
      NmAsli = txtBox(1)
   End If
'Else
'   I = InStr(txtBox(1), "\")
'   If I <> 0 Then
'      NmAsli = Trim(Mid(txtBox(1), I + 1, Len(txtBox(1)) - I))
'   Else
'      NmAsli = txtBox(1)
'   End If
'End If

Cmm.Parameters("@Username").Value = NmAsli
Cmm.Parameters("@Password").Value = txtBox(3)
Cmm.Execute
'MsgBox Cmm.Parameters("@IsOk")
AddUser = True
Exit Function
Hell:
    AddUser = False
    'MsgBox Err.Description
    Err.Clear
End Function

Private Sub OpenDB(ByVal Param As Integer)

RcLaporan.DBOpen "SELECT     [User Table].[User ID], [User Table].[User Name] AS [Sql User], [User Table].EmpID AS [Emp ID], [user table].employee as [full name], [user table].[group user],[user table].password, [user table].dept, [user table].name_dept FROM [user table] WHERE  ([User Table].[User ID] = " & Param & ")", CNN, lckLockBatch
'txtBox(3).Text = PROBAdecodeString(DecodeStr64(txtpassEnkrip.Text), key, True)
With RcLaporan.DBRecordset
     Set txtBox(0).DataSource = .DataSource
     Set txtBox(1).DataSource = .DataSource
     Set txtBox(2).DataSource = .DataSource
     Set txtpassEnkrip.DataSource = .DataSource 'password
     txtBox(3).Text = PROBAdecodeString(DecodeStr64(txtpassEnkrip.Text), Key, True) 'membaca data Enkrip password
     Set txtBox(4).DataSource = .DataSource
     Set txtkodeDept.DataSource = .DataSource
     Set txtNameDept.DataSource = .DataSource
     
     If .Recordcount <> 0 Then
        List1.Enabled = True
     Else
        List1.Enabled = False
     End If
     ParamUser = Param  'digunakan variabel untuk baca di user report
End With

End Sub

Private Sub OpenDBListUser(ByVal Param As Integer)

'RcLaporan.DBOpen "SELECT     [User Table].[User ID], [User Table].[User Name] AS [Sql User], [User Table].EmpID AS [Emp ID], [user table].employee as [full name], [user table].[group user],[user table].password, [user table].dept, [user table].name_dept FROM [user table] WHERE  ([User Table].[User ID] = " & Param & ")", CNN, lckLockBatch


RcLaporanUser.DBOpen "SELECT dbo.[user table].[User ID], dbo.user_table_group.ID, dbo.user_table_group.[Group Name], dbo.[user table].[User Name], dbo.[user table].EmpID, dbo.[user table].Name_Dept " & _
                  "FROM   dbo.user_table_group INNER JOIN " & _
                  "dbo.[user table] ON dbo.user_table_group.id = dbo.[user table].[Group User] WHERE     (dbo.user_table_group.ID = " & Param & ")", CNN, lckLockBatch

Set TDBListUser.DataSource = RcLaporanUser.DBRecordset

End Sub

Private Sub CreateList(sGroup As String)
Dim Rc As New DBQuick
Dim sql As String
Dim I As Integer
Dim Avdata As Variant
List1.Clear

sql = "SELECT   User_table_group.[Group Name], User_Table_Form.[Group Form]" & _
                " FROM  User_table_group INNER JOIN" & _
                " User_table_line ON dbo.User_table_group.id = dbo.User_table_line.[Group ID] INNER JOIN" & _
                " User_Table_Form ON dbo.User_table_line.[Form ID] = dbo.User_Table_Form.ID" & _
                " WHERE (User_table_group.[Group Name] = '" & sGroup & " ')" & _
                " GROUP BY User_table_group.[Group Name], User_Table_Form.[Group Form]"

Rc.DBOpen sql, CNN, lckLockReadOnly

With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            'List1.AddItem Avdata(0, I)
            List1.AddItem Avdata(1, I)
        Next I
     End If
End With
If List1.ListCount <> 0 Then List1.ListIndex = 0
Set Rc = Nothing
Set Avdata = Nothing
End Sub

Private Sub createID_lineGroup()
Dim Rc As New DBQuick
Dim sql As String
Dim I As Integer
Dim Avdata As Variant
'List2.Clear

sql = "SELECT   User_table_group.[Group Name], User_Table_line.[id]" & _
                " FROM  User_table_group INNER JOIN" & _
                " User_table_line ON dbo.User_table_group.id = dbo.User_table_line.[Group ID] INNER JOIN" & _
                " User_Table_Form ON dbo.User_table_line.[Form ID] = dbo.User_Table_Form.ID" & _
                " WHERE (User_table_group.[Group Name] = '" & DataCombo1(0).Text & " ')" & _
                " GROUP BY User_table_group.[Group Name], User_Table_line.[id]"

Rc.DBOpen sql, CNN, lckLockReadOnly

With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            'List1.AddItem Avdata(0, I)
            'List2.AddItem Avdata(1, I)
        Next I
     End If
End With
If List1.ListCount <> 0 Then List1.ListIndex = 0
Set Rc = Nothing
Set Avdata = Nothing
End Sub

Private Sub CopyForm(ByVal Param As Integer)
Dim Rc As New DBQuick
Dim RcLap As New DBQuick
Dim I As Integer
Dim Avdata As Variant

Rc.DBOpen "SELECT [Detail User Table].Idx FROM  [Detail User Table]  WHERE     ([Detail User Table].[User ID] = " & Param & ")  ", CNN, lckLockReadOnly
With Rc.DBRecordset
     'MsgBox .Source
     If .Recordcount <> 0 Then
        RcLap.DBOpen "SELECT     Idx, [Group Table], [Form List] FROM         [Form Table] WHERE     (NOT (Idx IN   (SELECT     Idx  FROM          [Detail User Table] WHERE      ([User ID] = " & Param & ")))) ORDER BY [Group Table]", CNN, lckLockReadOnly
        If RcLap.DBRecordset.Recordcount <> 0 Then
           Avdata = RcLap.DBRecordset.Getrows(RcLap.DBRecordset.Recordcount, adBookmarkFirst)
           For I = 0 To UBound(Avdata, 2)
               SendDataToServer (" INSERT INTO [Detail User Table]" & _
                                 " ([User ID], Idx)" & _
                                 " VALUES     (" & Param & ", '" & Avdata(0, I) & "')")
           Next I
        End If
     Else
        RcLap.DBOpen "select * from [Form Table] Order by IDx", CNN, lckLockReadOnly
        If RcLap.DBRecordset.Recordcount <> 0 Then
            Avdata = RcLap.DBRecordset.Getrows(RcLap.DBRecordset.Recordcount, adBookmarkFirst)
            For I = 0 To UBound(Avdata, 2)
                SendDataToServer (" INSERT INTO [Detail User Table]" & _
                                  " ([User ID], Idx)" & _
                                  " VALUES     (" & Param & ", '" & Avdata(0, I) & "')")
            Next I
        End If
     End If
End With
Set Rc = Nothing
Set Avdata = Nothing
End Sub

Private Sub CopyLaporan(ByVal Param As Integer)
Dim Rc As New DBQuick
Dim RcLap As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Rc.DBOpen "SELECT     [User ID] FROM         [Report Permit] WHERE     ([User ID] = " & Param & ")", CNN, lckLockReadOnly
With Rc.DBRecordset
     'MsgBox .Source
     If .Recordcount <> 0 Then
        RcLap.DBOpen "SELECT   Idx, [Group Table], [Form List] FROM         [Form Table] WHERE     (NOT (Idx IN   (SELECT     Idx  FROM          [Detail User Table] WHERE      ([User ID] = " & Param & ")))) ORDER BY [Group Table]", CNN, lckLockReadOnly
        If RcLap.DBRecordset.Recordcount <> 0 Then
           Avdata = RcLap.DBRecordset.Getrows(RcLap.DBRecordset.Recordcount, adBookmarkFirst)
           For I = 0 To UBound(Avdata, 2)
               SendDataToServer (" INSERT INTO [Detail User Table]" & _
                                 " ([User ID], Idx)" & _
                                 " VALUES   (" & Param & ", '" & Avdata(0, I) & "')")
           Next I
        End If
     Else
       ' RcLap.DBOpen "SELECT IDReport FROM [Report Modules] ORDER BY ModulesName", CNN, lckLockReadOnly
        RcLap.DBOpen "SELECT IDReport FROM [Report Modules] ", CNN, lckLockReadOnly
        If RcLap.DBRecordset.Recordcount <> 0 Then
            Avdata = RcLap.DBRecordset.Getrows(RcLap.DBRecordset.Recordcount, adBookmarkFirst)
            For I = 0 To UBound(Avdata, 2)
                SendDataToServer (" INSERT INTO  [Report Permit]" & _
                                  " ([User ID], IDReport)" & _
                                  " VALUES     (" & Param & ", '" & Avdata(0, I) & "')")
            Next I
        End If
     End If
End With
Set Rc = Nothing
Set Avdata = Nothing
End Sub

Private Function MaxForm() As Integer
Dim Rc As New DBQuick
Dim Avdata As Variant
Rc.DBOpen "SELECT     MAX([User ID]) AS MAXNOM FROM         [User Table]", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        Avdata = 0
     End If
     Avdata = Avdata + 1
     MaxForm = Avdata
End With
Set Rc = Nothing
Set Avdata = Nothing
End Function

Private Function MaxGroup() As Integer
Dim Rc As New DBQuick
Dim Avdata As Variant
Rc.DBOpen "SELECT     MAX(id) AS MAXNOM FROM  User_table_group", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        Avdata = 0
     End If
     Avdata = Avdata + 1
     MaxGroup = Avdata
End With
Set Rc = Nothing
Set Avdata = Nothing
End Function

Private Function MaxGroupLine() As Integer
Dim Rc As New DBQuick
Dim Avdata As Variant
Rc.DBOpen "SELECT     MAX(id) AS MAXNOM FROM  User_table_Line", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        Avdata = 0
     End If
     Avdata = Avdata + 1
     MaxGroupLine = Avdata
End With
Set Rc = Nothing
Set Avdata = Nothing
End Function

Private Sub EnabledControl(ByVal Tipical As Boolean)
txtBox(0).Enabled = False
txtBox(1).Enabled = Tipical
txtBox(2).Enabled = False
txtBox(3).Enabled = Tipical
DataCombo1(0).Enabled = Tipical
cmOK(4).Enabled = Tipical
cmOK(5).Enabled = Tipical
'cmOK(15).Enabled = Tipical
End Sub



Private Sub EnabledTombol(ByVal Tipical As Boolean)
CmdTombol(0).Enabled = Not Tipical
CmdTombol(1).Enabled = Not Tipical
CmdTombol(2).Enabled = Tipical
CmdTombol(3).Enabled = Tipical
CmdTombol(4).Enabled = Tipical
End Sub


'Private Sub OpenDetail(ByVal Param As Integer, ByVal GroupTable As String)
Private Sub OpenDetail(ByVal GroupTable As String, ByVal GroupForm As String)
Dim Avdata As Variant
Dim Fld As Field
Dim j, I As Integer
Dim sql As String

            
sql = "SELECT  user_table_line.id,user_table_line.[group id],user_table_line.[form id],User_Table_Form.[Form Name], User_table_line.Access, User_table_line.New, User_table_line.Edit, User_table_line.Del " & _
        "FROM  User_Table_Form INNER JOIN " & _
        "User_table_line ON User_Table_Form.ID = User_table_line.[Form ID] INNER JOIN " & _
        "User_table_group ON User_table_line.[Group ID] = User_table_group.id " & _
        "WHERE  (User_table_group.[Group Name] = N'" & GroupTable & "') AND (User_Table_Form.[Group Form] = N'" & GroupForm & "')"
            
RcDet.DBOpen sql, CNN, lckLockReadOnly

'With RcDet.DBRecordset
'     MSHFlexGrid1.Clear
'     MSHFlexGrid1.ClearStructure
'     MSHFlexGrid1.Cols = .Fields.Count + 1
'     MSHFlexGrid1.Rows = 2
'     MSHFlexGrid1.ColWidth(0) = 0
'    ' MSHFlexGrid1.ColWidth(7) = 0
'     MSHFlexGrid1.ColWidth(1) = 0
'     MSHFlexGrid1.ColWidth(2) = 0
'     MSHFlexGrid1.ColWidth(3) = 0
'     MSHFlexGrid1.ColWidth(4) = 2830
'     MSHFlexGrid1.ColWidth(5) = 750
'     MSHFlexGrid1.ColWidth(6) = 750
'     MSHFlexGrid1.ColWidth(7) = 750
     'MSHFlexGrid1.ColWidth(6) = 0
  '   If .Recordcount > 0 Then
       ' MSHFlexGrid1.Rows = .Recordcount + 1
'        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
'        For j = 0 To UBound(Avdata, 2)
'            With MSHFlexGrid1
'                 I = 1
'                 For Each Fld In RcDet.DBRecordset.Fields
'                     .TextMatrix(0, I) = Fld.Name  'nampilkan header kolom
'                     .TextMatrix(j + 1, I) = Avdata(I - 1, j) ' nampilkan name form
'                     Select Case Fld.Type
'                            Case 11:
'                            .col = I
'                            .row = j + 1
'                            .TextMatrix(j + 1, I) = ""
'                            '.TextMatrix(j + 1, I) = Avdata(I - 1, j)  'digunakan nampilkan edit,true,false
''                            Set .CellPicture = LoadResPicture(IIf(CBool(Avdata(I - 1, j)), 104, 103), vbResBitmap)
'                            ' .Text = IIf(True, strChecked, strUnChecked)
'                            .CellFontName = "Wingdings"
'                            .CellFontSize = 14
'                            .CellAlignment = flexAlignCenterCenter
'                            .Text = IIf(Avdata(I - 1, j) = True, "þ", "q")
'                            .CellPictureAlignment = flexAlignCenterCenter
'                     End Select
'                     I = I + 1
'                 Next
'            End With
'        Next
'        MSHFlexGrid1.Redraw = True
'        For I = 0 To List1.ListCount - 1
'          '  List1.Selected(I) = BoolToInt(OldData(RcLaporan.DBRecordset.Fields(0), List1.List(I)))
'        Next I
'      End If
'
'End With
'Set Avdata = Nothing
End Sub

Private Sub OpenLaporan(ByVal Param As Integer)
Dim Avdata As Variant
Dim Fld As Field
Dim j, I As Integer
'RcPermit.DBOpen "SELECT     [Report Modules].IDReport AS [ID_Report], [Report Modules].AliasReport AS [Report Name], [Report Modules].FileNameReport AS [File Name],  [Report Permit].Laporan AS Permitted FROM         [Report Permit] RIGHT OUTER JOIN [Report Modules] ON [Report Permit].IDReport = [Report Modules].IDReport WHERE     ([Report Permit].[User ID] = " & Param & " OR [Report Permit].[User ID] IS NULL) AND ([Report Permit].Laporan IS NULL) ORDER BY [Report Modules].FileNameReport", CNN, lckLockReadOnly
'With RcPermit.DBRecordset
'     If .Recordcount <> 0 Then
'        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
'        For j = 0 To UBound(Avdata, 2)
'            SendDataToServer " INSERT INTO [Report Permit]" & _
'                             " ([User ID], IDReport, Laporan)" & _
'                             " VALUES (" & Param & ", N'" & Avdata(0, j) & "', 0)"
'        Next j
'     End If
'     .Close
'End With
'RcPermit.DBOpen "SELECT     [Report Permit].IDReport AS [ID Report], [Report Modules].AliasReport AS [Report Name], [Report Modules].FileNameReport AS [File Name], [Report Permit].Laporan AS Permitted FROM         [Report Permit] INNER JOIN [Report Modules] ON [Report Permit].IDReport = [Report Modules].IDReport WHERE     ([Report Permit].[User ID] = " & Param & ") ORDER BY [Report Modules].AliasReport", CNN, lckLockBatch
'
'With RcPermit.DBRecordset
'     DgLaporan.Clear
'     DgLaporan.ClearStructure
'     DgLaporan.Cols = .Fields.Count + 1
'     DgLaporan.Rows = 2
'     DgLaporan.ColWidth(0) = 0
'     DgLaporan.ColWidth(1) = 0
'     DgLaporan.ColWidth(2) = 3400
'     DgLaporan.ColWidth(3) = 3400
'     DgLaporan.ColWidth(4) = 1200
'
'     If .Recordcount > 0 Then
'        DgLaporan.Rows = .Recordcount + 1
'        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
'        For j = 0 To UBound(Avdata, 2)
'            With DgLaporan
'                 I = 1
'                 For Each Fld In RcPermit.DBRecordset.Fields
'                     .TextMatrix(0, I) = Fld.Name
'                     .TextMatrix(j + 1, I) = Avdata(I - 1, j)
'                     Select Case Fld.Type
'                            Case 11:
'                            .Col = I
'                            .Row = j + 1
'                            .TextMatrix(j + 1, I) = ""
'                            Set .CellPicture = LoadResPicture(IIf(CBool(Avdata(I - 1, j)), 104, 103), vbResBitmap)
'                            .CellPictureAlignment = flexAlignCenterCenter
'                     End Select
'                     I = I + 1
'                 Next
'            End With
'        Next
'        DgLaporan.Redraw = True
''        For i = 0 To List1.ListCount - 1
''            List1.Selected(i) = BoolToInt(OldData(RcLaporan.DBRecordset.Fields(0), List1.List(i)))
''        Next i
'      End If
'
'End With
'Set Avdata = Nothing
End Sub

Private Function OldData(ByVal Param As Integer, ByVal GroupTable As String) As Boolean
Dim Rc As New DBQuick
Rc.DBOpen "SELECT  [Detail User Table].Laporan FROM [Detail User Table] INNER JOIN [Form Table] ON [Detail User Table].Idx = [Form Table].Idx WHERE     ([Detail User Table].[User ID] = " & Param & ") AND ([Form Table].[Group Table] = N'" & GroupTable & "') GROUP BY [Detail User Table].Laporan", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        OldData = IIf(Not IsNull(.Fields(0)), .Fields(0), False)
     Else
        OldData = False
     End If
End With
Set Rc = Nothing
End Function

Private Sub txtBox_Change(Index As Integer)
Dim rcFind As New DBQuick
If Index = 4 Then
    rcFind.DBOpen "select id , [group name] from user_table_group where id= '" & txtBox(4).Text & "'", CNN, lckLockReadOnly
    If rcFind.DBRecordset.Recordcount > 0 Then
        DataCombo1(0).Text = rcFind.DBRecordset.Fields(1)
    End If
ElseIf Index = 6 Then
'    'rcFind.DBOpen "select id , [form name],[group form] from user_table_form where [form name]= '" & txtBox(6).Text & "'", CNN, lckLockReadOnly
'    rcFind.DBOpen "select id , [form name],[group form] from user_table_form where [form name]= '" & txtBox(6).Text & "' AND [group form]='" & txttempgroupform.Text & "', CNN, lckLockReadOnly"
'    If rcFind.DBRecordset.Recordcount > 0 Then
'        txtTempform = rcFind.DBRecordset.Fields(0)
'    End If
    'sql = "select id , [form name],[group form] from user_table_form where [form name]= '" & txtBox(6).Text & "' AND [group form]='" & txttempgroupform.Text & "'"
    rcFind.DBOpen "select id , [form name],[group form] from user_table_form where [form name]= '" & txtBox(6).Text & "' AND [group form]='" & txttempgroupform.Text & "'", CNN, adOpenStatic, adLockReadOnly
    If rcFind.DBRecordset.Recordcount > 0 Then
        txtTempform = rcFind.DBRecordset.Fields(0)
    End If
End If

End Sub

Private Sub txtBox_GotFocus(Index As Integer)

Block txtBox(Index)

    
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Public Sub loadcombo()

End Sub


Private Sub GridLayout()
   Dim FlipFlop As Boolean
   Dim x As Integer
   Dim Y As Integer
   
'   With MSHFlexGrid1
'      .Rows = rsUsing.Recordcount + 1
'      .TextMatrix(0, 0) = "Triggered"
'      .TextMatrix(0, 1) = "Equipment No"
'
''      If rsUsing.Recordcount > 0 Then
''         rsUsing.MoveFirst
''         FlipFlop = False
''         For x = 1 To rsUsing.Recordcount
''               For y = 0 To 27
''                  .Col = y
''                  .Row = x
''                  If FlipFlop Then .CellBackColor = &HEEDAC1
''                  If rsUsing.Fields(0) = True Then .CellBackColor = RGB(252, 180, 180)
''                  If Not rsUsing.EOF Then
''                     If y = 0 Then
''                        .CellFontName = "Wingdings"
''                        .CellFontSize = 14
''                        .CellAlignment = flexAlignCenterCenter
'                        .Text = IIf(rsUsing.Fields(Y) = True, strChecked, strUnChecked)
''                     Else
''                        .Text = IIf(IsNull(rsUsing.Fields(y).value), "", rsUsing.Fields(y).value)
''                     End If
''
''                  End If
''               Next
''            rsUsing.MoveNext
''            FlipFlop = IIf(FlipFlop, False, True)
''         Next
''      End If
  ' End With
End Sub

Public Sub ViewGroupLine()
Dim sql As String
Set RcViewGroup = New Recordset
RcViewGroup.CursorLocation = adUseClient

sql = "SELECT  User_table_line.ID, User_table_group.[Group Name], User_Table_Form.[Group Form], User_Table_Form.[Form Name]," & _
        " User_table_line.Access , User_table_line.New, User_table_line.Edit, User_table_line.Del " & _
        "FROM  User_table_group INNER JOIN" & _
        " User_table_line ON User_table_group.id = User_table_line.[Group ID] INNER JOIN " & _
        " User_Table_Form ON User_table_line.[Form ID] = User_Table_Form.ID order by user_table_group.[group Name]"


RcViewGroup.Open sql, CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set RcViewGroup.ActiveConnection = Nothing
With RcViewGroup
     Set DGGrid.DataSource = RcViewGroup
End With
End Sub

Private Sub ViewListCntrol()
ListAccess.AddItem "True"
ListAccess.AddItem "False"

ListNew.AddItem "True"
ListNew.AddItem "False"

ListEdit.AddItem "True"
ListEdit.AddItem "False"

Listdel.AddItem "True"
Listdel.AddItem "False"
End Sub

Private Sub Clear()
DataCombo2(1).Text = ""
txtBox(6).Text = ""

ListAccess.Selected(0) = False
ListAccess.Selected(1) = False

ListNew.Selected(0) = False
ListNew.Selected(1) = False

ListEdit.Selected(0) = False
ListEdit.Selected(1) = False

Listdel.Selected(0) = False
Listdel.Selected(1) = False
End Sub

Private Sub aktifText()
DataCombo2(1).Enabled = True
cmOK(11).Enabled = True
ListAccess.Enabled = True
ListNew.Enabled = True
ListEdit.Enabled = True
Listdel.Enabled = True
End Sub

Private Sub NoaktifText()
DataCombo2(1).Enabled = False
cmOK(11).Enabled = False
ListAccess.Enabled = False
ListNew.Enabled = False
ListEdit.Enabled = False
Listdel.Enabled = False
End Sub

Private Sub txtcari1_Change()
On Error Resume Next
Dim strcari As String
If txtcari1 <> "" Then
    If Not RcViewGroup Is Nothing Then
       If RcViewGroup.Recordcount <> 0 Then
         Select Case RcViewGroup.Fields(DGGrid.Columns(mLoop1).DataField)
            Case adBigInt, adInteger, adCurrency, adDecimal, adDouble, adNumeric, adSingle, adSmallInt, adTinyInt, adVarNumeric
               strcari = "[" & DGGrid.Columns(mLoop1).DataField & "]" & " = " & txtcari1
            Case Else
               strcari = "[" & DGGrid.Columns(mLoop1).DataField & "]" & " Like '" & txtcari1 & "%'"
         End Select
          RcViewGroup.Filter = strcari ', 0, adSearchForward, adBookmarkFirst
          If RcViewGroup.Recordcount = 0 Then MessageBox "Kriteria Yang Dicari Tidak Ada..............!", vbCritical
       Else
          ViewGroupLine
       End If
    End If
Else
     ViewGroupLine
End If
End Sub

Private Sub ViewGroup()
Dim sql As String
Set RcGroupview = New Recordset
RcGroupview.CursorLocation = adUseClient
sql = "SELECT id as ID , [group name] as [Nama Group] from user_table_group order by ID"
RcGroupview.Open sql, CNN, adOpenKeyset, adLockOptimistic, adCmdText
Set RcGroupview.ActiveConnection = Nothing
With RcGroupview
     Set DataGroup.DataSource = RcGroupview
End With
Set txtBox(5).DataSource = RcGroupview
End Sub

Private Sub OpenDetailGrid(ByVal GroupTable As String, ByVal GroupForm As String)
Dim sql As String
            
sql = "SELECT  user_table_line.id,user_table_line.[group id],user_table_line.[form id],User_Table_Form.[Form Name], User_table_line.Access, User_table_line.New, User_table_line.Edit, User_table_line.Del " & _
        "FROM  User_Table_Form INNER JOIN " & _
        "User_table_line ON User_Table_Form.ID = User_table_line.[Form ID] INNER JOIN " & _
        "User_table_group ON User_table_line.[Group ID] = User_table_group.id " & _
        "WHERE  (User_table_group.[Group Name] = N'" & GroupTable & "') AND (User_Table_Form.[Group Form] = N'" & GroupForm & "')"
     
RcDet.DBOpen sql, CNN, lckLockBatch
Set TDBGrid1.DataSource = RcDet.DBRecordset

TDBGrid1.Columns(0).Visible = False 'hilangkan no ID
TDBGrid1.Columns(1).Locked = True 'Look form name

TDBGrid1.Columns(2).Alignment = dbgCenter
TDBGrid1.Columns(3).Alignment = dbgCenter
TDBGrid1.Columns(4).Alignment = dbgCenter
TDBGrid1.Columns(5).Alignment = dbgCenter

End Sub


Private Sub OpenDetailReport(ByVal Param As Integer)
Dim sql As String
            
sql = "SELECT  dbo.[report permit].idx,dbo.[report permit].[User ID], dbo.[report permit].noidx, dbo.[report modules].Description, dbo.[report modules].[Alias Report]," & _
     "dbo.[report modules].ReportGroup , dbo.[report modules].FileNameReport, dbo.[report modules].ViewObject, dbo.[report permit].Laporan " & _
     "FROM dbo.[report permit] INNER JOIN " & _
     "dbo.[report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx " & _
     " WHERE  (dbo.[report permit].[User ID] =" & Param & " )" & _
     " ORDER BY ReportGroup"
            
RcDetReport.DBOpen sql, CNN, lckLockBatch
'Set TDBGrid2.DataSource = RcDetReport.DBRecordset

'TDBGrid2.Columns(0).Visible = False 'hilangkan no idx
'TDBGrid2.Columns(1).Visible = False 'hilangkan no User ID
'TDBGrid1.Columns(1).Locked = True 'Look form name
'
'TDBGrid1.Columns(2).Alignment = dbgCenter
'TDBGrid1.Columns(3).Alignment = dbgCenter
'TDBGrid1.Columns(4).Alignment = dbgCenter
'TDBGrid1.Columns(5).Alignment = dbgCenter

End Sub


