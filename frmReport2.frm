VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmReport2 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Administration"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmReport2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9270
   Tag             =   "Administrasi Laporan"
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9270
      TabIndex        =   10
      Top             =   6615
      Width           =   9270
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   1560
         Picture         =   "frmReport2.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   840
         Picture         =   "frmReport2.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   3000
         Picture         =   "frmReport2.frx":138F6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   100
         Width           =   720
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   7935
         Top             =   195
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   3720
         Picture         =   "frmReport2.frx":1A148
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   2280
         Picture         =   "frmReport2.frx":1BC42
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   120
         Picture         =   "frmReport2.frx":22494
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   100
         Width           =   720
      End
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         Height          =   30
         Left            =   -45
         TabIndex        =   11
         Top             =   0
         Width           =   9390
      End
   End
   Begin TabDlg.SSTab TabReport 
      Height          =   6360
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Report"
      TabPicture(0)   =   "frmReport2.frx":28CE6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GridReport"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Design"
      TabPicture(1)   =   "frmReport2.frx":28D02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(2)"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid GridReport 
         Height          =   2700
         Left            =   150
         TabIndex        =   9
         Tag             =   "Design"
         ToolTipText     =   "Double Click to preview report"
         Top             =   3465
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   4763
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin VB.Frame Frame1 
         Caption         =   " Module "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   480
         Width           =   2760
         Begin VB.OptionButton chkOpt 
            Caption         =   "Purchasing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   300
            TabIndex        =   17
            Tag             =   "Purchasing"
            ToolTipText     =   "2"
            Top             =   1845
            Width           =   2325
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Mechanical Engineering"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   300
            TabIndex        =   16
            Tag             =   "Mechanical Engineering"
            Top             =   1221
            Width           =   2325
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Quality Control"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   300
            TabIndex        =   8
            Tag             =   "Quality Control"
            ToolTipText     =   "3"
            Top             =   2157
            Value           =   -1  'True
            Width           =   2325
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Logistic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   7
            Tag             =   "Logistic"
            Top             =   597
            Width           =   2325
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Marketing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   300
            TabIndex        =   6
            Tag             =   "Marketing"
            Top             =   909
            Width           =   2325
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Produksi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   300
            TabIndex        =   5
            Tag             =   "Produksi"
            ToolTipText     =   "1"
            Top             =   1533
            Width           =   2325
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Human Resources"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   4
            Tag             =   "Human Resources"
            Top             =   285
            Width           =   2325
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5670
         Index           =   2
         Left            =   -74880
         TabIndex        =   2
         Top             =   330
         Width           =   8670
         Begin VB.CommandButton cmdLink 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4080
            Picture         =   "frmReport2.frx":28D1E
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   360
            Width           =   345
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
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
            Height          =   345
            Index           =   0
            Left            =   1320
            TabIndex        =   40
            Tag             =   "ana"
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
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
            Height          =   345
            Index           =   1
            Left            =   1320
            TabIndex        =   39
            Tag             =   "ana"
            Top             =   855
            Width           =   2775
         End
         Begin VB.CommandButton cmdView 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4080
            Picture         =   "frmReport2.frx":2F570
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1800
            Width           =   345
         End
         Begin MSDataListLib.DataCombo CmbModule 
            DataField       =   "ReportGroup"
            Height          =   315
            Left            =   1320
            TabIndex        =   41
            Top             =   1335
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
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
         Begin MSDataGridLib.DataGrid DgDesign 
            Height          =   3285
            Left            =   165
            TabIndex        =   42
            Tag             =   "Design"
            Top             =   2265
            Width           =   8325
            _ExtentX        =   14684
            _ExtentY        =   5794
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BorderStyle     =   0
            HeadLines       =   2
            RowHeight       =   15
            RowDividerStyle =   6
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Module"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   47
            Top             =   1320
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   46
            Top             =   405
            Width           =   945
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   45
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View Object"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   44
            Top             =   1800
            Width           =   855
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1320
            X2              =   120
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   1320
            X2              =   120
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   1320
            X2              =   120
            Y1              =   1635
            Y2              =   1635
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   1320
            X2              =   120
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Label LblView 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataSource      =   "DataTrans"
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
            Height          =   345
            Left            =   1320
            TabIndex        =   43
            Tag             =   "BAHAN"
            Top             =   1800
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Filter Kriteria "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Left            =   2985
         TabIndex        =   1
         Top             =   480
         Width           =   5850
         Begin VB.PictureBox PictMain 
            BackColor       =   &H80000001&
            Height          =   2220
            Left            =   480
            ScaleHeight     =   2160
            ScaleWidth      =   5145
            TabIndex        =   21
            Top             =   840
            Visible         =   0   'False
            Width           =   5205
            Begin VB.PictureBox PictFilter 
               BackColor       =   &H00EAAF6F&
               BorderStyle     =   0  'None
               Height          =   2025
               Left            =   75
               ScaleHeight     =   2025
               ScaleWidth      =   4980
               TabIndex        =   22
               Top             =   45
               Width           =   4980
               Begin VB.Frame Frame1 
                  BackColor       =   &H00C0FFFF&
                  Height          =   30
                  Index           =   1
                  Left            =   -30
                  TabIndex        =   32
                  Top             =   1470
                  Width           =   4995
               End
               Begin VB.ComboBox CmbOperator 
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
                  ItemData        =   "frmReport2.frx":35DC2
                  Left            =   135
                  List            =   "frmReport2.frx":35DE5
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   480
                  Width           =   2895
               End
               Begin VB.CommandButton cmdCancel 
                  Caption         =   "&Cancel"
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
                  Index           =   1
                  Left            =   1155
                  TabIndex        =   30
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.CommandButton cmdOK 
                  Caption         =   "&OK"
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
                  Index           =   0
                  Left            =   135
                  TabIndex        =   29
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.ComboBox CmbFilter 
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
                  Index           =   0
                  Left            =   105
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   825
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.ComboBox CmbFilter 
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
                  Index           =   1
                  Left            =   105
                  Style           =   2  'Dropdown List
                  TabIndex        =   27
                  Top             =   1065
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.TextBox TxtFilter 
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
                  Index           =   1
                  Left            =   2370
                  TabIndex        =   26
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.TextBox TxtFilter 
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
                  Index           =   0
                  Left            =   2400
                  TabIndex        =   25
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H80000001&
                  BorderStyle     =   0  'None
                  FillColor       =   &H80000002&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   345
                  Left            =   0
                  ScaleHeight     =   345
                  ScaleWidth      =   4980
                  TabIndex        =   23
                  Top             =   0
                  Width           =   4980
                  Begin VB.Label LBLFilter 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "DATA SELECTION"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000005&
                     Height          =   195
                     Left            =   30
                     TabIndex        =   24
                     Top             =   45
                     Width           =   1395
                  End
               End
               Begin MSComCtl2.DTPicker DTPickFilter 
                  Height          =   315
                  Index           =   1
                  Left            =   3060
                  TabIndex        =   33
                  Top             =   1050
                  Visible         =   0   'False
                  Width           =   1815
                  _ExtentX        =   3201
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
                  CustomFormat    =   "dd MMM yyyy"
                  Format          =   20381699
                  CurrentDate     =   36877
               End
               Begin MSComCtl2.DTPicker DTPickFilter 
                  Height          =   315
                  Index           =   0
                  Left            =   3060
                  TabIndex        =   34
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1815
                  _ExtentX        =   3201
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
                  CustomFormat    =   "dd MMM yyyy"
                  Format          =   20381699
                  CurrentDate     =   36877
               End
               Begin VB.Label lblTo 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "To"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   3270
                  TabIndex        =   36
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   225
               End
               Begin VB.Label lblAnd 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "And"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   3750
                  TabIndex        =   35
                  Top             =   825
                  Visible         =   0   'False
                  Width           =   345
               End
            End
         End
         Begin MSComctlLib.ListView ListFilter 
            Height          =   2475
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   4366
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Filter By"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "From"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "To"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "To"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "FieldType"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Operator"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.Label lblReportIndex 
      Caption         =   "0"
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim myPart As New Utility

Private Rc As New DBQuick

Private WithEvents rsReport As ADODB.Recordset
Attribute rsReport.VB_VarHelpID = -1

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private RcFilter As New DBQuick

Private RcFlt As New DBQuick

Private Midex As String

Private IdxOpt As Integer
Dim mVarTmp As String

Private StrLaporan As String
Dim strSQL As String

Private RcProses As New DBQuick

Private Enum TypeFld
  fldTanggal = 0
  fldString = 1
  fldNumeric = 2
End Enum

Private Enum FieldControlType
  oTextBx
  oComboBx
  oMaskBx
  oDTPicker
End Enum

Const ObjLeft = 3060
Const ObjTop1 = 480
Const ObjTop2 = 1050
Dim ReportId As String
Dim obj As Object

Private TipeFld As TypeFld
Dim sViewObject As String
Dim rowN As Integer

Private Sub ActivateObject(sObj As FieldControlType, _
                           bStatus As Boolean)
  Dim I As Integer

  '    oTextBx
  '    oComboBx
  '    oMaskBx
  '    fcCheckBx
  '    oDTPicker
  '
  sViewObject = GridReport.Columns(6).Text

  For I = 0 To 1
    TxtFilter(I).Visible = False
    TxtFilter(I).Text = ""
    DTPickFilter(I).Visible = False
    CmbFilter(I).Visible = False
  Next

  Select Case sObj

    Case FieldControlType.oTextBx
      TxtFilter(0).Visible = True
      TxtFilter(1).Visible = bStatus
      TxtFilter(0).Left = ObjLeft
      TxtFilter(1).Left = ObjLeft
      TxtFilter(0).Top = ObjTop1
      TxtFilter(1).Top = ObjTop2

      'For I = 0 To 1
      '    TxtFilter(0).Visible = bStatus
      '    TxtFilter(1).Visible = Not bStatus
      '    DTPickFilter(I).Visible = Not bStatus
      '    CmbFilter(I).Visible = Not bStatus
      'Next
    Case FieldControlType.oComboBx
      Dim rsCombo As Recordset
      CmbFilter(0).Visible = True
      CmbFilter(1).Visible = bStatus
      CmbFilter(0).Left = ObjLeft
      CmbFilter(1).Left = ObjLeft
      CmbFilter(0).Top = ObjTop1
      CmbFilter(1).Top = ObjTop2
        
      Set rsCombo = New Recordset
        
      CmbFilter(0).Clear
      CmbFilter(1).Clear

      If Trim(sViewObject) = "" Then
        MsgBox "Object Table Not Defined"
      Else
        strSQL = "select Distinct [" & Trim(ListFilter.SelectedItem.Text) & "] from [" & sViewObject & "] where [" & _
                ListFilter.SelectedItem.Text & "] is not null "
        Debug.Print strSQL
        rsCombo.Open strSQL, CNN, adOpenStatic, adLockReadOnly

        If rsCombo.Recordcount > 0 Then
          While Not rsCombo.EOF
            CmbFilter(0).AddItem Trim(rsCombo.Fields(0).Value)
            CmbFilter(1).AddItem Trim(rsCombo.Fields(0).Value)
            rsCombo.MoveNext
          Wend
        End If
      End If

      'For I = 0 To 1
      '    TxtFilter(I).Visible = Not bStatus
      '    DTPickFilter(I).Visible = Not bStatus
      '    CmbFilter(0).Visible = bStatus
      '    CmbFilter(1).Visible = Not bStatus
      'Next
    Case FieldControlType.oMaskBx

    Case FieldControlType.oDTPicker
      DTPickFilter(0).Visible = True
      DTPickFilter(1).Visible = bStatus
      DTPickFilter(0).Left = ObjLeft
      DTPickFilter(1).Left = ObjLeft
      DTPickFilter(0).Top = ObjTop1
      DTPickFilter(1).Top = ObjTop2
      'For I = 0 To 1
      '    TxtFilter(I).Visible = Not bStatus
      '    DTPickFilter(0).Visible = bStatus
      '    DTPickFilter(0).Value = Date
      '    DTPickFilter(1).Visible = Not bStatus
      '            DTPickFilter(1).Value = Date
      '    CmbFilter(I).Visible = Not bStatus
      'Next
  End Select

End Sub

'Private Sub ComboMove()
'  On Error Resume Next
'
'  With GridFilter
'
'    If .ApproxCount <> 0 Then
'      cboFilter.Visible = True
'      cboFilter.Move .Columns(1).Left + 100, (.RowTop(.Row) + .RowHeight) + 15, .Columns(1).Width, .RowHeight
'      '        GridFilter.SetFocus
'      OpenFlt
'    End If
'
'  End With
'
'  Err.Clear
'End Sub

Private Sub ExecSave()
  Dim obj As Object

  For Each obj In Me

    If TypeOf obj Is OptionButton Then
      If obj.Value = True Then ReportId = obj.ToolTipText
    End If

  Next obj

  '   Select Case rsReport.Status
  '
  '      Case adRecModified
  '         strSQL = "UPDATE [Report Modules] Set Description = N'" & txtBox(1).Text & "', " & " ReportGroup = N'" & CmbModule.BoundText & "', FileNameReport = N'" & txtBox(0).Text & "', ViewObject = N'" & CmbView.BoundText & "'" & " Where (NoIdx = " & rsReport.Fields("NoIdx").Value & ")"
  '
  '      Case adRecNew
  strSQL = " INSERT INTO [Report Modules] (Description, ReportGroup, FileNameReport, ViewObject,REPORT_ID,[Alias Report]) " & _
          " VALUES (N'" & txtBox(1).Text & "', N'" & CmbModule.BoundText & "', N'" & txtBox(0).Text & "', N'" & LblView & _
          "', N'" & ReportId & "','" & Mid(txtBox(0), 1, Len(txtBox(0)) - 4) & "')"
  '
  '      Case Else
  '         strSQL = ""
  '   End Select

  SendDataToServer strSQL
  '   If Len(strSQL) <> 0 Then
  '      If myPart.SendCommandToServer(strSQL) Then
  '         MsgBox "Data Saved Successfully...", vbInformation, App.ProductName
  '      Else
  '         MsgBox "Save data error...", vbCritical, App.ProductName
  '      End If
  '   End If

End Sub

Private Sub GenerateFld()
  Dim I As Integer
  Dim vLst As ListItem
  Dim rsView As ADODB.Recordset
  Dim ReportIndex As Integer

  For Each obj In Me

    If TypeOf obj Is OptionButton Then
      If obj.Value = True Then ReportId = obj.ToolTipText
    End If

  Next obj
   
  If Not rsReport Is Nothing Then
    If Not rsReport.EOF Then
      If Not IsNull(rsReport.Fields("ViewObject").Value) Then
        sViewObject = rsReport.Fields("ViewObject").Value
      Else
        Exit Sub
      End If

    Else
      sViewObject = ""
    End If

  Else
    Exit Sub
  End If

  If Rc.DBOpen( _
          "SELECT Tools_Filter.REPORT_ID, Tools_Filter.FIELD_NAME, Tools_Filter.FIELD_TYPE, Tools_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Tools_Filter ON ([Report Modules].NoIdx = Tools_Filter.REPORT_ID) Where [Report Modules].NoIdx = '" _
          & CInt(lblReportIndex) & "' AND [Report Modules].ViewObject = '" & sViewObject & "'", CNN, lckLockBatch) = True Then

    If Rc.Recordcount <> 0 Then
      ListFilter.ListItems.Clear

      With Rc.DBRecordset

        Do While Not .EOF
          Set vLst = ListFilter.ListItems.Add(, , UCase(.Fields("FIELD_NAME").Value))
          vLst.SubItems(3) = UCase(.Fields("FIELD_NAME").Value)
          vLst.SubItems(3) = IIf(IsNull(.Fields("FIELD_TYPE").Value), "", UCase(.Fields("FIELD_TYPE").Value))
          .MoveNext
        Loop

      End With

    Else
      ListFilter.ListItems.Clear
    End If
   
    '////////////////////////////////////////////////////////////////////

  End If

End Sub

Private Function GetOperator(nKode As Integer) As String
  On Error Resume Next
  '0   is equal to (=)
  '1   is not equal to (<>)
  '2   is less than (<)
  '3   is less than and equal to (<=)
  '4   is greater than (>)
  '5   is greater than and equal to (>=)
  '6   is between
  '7   is not between
  '8   is like
  '9   is not like

  Select Case nKode

    Case 0
      GetOperator = "="

    Case 1
      GetOperator = "<>"

    Case 2
      GetOperator = "<"

    Case 3
      GetOperator = "<="

    Case 4
      GetOperator = ">"

    Case 5
      GetOperator = ">="

    Case 6
      GetOperator = "BETWEEN"

    Case 7
      GetOperator = "NOT BETWEEN"

    Case 8
      GetOperator = "LIKE"

    Case 9
      GetOperator = "NOT LIKE"
  End Select

End Function

Private Sub GridLayout()
  GridReport.Columns(0).Visible = False
  GridReport.Columns(1).Width = 5000
  DgDesign.Columns(0).Width = 1964.976
  DgDesign.Columns(1).Width = 1964.976
  'DgDesign.Columns(2).Width = 3834.835
  'DgDesign.Height = 3090
  GridReport.HoldFields
End Sub

Private Sub indexOption()

  For Each obj In Me

    If TypeOf obj Is OptionButton Then
      If obj.Value = True Then ReportId = obj.ToolTipText
    End If

  Next obj

End Sub

Private Function LocalPeriodeActive() As Integer

  Select Case mVarTempPeriode

    Case 1
      LocalPeriodeActive = 12

    Case 2
      LocalPeriodeActive = 1

    Case 3
      LocalPeriodeActive = 2

    Case 4
      LocalPeriodeActive = 3

    Case 5
      LocalPeriodeActive = 4

    Case 6
      LocalPeriodeActive = 5

    Case 7
      LocalPeriodeActive = 6

    Case 8
      LocalPeriodeActive = 7

    Case 9
      LocalPeriodeActive = 8

    Case 10
      LocalPeriodeActive = 9

    Case 11
      LocalPeriodeActive = 10

    Case 12
      LocalPeriodeActive = 1
  End Select

End Function

Private Sub OpenDB(ByVal Nodetext As String)
  Dim I As Long

  strSQL = "SELECT * FROM [Report Modules] WHERE (ReportGroup = N'" & Nodetext & "')"
  Set rsReport = New ADODB.Recordset
  rsReport.CursorLocation = adUseClient
  rsReport.Open strSQL, CNN, adOpenKeyset, adLockOptimistic
  Set GridReport.DataSource = rsReport
  GridReport.ReBind
  Set CmbModule.DataSource = rsReport

  If rsReport.Recordcount < 1 Then
    GridReport.Tag = "0"
    Exit Sub
  Else
    GridReport.Tag = "1"
  End If

  lblReportIndex.Caption = GridReport.Columns(0).Text

  For I = 0 To 1
    Set txtBox(I).DataSource = rsReport
  Next

  Dim rsDisain As New ADODB.Recordset
  strSQL = "SELECT * FROM [REPORT MODULES] WHERE REPORTGROUP = '" & CmbModule.BoundText & "'"
  rsDisain.CursorLocation = adUseClient
  rsDisain.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
  Set DgDesign.DataSource = rsDisain

  '    m_data = Nothing
  '        m_data = New SqlClient.SqlDataAdapter(StrSql, Cnn)
  '    If Not IsNothing(m_set) Then m_set = Nothing
  '    m_set = New DataSet
  '    m_data.Fill (m_set)
  '        GrdPreview.TableStyles.Clear()
  '        GrdPreview.SetDataBinding(m_set, m_set.Tables(0).ToString)
  '        GridLayout(m_set, GrdPreview, Nodetext)
  '        FilterCriteria()

End Sub

'Private Sub OpenFlt()
'
'  If Not rsReport Is Nothing Then
'    If rsReport.State = 1 Then
'      If rsReport.Recordcount <> 0 And Len(sViewObject) <> 0 Then
'        RcFlt.DBOpen " select [" & RcFilter.Fields(0) & "] from [" & sViewObject & "] group By [" & RcFilter.Fields(0) & _
'                "] order by [" & RcFilter.Fields(0) & "]", CNN, lckLockReadOnly
'        cboFilter.Text = ""
'        cboFilter.ListField = RcFilter.Fields(0)
'        Set cboFilter.RowSource = RcFlt.DBRecordset
'      Else
'        cboFilter.Text = ""
'        Set cboFilter.RowSource = Nothing
'      End If
'    End If
'  End If
'
'End Sub

Private Sub Preview()
  On Error GoTo Hell
  Dim RcTes As New DBQuick

  If cekListKosong = False Then
    strSQL = " SELECT * FROM [" & rsReport.Fields("ViewObject").Value & "]" & ScanFilter2
    strSQL = strSQL & mVarTmp
  Else
    strSQL = " SELECT * FROM [" & rsReport.Fields("ViewObject").Value & "]"
    
  End If
  
  RcTes.DBOpen strSQL, CNN ', lckLockBatch, lckLockSync
  ReportPos = PathRPT

  If RcTes.Recordcount <> 0 Then
    myPart.CallReportView strSQL, rsReport.Fields("FileNameReport").Value, ReportPos, rsReport.Fields("Description").Value 'App.Path & "\" & "Report"
  Else
    MsgBox "Laporan Belum Ada Datanya. Harap Diperiksa Filter Kriterianya", "Peringatan", msgOkOnly
  End If

  RcTes.CloseDB
  Exit Sub
Hell:


  MsgBox Err.Description, vbCritical, App.ProductName
  Err.Clear
End Sub

Private Function cekListKosong() As Boolean
  Dim ncount As Integer

  For ncount = 1 To ListFilter.ListItems.Count

    If ListFilter.ListItems(ncount).Checked = True Then cekListKosong = False
  Next

End Function

Private Function ScanFilter2() As String
  Dim ListCount As Byte
  mVarTmp = ""

  For ListCount = 1 To ListFilter.ListItems.Count
    ListFilter.ListItems(ListCount).Selected = IIf(ListFilter.ListItems(ListCount).Checked, "True", "False")

    If ListFilter.ListItems(ListCount).Checked And ListFilter.SelectedItem.ListSubItems(4).Text = "202" And Trim( _
            ListFilter.SelectedItem.ListSubItems(1).Text) <> "" And Trim(ListFilter.SelectedItem.ListSubItems(2).Text) <> "" _
            Then
      mVarTmp = mVarTmp & ListFilter.ListItems(ListCount).Text & " " & ListFilter.SelectedItem.ListSubItems(5).Text & " '" & _
              ListFilter.ListItems(1).SubItems(ListCount) & "' AND '" & ListFilter.ListItems(1).SubItems(ListCount + 1) & _
              "' And "
    ElseIf ListFilter.ListItems(ListCount).Checked And ListFilter.SelectedItem.ListSubItems(4).Text = "0" And Trim( _
            ListFilter.SelectedItem.ListSubItems(1).Text) <> "" And Trim(ListFilter.SelectedItem.ListSubItems(2).Text) <> "" _
            Then
      mVarTmp = mVarTmp & ListFilter.ListItems(ListCount).Text & " BETWEEN '" & Format(ListFilter.SelectedItem.ListSubItems( _
              1).Text, "yyyy-mm-dd") & " ' AND '" & Format(ListFilter.SelectedItem.ListSubItems(2).Text, "yyyy-mm-dd") & _
              "' And "
    ElseIf ListFilter.ListItems(ListCount).Checked And ListFilter.SelectedItem.ListSubItems(4).Text = "202" And Trim( _
            ListFilter.SelectedItem.ListSubItems(1).Text) <> "" And Trim(ListFilter.SelectedItem.ListSubItems(2).Text) = "" Then
      mVarTmp = mVarTmp & ListFilter.ListItems(ListCount).Text & " " & ListFilter.SelectedItem.ListSubItems(5).Text & " '" & _
              ListFilter.SelectedItem.ListSubItems(1).Text & "' AND'"
    End If

  Next

  mVarTmp = " WHERE " & Mid(mVarTmp, 1, Len(mVarTmp) - 4)
End Function

Private Function ScanFilter() As String
  Dim mVarI As Integer
  Dim mVarTmp As String
  Dim RcFlt As New DBQuick
  ScanFilter = ""
  RcFlt.DBOpen _
          "SELECT Tools_Filter.REPORT_ID, Tools_Filter.FIELD_NAME, Tools_Filter.FIELD_TYPE, Tools_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Tools_Filter ON ([Report Modules].NoIdx = Tools_Filter.REPORT_ID) Where [Report Modules].NoIdx = '" _
          & GridReport.Columns(0).Text & "' AND [Report Modules].ViewObject = '" & GridReport.Columns(6).Text & "'", CNN, _
          lckLockBatch
  Debug.Print _
          "SELECT Tools_Filter.REPORT_ID, Tools_Filter.FIELD_NAME, Tools_Filter.FIELD_TYPE, Tools_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Tools_Filter ON ([Report Modules].NoIdx = Tools_Filter.REPORT_ID) Where [Report Modules].NoIdx = '" _
          & GridReport.Columns(0).Text & "' AND [Report Modules].ViewObject = '" & GridReport.Columns(6).Text & "'"
  Debug.Print RcFlt.Recordcount

  If Not RcFlt Is Nothing Then
    If RcFlt.Recordcount <> 0 Then
      'Set RcFlt = RcFilter.DBRecordset.Clone(adLockReadOnly)

      With RcFlt

        If .Recordcount <> 0 Then
          ' .Filter = "Isi <> ''"

          If .Recordcount <> 0 Then
            .DBRecordset.MoveFirst

            Do

              If .DBRecordset.EOF = True Then Exit Do

              Select Case .Fields(2).Value

                Case 220
                  mVarTmp = Trim(mVarTmp & "[" & .Fields(0) & "] Like N'" & .Fields(1) & "%'") & " AND "

                Case 63
                  mVarTmp = Trim(mVarTmp & "[" & .Fields(0) & "] = " & .Fields(1)) & " AND "

                Case 0
                  mVarTmp = mVarTmp & "[" & .Fields(0) & "] >= Convert(datetime,'" & Format(.Fields(1), "dd/mm/yy") & "',3)" & _
                          " AND "
              End Select

              .DBRecordset.MoveNext
            Loop

            .DBRecordset.MoveFirst
            ScanFilter = " WHERE " & Left(mVarTmp, Len(mVarTmp) - 5)
          End If
        End If

      End With

    End If
  End If

  'CloseDB RcFlt
End Function

Private Function ScanTypeFld(ByVal Param As String) As TypeFld

  Select Case Rc.DBRecordset.Fields(Param).Type

    Case 135

      ScanTypeFld = fldTanggal

    Case 131, 3, 4, 5, 6

      ScanTypeFld = fldNumeric

    Case Else
      ScanTypeFld = fldString
  End Select

End Function

Private Sub chkOpt_Click(Index As Integer)
  OpenDB chkOpt(Index).Tag
  GridLayout
End Sub

Private Sub CmbFilter_Click(Index As Integer)

  'If Index = 0 Then ListFilter.SelectedItem.SubItems(0) = CmbFilter(0).Text
  If Index = 0 Then ListFilter.SelectedItem.SubItems(1) = CmbFilter(0).Text
  If Index = 1 Then ListFilter.SelectedItem.SubItems(2) = CmbFilter(1).Text
End Sub

Private Sub CmbOperator_Click()

  If (ListFilter.SelectedItem.SubItems(4) = "0" And CmbOperator.Text = "is between") Or (ListFilter.SelectedItem.SubItems(4) = _
          "0" And CmbOperator.Text = "is not between") Then
    DTPickFilter(0).Visible = True
    DTPickFilter(1).Visible = True
  ElseIf (ListFilter.SelectedItem.SubItems(4) = "202" And CmbOperator.Text = "is between") Or (ListFilter.SelectedItem.SubItems( _
          4) = "202" And CmbOperator.Text = "is not between") Then
    CmbFilter(0).Visible = True
    CmbFilter(1).Visible = True
  ElseIf (ListFilter.SelectedItem.SubItems(4) = "202" And CmbOperator.Text <> "is between") Or ( _
          ListFilter.SelectedItem.SubItems(4) = "202" And CmbOperator.Text <> "is not between") Then
    CmbFilter(0).Visible = True
    CmbFilter(1).Visible = False
  End If

End Sub

Private Sub CmdCancel_Click(Index As Integer)
  PictMain.Visible = False
End Sub

Private Sub cmdLink_Click()
  On Error GoTo RepERR

  With Dialog
    .InitDir = ReportPos  'App.Path & "\Report"
    .Filter = "*.rpt|*.rpt" '"Crystal Report"
    .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
    .ShowOpen

    If .Filename = "" Then
    Else
      txtBox(0) = .FileTitle
    End If

  End With

  'txtBox(0).SetFocus
RepERR:


  If Err <> 0 Then
    MsgBox Err.Description & " - " & Err.Number, vbCritical, App.ProductName
  End If

End Sub

Private Sub CmdOK_Click(Index As Integer)

  If Index = 0 Then

    '    Debug.Print CmbOperator.ListIndex
    If CmbOperator.ListIndex <> -1 Then

      Select Case ListFilter.SelectedItem.SubItems(4)

        Case 0 'DATE PICKER

          If Trim(ListFilter.SelectedItem.SubItems(4)) = "0" Then
            If DTPickFilter(0).Value = "" Or DTPickFilter(1).Value = "" Then
              MessageBox "Data Belum Lengkap!!"
              Exit Sub
            End If

            ActivateObject oDTPicker, False
            ListFilter.SelectedItem.SubItems(5) = GetOperator(CmbOperator.ListIndex)
            ListFilter.SelectedItem.SubItems(1) = DTPickFilter(0).Value
            ListFilter.SelectedItem.SubItems(2) = DTPickFilter(1).Value
          End If

        Case 202 ' BERARTI COMBO

          If CmbFilter(0).ListIndex = -1 Or CmbFilter(1).ListIndex = -1 And CmbFilter(1).Visible = True Then
            MessageBox "Data Belum Lengkap!!"
            Exit Sub
          End If

          ActivateObject oComboBx, False
          ListFilter.SelectedItem.SubItems(5) = GetOperator(CmbOperator.ListIndex)
          ' ListFilter.SelectedItem.SubItems(1) = CmbFilter(0).Text
          ' ListFilter.SelectedItem.SubItems(2) = CmbFilter(1).Text
         
      End Select

      '        ListFilter.SelectedItem.SubItems(6) = GetOperator(CmbOperator.ListIndex)
      PictMain.Visible = False
    Else
      MsgBox "Select Operator", vbExclamation, App.ProductName
    End If

  Else
    PictMain.Visible = False
    ListFilter.SelectedItem.Checked = False
  End If

End Sub

Private Sub CmdTombol_Click(Index As Integer)

  Select Case Index

    Case 0
      Unload Me

    Case 1
      'PREVIEW
      Screen.MousePointer = vbHourglass
      Preview
      Screen.MousePointer = 0

    Case 2
      'NEW
      TabReport.Tab = 1
      rsReport.AddNew
      txtBox(0) = ""
      txtBox(1) = ""

    Case 3
      '        Save
      ExecSave

    Case 4

      'DELETE
      If rsReport.Recordcount <> 0 Then
        If TabReport.TabIndex = 1 Then
          If MsgBox("Delete Report ' " & DgDesign.Columns("Description").Text & " ' ?", vbQuestion + vbYesNo, App.ProductName) _
                  = vbYes Then
            strSQL = "DELETE FROM [Report Modules] Where (NoIdx = '" & DgDesign.Columns(0).Value & "')"
            SendDataToServer strSQL
          End If

        Else

          If MsgBox("Delete Report ' " & GridReport.Columns("Description").Text & " ' ?", vbQuestion + vbYesNo, _
                  App.ProductName) = vbYes Then
            strSQL = "DELETE FROM [Report Modules] Where (NoIdx = '" & GridReport.Columns(0).Value & "')"
            SendDataToServer strSQL
          End If
        End If
      End If

    Case 5
      GenerateFld
      '        OpenFlt
  End Select

  Call chkOpt_Click(IdxOpt)
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
  Set mCall = New frmCaller
    
  Select Case Index

    Case 0
      RcProses.DBOpen "Select name as [Daftar View] From sysobjects where xtype ='V' order by name", CNN, lckLockReadOnly

  End Select
    
  If RcProses.Recordcount <> 0 Then

    Select Case Index

      Case 0
        mCall.FromTagActive = "LIST VIEW"

      Case 1
        mCall.FromTagActive = "FORMULA EKSTRAKSI"
    End Select

    Set mCall.FormData = RcProses.DBRecordset
    mCall.LookUp Me
  Else
    MessageBox "Konfigurasi Pra Produksi Masih Kosong" & vbCrLf & "Silakan Isi Form Formula Ekstraksi", "Peringatan", msgOkOnly
    OpenPartner = True
  End If

End Function

Private Sub DgDesign_Click()
  'GridReport.Row = DgDesign.Bookmark
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

  Select Case TagForm

    Case "LIST VIEW"

      With RcProses
        LblView = RcProses.Fields("Daftar View")
      End With
            
  End Select

End Sub

Private Sub cmdView_Click()
  OpenPartner 0
End Sub

Private Sub DTPickFilter_Change(Index As Integer)

  If Index = 0 Then ListFilter.SelectedItem.SubItems(1) = DTPickFilter(0).Value
  If Index = 1 Then ListFilter.SelectedItem.SubItems(2) = DTPickFilter(1).Value
End Sub

Private Sub Form_Activate()
  'If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

  If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  Dim rsTemp As ADODB.Recordset
  PassTengah Me, MainMenu
  '## Set Priority Code on WO ##
  Set rsTemp = New ADODB.Recordset
  rsTemp.CursorLocation = adUseClient
  rsTemp.Open "SELECT GroupID, GroupName FROM [Report Group]", CNN, adOpenStatic, adLockReadOnly, adCmdText
  Set CmbModule.RowSource = rsTemp
  CmbModule.ListField = rsTemp.Fields(1).Name
  CmbModule.BoundColumn = rsTemp.Fields(0).Name

  TabReport.Tab = 0
  IdxOpt = 6
  Call chkOpt_Click(IdxOpt)
  GridReport_Click
  GridReport.Columns(1).Width = 0
  DgDesign.Columns(0).Width = 0
  DgDesign.Columns(1).Width = 0
  ' ListFilter.ColumnHeaders(4).Width = 0
  Set mCall = New frmCaller
  ' DgDesign.Row = GridReport.Bookmark
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
  Rc.CloseDB
  RcFlt.CloseDB
  RcFilter.CloseDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmReport = Nothing
End Sub

Private Sub GridFilter_Error(ByVal DataError As Integer, _
                             Response As Integer)
  DataError = 0
  Response = 0
End Sub

Private Sub FrTombol_DragDrop(Source As Control, _
                              X As Single, _
                              Y As Single)

End Sub

Private Sub GridReport_Click()
  Dim vLst As ListItem
  ListFilter.ListItems.Clear

  If GridReport.Tag = "0" Then Exit Sub
  If Rc.DBOpen( _
          "SELECT Tools_Filter.REPORT_ID, Tools_Filter.FIELD_NAME, Tools_Filter.FIELD_TYPE, Tools_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Tools_Filter ON ([Report Modules].NoIdx = Tools_Filter.REPORT_ID) Where [Report Modules].NoIdx = '" _
          & GridReport.Columns(0).Text & "' AND [Report Modules].ViewObject = '" & GridReport.Columns(6).Text & "'", CNN, _
          lckLockBatch) = True Then
    ListFilter.ListItems.Clear

    If Rc.Recordcount <> 0 Then

      With Rc.DBRecordset

        Do While Not .EOF
          Set vLst = ListFilter.ListItems.Add(, , UCase(.Fields("FIELD_NAME").Value))
          vLst.SubItems(4) = IIf(IsNull(.Fields("FIELD_TYPE").Value), "", UCase(.Fields("FIELD_TYPE").Value))
          .MoveNext
        Loop

      End With

    Else
      ListFilter.ListItems.Clear
    End If
  End If
'Debug.Print GridReport.Row
  ' DgDesign.Row = GridReport.Bookmark
End Sub

Private Sub GridReport_DblClick()
  Dim mVarWindState As Integer
  mVarWindState = Me.WindowState
  Preview
  Me.WindowState = mVarWindState
End Sub

Private Sub ListFilter_ItemCheck(ByVal Item As MSComctlLib.ListItem)

  If Item.Checked = False Then
    PictMain.Visible = False
    ListFilter.SelectedItem.SubItems(1) = ""
    ListFilter.SelectedItem.SubItems(2) = ""
  Else
    PictMain.Visible = True
    CmbOperator.ListIndex = -1
    Set ListFilter.SelectedItem = Item
    LBLFilter = _
            "Filter Selection SELECT [PO Order].PurchaseID,Inventory.ItemName,PartnerDB.CompanyName,PartnerDB.Address,PartnerDB.City From [PO Order] INNER JOIN PartnerDB ON ([PO Order].PartnerID = PartnerDB.PartnerID) INNER JOIN [Detail PO] ON ([PO Order].PurchaseID = [Detail PO].PurchaseID)  INNER JOIN Inventory ON ([Detail PO].NoItem = Inventory.NoItem) INNER JOIN [Inventory Group] ON (Inventory.NoGroup = [Inventory Group].NoGroup) Where [PO Order].StatusSJ = 0 AND LEFT([PO Order].PurchaseID, 2) = 'PO' AND Inventory.NoGroup = 'PBP' Order By  [PO Order].PurchaseID by : " _
            & ListFilter.SelectedItem.Text
    
    Select Case ListFilter.SelectedItem.SubItems(4)

      Case 202

        ActivateObject oComboBx, False

      Case 203, 3, 5, 6  'TEXTBOX

        If ListFilter.SelectedItem.SubItems(3) = "" Then
          ActivateObject oTextBx, False
          TxtFilter(0).Left = ObjLeft
          TxtFilter(1).Left = ObjLeft
          TxtFilter(0).Top = ObjTop1
          TxtFilter(1).Top = ObjTop2
          TxtFilter(0).Text = ""
          TxtFilter(1).Text = ""
        Else

          If Trim(ListFilter.SelectedItem.SubItems(3)) = "0" Then
            ActivateObject oDTPicker, False
          Else
            ActivateObject oComboBx, False
          End If
        End If

        '            TxtFilter(0).SetFocus
      Case 0    'DATEPICKER
        ActivateObject oDTPicker, False
    End Select

  End If

End Sub

Private Sub rsReport_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                                  ByVal pError As ADODB.Error, _
                                  adStatus As ADODB.EventStatusEnum, _
                                  ByVal pRecordset As ADODB.Recordset)
   
  ' GenerateFld
  ' OpenFlt
End Sub

