VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportConfig 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manajemen Laporan"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10395
   Begin MSComDlg.CommonDialog dialog 
      Left            =   4380
      Top             =   4725
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabReport 
      Height          =   5445
      Left            =   5220
      TabIndex        =   1
      Top             =   135
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9604
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   15380335
      TabCaption(0)   =   "Detil"
      TabPicture(0)   =   "frmReportConfig.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "PictReport(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Filter"
      TabPicture(1)   =   "frmReportConfig.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PictReport(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Group Access"
      TabPicture(2)   =   "frmReportConfig.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LViewReport"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Report Module"
      TabPicture(3)   =   "frmReportConfig.frx":68A6
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "PictReport(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.PictureBox PictReport 
         BackColor       =   &H00EAAF6F&
         Height          =   4995
         Index           =   2
         Left            =   75
         ScaleHeight     =   4935
         ScaleWidth      =   4845
         TabIndex        =   32
         Top             =   375
         Width           =   4900
         Begin MSDataGridLib.DataGrid GridModule 
            Height          =   4050
            Left            =   0
            TabIndex        =   13
            Tag             =   "KP"
            Top             =   885
            Width           =   4830
            _ExtentX        =   8520
            _ExtentY        =   7144
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BorderStyle     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   6
            FormatLocked    =   -1  'True
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "GroupID"
               Caption         =   "No Group"
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
               DataField       =   "GroupName"
               Caption         =   "Nama Group"
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
               DataField       =   "Report"
               Caption         =   "Report"
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtReport 
            Appearance      =   0  'Flat
            DataField       =   "GroupName"
            Enabled         =   0   'False
            Height          =   345
            Index           =   4
            Left            =   1095
            TabIndex        =   12
            Top             =   435
            Width           =   2775
         End
         Begin VB.TextBox txtReport 
            Appearance      =   0  'Flat
            DataField       =   "GroupID"
            Enabled         =   0   'False
            Height          =   330
            Index           =   3
            Left            =   1095
            TabIndex        =   11
            Top             =   75
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Group"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   105
            TabIndex        =   34
            Top             =   510
            Width           =   885
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   1640
            X2              =   45
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Group"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   33
            Top             =   150
            Width           =   675
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1625
            X2              =   45
            Y1              =   390
            Y2              =   390
         End
      End
      Begin VB.PictureBox PictReport 
         BackColor       =   &H00EAAF6F&
         Height          =   4980
         Index           =   0
         Left            =   -74925
         ScaleHeight     =   4920
         ScaleWidth      =   4845
         TabIndex        =   23
         Top             =   375
         Width           =   4900
         Begin VB.TextBox txtReport 
            Appearance      =   0  'Flat
            DataField       =   "FileNameReport"
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   1215
            TabIndex        =   3
            Top             =   165
            Width           =   2775
         End
         Begin VB.TextBox txtReport 
            Appearance      =   0  'Flat
            DataField       =   "Description"
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   1215
            TabIndex        =   5
            Top             =   900
            Width           =   2775
         End
         Begin VB.CommandButton cmdLook 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   3990
            Picture         =   "frmReportConfig.frx":68C2
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   180
            Width           =   330
         End
         Begin VB.CommandButton cmdLook 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   3990
            Picture         =   "frmReportConfig.frx":D114
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1635
            Width           =   330
         End
         Begin VB.TextBox txtReport 
            Appearance      =   0  'Flat
            DataField       =   "alias report"
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   1215
            TabIndex        =   4
            Top             =   525
            Width           =   2775
         End
         Begin MSDataListLib.DataCombo CmbModule 
            DataField       =   "ReportGroup"
            Height          =   315
            Left            =   1215
            TabIndex        =   6
            Top             =   1275
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   714
            _Version        =   393216
            IntegralHeight  =   0   'False
            Enabled         =   0   'False
            Appearance      =   0
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            DataField       =   "noIdx"
            Height          =   285
            Left            =   2640
            TabIndex        =   31
            Top             =   2040
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   1745
            X2              =   165
            Y1              =   1935
            Y2              =   1935
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   1745
            X2              =   165
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Module"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   225
            TabIndex        =   28
            Top             =   1335
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Laporan"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   27
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   26
            Top             =   975
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View Object"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   25
            Top             =   1703
            Width           =   855
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   1745
            X2              =   165
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   1745
            X2              =   165
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label LBLReport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "ViewObject"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   1215
            TabIndex        =   8
            Top             =   1620
            Width           =   2775
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   1760
            X2              =   165
            Y1              =   855
            Y2              =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Judul"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   24
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.PictureBox PictReport 
         Height          =   4980
         Index           =   1
         Left            =   -74925
         ScaleHeight     =   4920
         ScaleWidth      =   4845
         TabIndex        =   22
         Top             =   375
         Width           =   4900
         Begin MSDataGridLib.DataGrid GridConfig 
            Bindings        =   "frmReportConfig.frx":13966
            Height          =   5505
            Left            =   0
            TabIndex        =   9
            Tag             =   "KP"
            Top             =   0
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   9710
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   6
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "FORM_NAME"
               Caption         =   "Form Name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "KODE_FORM"
               Caption         =   "Kode Form"
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
               DataField       =   "FIELD_NAME"
               Caption         =   "FILTER"
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
               DataField       =   "FIELD_TYPE"
               Caption         =   "TIPE"
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
            BeginProperty Column04 
               DataField       =   "OBJECT_TYPE"
               Caption         =   "Object Type"
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
            BeginProperty Column05 
               DataField       =   "idx"
               Caption         =   "idx"
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
                  DividerStyle    =   6
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column02 
                  DividerStyle    =   6
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView LViewReport 
         Height          =   5040
         Left            =   -74955
         TabIndex        =   10
         Top             =   345
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8890
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "GroupName"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgAccount 
      Left            =   4650
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportConfig.frx":1397A
            Key             =   "SEGITIGA"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportConfig.frx":1A1DC
            Key             =   "ABANG"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportConfig.frx":20A3E
            Key             =   "BIRU"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportConfig.frx":272A0
            Key             =   "IJO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TView 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   9234
      _Version        =   393217
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgAccount"
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
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   10395
      TabIndex        =   19
      Top             =   5805
      Width           =   10395
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Edit"
         Height          =   555
         Index           =   2
         Left            =   870
         Picture         =   "frmReportConfig.frx":2DB02
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Refresh"
         Height          =   555
         Index           =   6
         Left            =   3750
         Picture         =   "frmReportConfig.frx":34354
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Tambah"
         Height          =   555
         Index           =   1
         Left            =   150
         Picture         =   "frmReportConfig.frx":3ABA6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   100
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
         TabIndex        =   20
         Top             =   0
         Width           =   10530
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Batal"
         Height          =   555
         Index           =   4
         Left            =   2310
         Picture         =   "frmReportConfig.frx":413F8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Hapus"
         Height          =   555
         Index           =   5
         Left            =   3030
         Picture         =   "frmReportConfig.frx":47C4A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Save"
         Height          =   555
         Index           =   3
         Left            =   1590
         Picture         =   "frmReportConfig.frx":4E49C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   4470
         Picture         =   "frmReportConfig.frx":54CEE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   100
         Width           =   720
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "DAFTAR LAPORAN"
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
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   4980
   End
End
Attribute VB_Name = "frmReportConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsConfig As ADODB.Recordset
Dim rsModule As ADODB.Recordset
Dim myPart As New utility
Dim vNode As Node
Dim strSQL As String
Dim SelectNode As Node
Dim RsDetail As ADODB.Recordset
Dim rsTemp As New DBQuick
Dim RcProses As New DBQuick
Dim Posisi As String
Dim nPosisi As TypePost
Dim ReportId As String
Dim AliasReport As String
Dim MvarItem As Integer
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Public Enum TypePost
    TypeEDIT = 0
    TypeNEW = 1
End Enum

Private Sub ConvertFieldType()
  '11  Bit
  '135 DateTime
  '202 nvarchar
  '203 ntext
  '3   int
  '5   Float
  '6   Money

End Sub

Private Sub GridLayout()

With GridConfig
  .Columns(0).Visible = False
  .Columns(1).Visible = False
  .Columns(5).Visible = False
  .Columns(2).Locked = True
  .Columns(3).Locked = True
  .Columns(4).Locked = True
  .Columns(5).Locked = True
  .Columns(4).Button = True
  .Columns(5).Button = True
  .Columns(3).Alignment = dbgCenter
End With

With GridModule
    
    If rsModule.Recordcount <> 0 Then
        If rsModule.Recordcount > 16 Then
            .Columns(1).width = 2480
        Else
            .Columns(1).width = 2710
        End If
    End If
End With
End Sub

Private Sub LoadTree()
On Error GoTo xErr


Dim rsForms As DBQuick
Dim rsChild As New Recordset
Dim No  As Integer

TView.Nodes.Clear
Set rsForms = New DBQuick

strSQL = "Shape{SELECT GroupID , GroupName FROM [report group]} as ParentNode append " & _
  " ({SELECT * FROM [Report Modules] ORDER BY ReportGroup, [Alias Report] } as ChildNode relate GroupID to ReportGroup)"
  
rsForms.DBOpen strSQL, CNN, lckLockReadOnly
No = 1

If rsForms.Recordcount > 0 Then
'    rsForms.MoveFirst
    FirstNode = Trim(rsForms.DBRecordset.Fields(0).Value)
    Set rsChild = rsForms.DBRecordset("ChildNode").Value
    With rsForms.DBRecordset
        Do While Not .EOF
            With TView.Nodes.Add(, , .Fields(0).Value, .Fields(1).Value, "BIRU")
                .Bold = True
'                .Expanded = True
            End With
            If rsChild.Recordcount <> 0 Then
                Do While Not rsChild.EOF
                    Set vNode = TView.Nodes.Add(.Fields(0).Value, tvwChild, CStr(rsChild.Fields("NoIdx").Value) & "A", rsChild.Fields("Alias Report").Value, "IJO", "IJO")
                    vNode.Tag = rsChild.Fields("ViewObject").Value
                    rsChild.MoveNext
                Loop
            End If
            .MoveNext
        Loop
    End With
    
'    While Not rsForms.EOF
'        Set vNode = TView.Nodes.Add(, , CStr(rsForms.Fields("NoIdx").Value) & "A", Trim(IIf(IsNull(rsForms.Fields( _
'                "Alias Report").Value), " ", rsForms.Fields("Alias Report").Value)))
'        'Set vNode = TVConfig.Nodes.Add("A", tvwChild, "A" & Trim(rsForms.Fields(0).Value), Trim(rsForms.Fields(0).Value))
'        vNode.Tag = No
'        vNode.Expanded = True
'        No = No + 1
'        rsForms.MoveNext
'    Wend
End If

  rsForms.CloseDB
  Set rsForms = Nothing
  
'  Set rsConfig = myPart.OpenDB("SELECT [REPORT_ID], [FIELD_NAME], [FIELD_TYPE] From Report_Filter WHERE ([REPORT_ID]= N'" & _
'          UCase(FirstNode) & "')")
'  Set GridConfig.DataSource = rsConfig
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub
Private Sub LoadGroup()
strSQL = "SELECT [report group].GroupID, [report group].GroupName, COUNT([report modules].[Alias Report]) AS Report " & _
        " FROM [report group]  LEFT OUTER JOIN [report modules] ON [report group].GroupID = [report modules].ReportGroup  " & _
        " GROUP BY [report group].GroupID, [report group].GroupName ORDER BY [report group].GroupName"

Set rsModule = myPart.OpenDB(strSQL)
Set GridModule.DataSource = rsModule
Set txtReport(3).DataSource = rsModule
Set txtReport(4).DataSource = rsModule
End Sub
Private Sub cmdLook_Click(Index As Integer)
On Error GoTo RepERR

Select Case Index
    
    Case 0
        With dialog
            .InitDir = ReportPos  'App.Path & "\Report"
            .Filter = "*.rpt|*.rpt" '"Crystal Report"
            .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
            .ShowOpen
        
            If .Filename = "" Then
            Else
                txtReport(0) = .FileTitle
            End If
        
        End With
    
    Case 1
        OpenPartner
RepERR:
    If Err <> 0 Then
        MessageBox Err.Description & " - " & Err.Number, App.ProductName, msgOkOnly, msgExclamation
    End If

End Select
End Sub
Private Sub OpenPartner()


RcProses.DBOpen "Select name as [Daftar View] From sysobjects where xtype ='V' order by name", CNN, lckLockReadOnly

If RcProses.Recordcount <> 0 Then
    mCall.FromTagActive = "LIST VIEW"
    Set mCall.FormData = RcProses.DBRecordset
    mCall.LookUp Me
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
rsTemp.CloseDB
Set rsTemp = Nothing
'rsConfig.Close
Set rsConfig = Nothing
rsModule.Close
Set rsModule = Nothing
'RsDetail.Close
Set RsDetail = Nothing
RcProses.CloseDB
Set RcProses = Nothing
End Sub

Private Sub LViewReport_ItemCheck(ByVal Item As MSComctlLib.ListItem)
MvarItem = Item.Index
InsertToAccess FirstNode, Item.Text
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm

        Case "LIST VIEW"

            With RcProses
                LBLReport(0) = RcProses.DBRecordset.Fields("Daftar View").Value
            End With
            
    End Select

End Sub
Private Sub TView_NodeClick(ByVal Node As MSComctlLib.Node)
  Set SelectNode = Node
  If SelectNode.Root.Text = "" Then Exit Sub
  SetDataConfig
  Label2(1).Caption = "  DAFTAR LAPORAN - " & Node.Text
End Sub
Private Sub SetDataConfig()
Dim idxToStr As String
Dim I As Long

idxToStr = TView.SelectedItem.Key
If Not SelectNode.Parent Is Nothing Then
    Set rsConfig = myPart.OpenDB("SELECT * From Report_Filter WHERE (REPORT_ID = N'" & Replace(idxToStr, "A", "", 1, Len( _
            TView.SelectedItem.Key)) & "')")
            
    Set GridConfig.DataSource = rsConfig
    strSQL = "SELECT NoIdx, Description, [Alias Report], ReportGroup, FileNameReport, ViewObject " & _
            " From dbo.[report modules] Where (NoIdx = N'" & Replace(idxToStr, "A", "", 1, Len( _
            TView.SelectedItem.Key)) & "') ORDER BY [Alias Report]"
    Set RsDetail = myPart.OpenDB(strSQL)
    Set LBLReport(0).DataSource = RsDetail
    For I = 0 To 2
        Set txtReport(I).DataSource = RsDetail
    Next
    Set CmbModule.DataSource = RsDetail
    Set Label3.DataSource = RsDetail
    FirstNode = Replace(idxToStr, "A", "", 1, Len(TView.SelectedItem.Key))
    If TabReport.Tab = 2 Then
        'ListGroupAkses FirstNode, aksess.GetID
        ListGroupAkses FirstNode 'Tampil List DI List View
    End If
End If
End Sub
Private Sub LockButton(bStatus As Boolean)
CmdTombol(1).Enabled = bStatus      'NEW
CmdTombol(2).Enabled = bStatus      'EDIT
CmdTombol(3).Enabled = Not bStatus      'SAVE
CmdTombol(4).Enabled = Not bStatus      'CANCEL
CmdTombol(5).Enabled = Not bStatus      'DELETE
CmdTombol(6).Enabled = bStatus      'REFRESH
TView.Enabled = bStatus

End Sub
Private Sub CmdTombol_Click(Index As Integer)
On Error GoTo Hell
Dim NodeIdx As Long

If TabReport.Tab <> 3 Then
    If SelectNode Is Nothing Then
        If Index = 0 Then
            Unload Me
        Else
            Exit Sub
        End If
    Else
        If Index = 0 Then
            Unload Me
            Exit Sub
        End If
        If SelectNode.Parent Is Nothing Then
            Exit Sub
        Else
            NodeIdx = SelectNode.Index
        End If
    End If
End If
Select Case Index
    Case 0  'EXIT
        Unload Me

    Case 1  'NEW
        Select Case TabReport.Tab
            Case 0  'DETIL
                'BindControlToData
                Posisi = "new"
                nPosisi = TypeEDIT
                TombolEdit True
                txtReport(0) = ""
                txtReport(1) = ""
                txtReport(2) = ""
                LBLReport(0) = ""
                CmbModule.Text = ""
                LockButton False
            
            Case 1  'FILTER
                frmEntryForm.OperationMode = "Insert"
                frmEntryForm.FormName = Trim(TView.SelectedItem.Text)
                frmEntryForm.Show vbModal
                LoadTree
                TView.SetFocus
                TView.Nodes(NodeIdx).Selected = True
                TView_NodeClick TView.SelectedItem
            Case 3  'MODULE REPORT
                Posisi = "new"
                nPosisi = TypeNEW
                rsModule.AddNew
                txtReport(3) = ""
                txtReport(4) = ""
                TombolReportModule True
                LockButton False
            Case Else
                
        End Select
        
    Case 2  'EDIT
        
        Select Case TabReport.Tab
            Case 0
                TabReport.Tab = 0
                Posisi = "edit"
                TombolEdit True
                LockButton False
            Case 1
                If Not SelectNode.Parent Is Nothing Then
                    LockButton False
                    frmEntryForm.OperationMode = "Edit"
                    frmEntryForm.FormName = Trim(TView.SelectedItem.Text)
        '            frmEntryForm.ViewObject = FirstNode
                    frmEntryForm.Show vbModal
                    LoadTree
                    TView.Enabled = True
'                    LockButton False
                    TView.SetFocus
                    TView.Nodes(NodeIdx).Selected = True
                    TView_NodeClick TView.SelectedItem
                End If
            Case 3  'REPORT MODULE
                Posisi = "edit"
                nPosisi = TypeEDIT
                TombolReportModule True
                LockButton False
             
        End Select

    
    Case 3  'SAVE
        Select Case TabReport.Tab
            Case 0
                ExecSave
                LoadTree
                TombolEdit False
                LockButton True
            
            Case 1
                If Not rsConfig Is Nothing Then
                    rsConfig.UpdateBatch adAffectAllChapters
                    MessageBox "Save data successfully...", "Konfigurasi Laporan", msgOkOnly, msgInfo
                    LockButton True
                End If
            Case 3
                If Posisi = "new" Then
                    ExecSaveReportModule nPosisi
                Else
                    ExecSaveReportModule nPosisi
                End If
                LoadTree
                LoadGroup
                TombolReportModule False
                LockButton True
                
        End Select
    Case 4  'BATAL
        LockButton True
        Select Case TabReport.Tab
            Case 0
                TombolEdit False
            Case 3
                LoadGroup
                TombolReportModule False
        End Select
    Case 5  'DELETE
        Select Case TabReport.Tab
            Case 0
                If MessageBox("Yakin data akan dihapus ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
                   SendDataToServer "delete from [report modules] where noIdx = '" & Label3.Caption & "'"
                End If
                LoadTree
                txtReport(0) = ""
                txtReport(1) = ""
                txtReport(2) = ""
                LBLReport(0) = ""
                CmbModule.Text = ""
                LockButton True
                TombolEdit False
            Case 3
                If rsModule.Fields(2).Value = 0 Then
                    If MessageBox("Yakin data akan dihapus ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
                       SendDataToServer "delete from [report group] where groupid = '" & txtReport(3).Text & "'"
                    End If
                    LoadGroup
                    LoadTree
                Else
                    MessageBox "Masih terdapat file report dalam group ' " & UCase(txtReport(4).Text) & " '", "Kontrol Delete", msgOkOnly, msgExclamation
                End If
                LockButton True
                TombolReportModule False
        End Select

     Case 6
        LoadTree
End Select

Exit Sub
Hell:
    If Err.Number = 35605 Then
        MessageBox "Baris ini sudah dihapus", "Informasi", msgOkOnly, msgExclamation
    Else
        MessageBox Err.Description, "Kontrol Konfigurasi", msgOkOnly, msgExclamation
    End If
  Err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

  If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    '*** Load data group / module untuk data combo di tab design
    rsTemp.DBOpen "SELECT GroupID, GroupName FROM [Report Group]", CNN
    Set CmbModule.RowSource = rsTemp.DBRecordset
    CmbModule.ListField = rsTemp.DBRecordset.Fields(1).Name
    CmbModule.BoundColumn = rsTemp.DBRecordset.Fields(0).Name
    
    LoadTree
    LoadGroup
    GridLayout
    GridConfig.Columns(4).width = 0
    CenterForm Me, Me
    LockButton True
    TabReport.Tab = 0
    
    Set mCall = New frmCaller
    
End Sub

Private Sub GridConfig_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo 1
  Select Case ColIndex

    Case 4
      If GridConfig.AllowUpdate = True Then
        FrmLookUp.TitleForm = "Object Type"
        Set FrmLookUp.FormCaller = Me
        Set FrmLookUp.FormContainer = GridConfig.Columns(ColIndex)
        '    Set FrmLookUp.FormContainer2 = GridConfig.Columns(ColIndex + 1)
        
        '   strSQL = "SELECT [Catalogue No], [Engineering Description], [Equipment ID] From [Equipment APLs Table]"
        strSQL = "SELECT Object, [Object Name] From Tools_ObjectView ORDER BY [Object Name]"
        FrmLookUp.SQLScript = strSQL
        FrmLookUp.ColRefNumber = 0
        FrmLookUp.ColRefNumber2 = 0
        FrmLookUp.ColRefNumber3 = 0
        Load FrmLookUp
        FrmLookUp.Show vbModal
        GridConfig.SetFocus
        '    GridConfig.Col = 4
        '    GridConfig.EditActive = True
      End If

    Case 5

      If GridConfig.AllowUpdate = True Then
        
        FrmLookUp.TitleForm = "Data Combo"
        Set FrmLookUp.FormCaller = Me
        Set FrmLookUp.FormContainer = GridConfig.Columns(ColIndex)
        '    Set FrmLookUp.FormContainer2 = GridConfig.Columns(ColIndex + 1)
        
        '   strSQL = "SELECT [Catalogue No], [Engineering Description], [Equipment ID] From [Equipment APLs Table]"
        strSQL = "SELECT Table_Name,Table_Type From kolom_object where table_type='VIEW' Group by Table_Name,Table_Type"
        FrmLookUp.SQLScript = strSQL
        FrmLookUp.ColRefNumber = 0
        FrmLookUp.ColRefNumber2 = 0
        FrmLookUp.ColRefNumber3 = 0
        Load FrmLookUp
        FrmLookUp.Show vbModal
        GridConfig.SetFocus
      End If

  End Select
Exit Sub
1:
MessageBox Err.Description, "frmReportConfig:gridconfig_buttonclick" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub TombolEdit(TombolEdit As Boolean)
    cmdLook(0).Enabled = TombolEdit
    cmdLook(1).Enabled = TombolEdit
    txtReport(1).Enabled = TombolEdit
    txtReport(2).Enabled = TombolEdit
    CmbModule.Enabled = TombolEdit
End Sub
Private Sub TombolReportModule(TombolEdit As Boolean)
    txtReport(3).Enabled = TombolEdit
    txtReport(4).Enabled = TombolEdit
    GridModule.Enabled = Not TombolEdit
    TabReport.TabEnabled(0) = Not TombolEdit
    TabReport.TabEnabled(1) = Not TombolEdit
    TabReport.TabEnabled(2) = Not TombolEdit
End Sub
Private Sub ExecSaveReportModule(nStatus As TypePost)
If MessageBox("Apakah Data ini akan disimpan ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
    Select Case nStatus
        Case TypeEDIT
            SendDataToServer "update [report group]  set GroupID = '" & UCase(txtReport(3).Text) & "', GroupName = '" & txtReport(4).Text & "' where GroupID = '" & rsModule.Fields(0).UnderlyingValue & "'"
        Case TypeNEW
            SendDataToServer "insert into [report group] (GroupID, GroupName) values ('" & UCase(txtReport(3).Text) & "', '" & txtReport(4).Text & "')"
    End Select
End If
End Sub

Private Sub ExecSave()
    Dim obj As Object

    For Each obj In Me

        If TypeOf obj Is OptionButton Then
            If obj.Value = True Then ReportId = obj.ToolTipText
        End If

    Next obj
    AliasReport = Mid(txtReport(0), 1, Len(txtReport(0)) - 4)
    If Posisi = "edit" Then
       If MessageBox("Apakah Data Ini akan Diubah ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
          SendDataToServer "update [report modules] set Description = '" & txtReport(2).Text & "', ReportGroup = '" & CmbModule.BoundText & "', FileNameReport = '" & txtReport(0).Text & "', ViewObject = '" & LBLReport(0).Caption & "' , IDreport = '" & ReportId & "' , [Alias Report] = '" & txtReport(1).Text & "' where noIdx = '" & Label3.Caption & "'"
         ' MessageBox "Ubah Module Laporan Sukses", "Konfirmasi", msgOkOnly, msgInfo
       End If
    Else
       If MessageBox("Apakah Data ini akan disimpan ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
          SendDataToServer "insert into [report modules] (Description, ReportGroup, FileNameReport, ViewObject, IDreport, [Alias Report]) values ('" & txtReport(2).Text & "', '" & CmbModule.BoundText & "', '" & txtReport(0).Text & "', '" & LBLReport(0).Caption & "', '" & ReportId & "', '" & txtReport(1).Text & "')"
         ' InsReportPermit
         ' MessageBox "Simpan Module Laporan Sukses", "Konfirmasi", msgOkOnly, msgInfo
       End If
    End If

End Sub


'Private Sub GenerateFld()
'    Dim I As Integer
'    Dim vLst As ListItem
'    Dim rsView As ADODB.Recordset
'    Dim ReportIndex As Integer
'
'    For Each obj In Me
'
'        If TypeOf obj Is OptionButton Then
'            If obj.Value = True Then ReportId = obj.ToolTipText
'        End If
'
'    Next obj
'
'    If Not rsReport Is Nothing Then
'        If Not rsReport.EOF Then
'            If Not IsNull(rsReport.Fields("ViewObject").Value) Then
'                sViewObject = rsReport.Fields("ViewObject").Value
'            Else
'                Exit Sub
'            End If
'
'        Else
'            sViewObject = ""
'        End If
'
'    Else
'        Exit Sub
'    End If
'
'    If Rc.DBOpen("SELECT Report_Filter.REPORT_ID, Report_Filter.FIELD_NAME, Report_Filter.FIELD_TYPE, Report_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Report_Filter ON ([Report Modules].NoIdx = Report_Filter.REPORT_ID) Where [Report Modules].NoIdx = '" & CInt(lblReportIndex) & "' AND [Report Modules].ViewObject = '" & sViewObject & "'", CNN, lckLockBatch) = True Then
'
'        If Rc.Recordcount <> 0 Then
'            ListFilter.ListItems.Clear
'
'            With Rc.DBRecordset
'
'                Do While Not .EOF
'                    Set vLst = ListFilter.ListItems.Add(, , UCase(.Fields("FIELD_NAME").Value))
'                    vLst.SubItems(3) = UCase(.Fields("FIELD_NAME").Value)
'                    vLst.SubItems(3) = IIf(IsNull(.Fields("FIELD_TYPE").Value), "", UCase(.Fields("FIELD_TYPE").Value))
'                    .MoveNext
'                Loop
'
'            End With
'
'        Else
'            ListFilter.ListItems.Clear
'        End If
'
'        '////////////////////////////////////////////////////////////////////
'
'    End If
'
'End Sub


'Private Sub ListGroupAkses(ByVal noidx As Integer, ByVal ID As Integer)
Private Sub ListGroupAkses(ByVal noidx As Integer)

Dim rsFieldGroup As ADODB.Recordset
Dim rsChekFieldGroup As ADODB.Recordset
Dim I, x As Integer
Dim strSQLt As String
Dim vLst As ListItem
        
        strSQL = "SELECT     [group name],id From [user_table_group]GROUP BY [group name],id"
        
        LViewReport.ListItems.Clear
        Set rsChekFieldGroup = myPart.OpenDB(strSQL)
        LViewReport.ColumnHeaders(1).Text = "Group Access"
        If rsChekFieldGroup.Recordcount > 0 Then
            While Not rsChekFieldGroup.EOF
                Set vLst = LViewReport.ListItems.Add(, , UCase(rsChekFieldGroup.Fields(0).Value))
                rsChekFieldGroup.MoveNext
            Wend
        End If
        
'        strSQL = "SELECT     dbo.[report permit].[User ID], dbo.[report permit].noidx, dbo.[report permit].Laporan, dbo.user_table_group.[Group Name] " & _
'                  "FROM         dbo.[report permit] RIGHT OUTER JOIN " & _
'                  "dbo.user_table_group ON dbo.[report permit].IDGroup = dbo.user_table_group.id ORDER BY dbo.user_table_group.[Group Name]"
'
        
        strSQLt = "SELECT     TOP (100) PERCENT dbo.[report permit].[User ID], dbo.[report permit].noidx, dbo.[report permit].Laporan, dbo.user_table_group.[Group Name]," & _
                 " dbo.[report permit].IDGroup " & _
                 " FROM         dbo.[report permit] RIGHT OUTER JOIN  " & _
                 " dbo.user_table_group ON dbo.[report permit].IDGroup = dbo.user_table_group.id " & _
                 " Where (dbo.[report permit].noidx = " & noidx & ") " & _
                 " ORDER BY dbo.user_table_group.[Group Name]"

        Set rsFieldGroup = myPart.OpenDB(strSQLt)
        For x = 1 To LViewReport.ListItems.Count
            If rsFieldGroup.Recordcount > 0 Then
                rsFieldGroup.MoveFirst
                Do While rsFieldGroup.EOF <> True
                    If UCase$(LViewReport.ListItems(x).Text) = UCase$(rsFieldGroup.Fields(3).Value) Then
                       LViewReport.ListItems(x).Checked = True
                    Else
                        If LViewReport.ListItems(x).Checked = True Then
                        Else
                           LViewReport.ListItems(x).Checked = False '
                        End If
                    End If
                    rsFieldGroup.MoveNext
                Loop
            End If
        Next
''
                 
       
End Sub

Private Sub InsertToAccess(ByVal noidx As Integer, Groups As String)
Dim rsChekFieldGroup As ADODB.Recordset
Dim x As Integer
    strSQL = "SELECT id,[group name] From [user_table_group] where [group name]='" & Groups & "'"
    Set rsChekFieldGroup = myPart.OpenDB(strSQL)
    If rsChekFieldGroup.Recordcount > 0 Then
       If LViewReport.ListItems(MvarItem).Checked = True Then
           SendDataToServer (" INSERT INTO [report permit] " & _
                         " ([User ID], noidx, laporan,idgroup)" & _
                         " VALUES  (" & aksess.GetID & "," & noidx & ", 1," & rsChekFieldGroup.Fields(0).Value & ")")
           
        Else
            SendDataToServer " DELETE FROM [report permit] where [user id]=" & aksess.GetID & " and noidx=" & noidx & " and idgroup=" & rsChekFieldGroup.Fields(0).Value & ""
        End If
    End If
End Sub


