VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pelaporan"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10305
   Tag             =   "Report Administrator"
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   10305
      TabIndex        =   22
      Top             =   7530
      Width           =   10305
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Refresh"
         Height          =   555
         Index           =   7
         Left            =   5280
         Picture         =   "frmReport.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   100
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Edit"
         Height          =   555
         Index           =   6
         Left            =   2400
         Picture         =   "frmReport.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   100
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         Height          =   30
         Left            =   -45
         TabIndex        =   32
         Top             =   0
         Width           =   10605
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Preview"
         Height          =   555
         Index           =   1
         Left            =   120
         Picture         =   "frmReport.frx":138F6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Delete"
         Height          =   555
         Index           =   4
         Left            =   3840
         Picture         =   "frmReport.frx":1A148
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   100
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   840
         Picture         =   "frmReport.frx":2099A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Batal"
         Height          =   555
         Index           =   5
         Left            =   4560
         Picture         =   "frmReport.frx":22494
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   100
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "New"
         Height          =   555
         Index           =   2
         Left            =   1680
         Picture         =   "frmReport.frx":28CE6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   100
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Save"
         Height          =   555
         Index           =   3
         Left            =   3120
         Picture         =   "frmReport.frx":2F538
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   100
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7935
         Top             =   195
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   10395
      TabIndex        =   19
      Top             =   0
      Width           =   10395
      Begin TabDlg.SSTab SSTab1 
         Height          =   7275
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   12832
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
         TabCaption(0)   =   "Laporan"
         TabPicture(0)   =   "frmReport.frx":35D8A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "ListFilter"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "TVConfig"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "ImageList1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "PictMain"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "GridReport"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Disain"
         TabPicture(1)   =   "frmReport.frx":35DA6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(1)"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid GridReport 
            Height          =   3300
            Left            =   120
            TabIndex        =   3
            Tag             =   "Design"
            Top             =   3720
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   5821
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   2
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "NoIdx"
               Caption         =   "Index"
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
               DataField       =   "viewObject"
               Caption         =   "View Object"
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
               DataField       =   "Alias Report"
               Caption         =   "Judul Report"
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
               DataField       =   "ReportGroup"
               Caption         =   "Report Group"
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
               DataField       =   "FileNameReport"
               Caption         =   "File Name Report"
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
               DataField       =   "Description"
               Caption         =   "Keterangan"
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
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox PictMain 
            BackColor       =   &H80000001&
            Height          =   2220
            Left            =   3600
            ScaleHeight     =   2160
            ScaleWidth      =   5145
            TabIndex        =   38
            Top             =   1200
            Visible         =   0   'False
            Width           =   5205
            Begin VB.PictureBox PictFilter 
               BackColor       =   &H00EAAF6F&
               BorderStyle     =   0  'None
               Height          =   2025
               Left            =   75
               ScaleHeight     =   2025
               ScaleWidth      =   4980
               TabIndex        =   39
               Top             =   45
               Width           =   4980
               Begin VB.PictureBox Picture1 
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
                  TabIndex        =   48
                  Top             =   0
                  Width           =   4980
                  Begin VB.Label LBLFilter 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "DATA SELECTION"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000005&
                     Height          =   195
                     Left            =   30
                     TabIndex        =   49
                     Top             =   45
                     Width           =   1395
                  End
               End
               Begin VB.TextBox TxtFilter 
                  Height          =   315
                  Index           =   0
                  Left            =   2400
                  TabIndex        =   47
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.TextBox TxtFilter 
                  Height          =   315
                  Index           =   1
                  Left            =   2370
                  TabIndex        =   46
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.ComboBox CmbFilter 
                  Height          =   315
                  Index           =   1
                  Left            =   105
                  TabIndex        =   45
                  Text            =   "CmbFilter"
                  Top             =   1065
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.ComboBox CmbFilter 
                  Height          =   315
                  Index           =   0
                  Left            =   105
                  TabIndex        =   44
                  Text            =   "CmbFilter"
                  Top             =   825
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.CommandButton cmdOK 
                  Caption         =   "&OK"
                  Height          =   315
                  Index           =   0
                  Left            =   135
                  TabIndex        =   43
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.CommandButton cmdCancel 
                  Caption         =   "&Cancel"
                  Height          =   315
                  Index           =   1
                  Left            =   1155
                  TabIndex        =   42
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.ComboBox CmbOperator 
                  Height          =   315
                  ItemData        =   "frmReport.frx":35DC2
                  Left            =   135
                  List            =   "frmReport.frx":35DE5
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   480
                  Width           =   2895
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00C0FFFF&
                  Height          =   30
                  Index           =   0
                  Left            =   -30
                  TabIndex        =   40
                  Top             =   1470
                  Width           =   4995
               End
               Begin MSComCtl2.DTPicker DTPickFilter 
                  Height          =   315
                  Index           =   1
                  Left            =   3060
                  TabIndex        =   50
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
                  Format          =   71630851
                  CurrentDate     =   36877
               End
               Begin MSComCtl2.DTPicker DTPickFilter 
                  Height          =   315
                  Index           =   0
                  Left            =   3060
                  TabIndex        =   51
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
                  Format          =   71630851
                  CurrentDate     =   36877
               End
               Begin VB.Label lblAnd 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "And"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   3750
                  TabIndex        =   53
                  Top             =   825
                  Visible         =   0   'False
                  Width           =   345
               End
               Begin VB.Label lblTo 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "To"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   3270
                  TabIndex        =   52
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   225
               End
            End
         End
         Begin VB.Frame Frame1 
            Height          =   6750
            Index           =   1
            Left            =   -74865
            TabIndex        =   33
            Top             =   360
            Width           =   9780
            Begin MSDataGridLib.DataGrid DgDesign 
               Height          =   4305
               Left            =   120
               TabIndex        =   23
               Tag             =   "Design"
               Top             =   2280
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   7594
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               BorderStyle     =   0
               HeadLines       =   2
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
               ColumnCount     =   6
               BeginProperty Column00 
                  DataField       =   "NoIdx"
                  Caption         =   "Index"
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
                  DataField       =   "reportGroup"
                  Caption         =   "Report Group"
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
                  DataField       =   "fileNameReport"
                  Caption         =   "Nama File Report"
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
                  DataField       =   "Alias Report"
                  Caption         =   "Alias Report"
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
                  DataField       =   "viewObject"
                  Caption         =   "View Object"
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
                  DataField       =   "description"
                  Caption         =   "Keterangan"
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
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "alias report"
               Enabled         =   0   'False
               Height          =   345
               Index           =   3
               Left            =   1320
               TabIndex        =   14
               Tag             =   "ana"
               Top             =   660
               Width           =   2775
            End
            Begin VB.CommandButton cmdView 
               Enabled         =   0   'False
               Height          =   315
               Left            =   4095
               Picture         =   "frmReport.frx":35EAC
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   1770
               Width           =   330
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Left            =   4095
               Picture         =   "frmReport.frx":3C6FE
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   315
               Width           =   330
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Description"
               Enabled         =   0   'False
               Height          =   345
               Index           =   1
               Left            =   1320
               TabIndex        =   15
               Tag             =   "ana"
               Top             =   1035
               Width           =   2775
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "FileNameReport"
               Enabled         =   0   'False
               Height          =   330
               Index           =   0
               Left            =   1320
               TabIndex        =   12
               Tag             =   "ana"
               Top             =   300
               Width           =   2775
            End
            Begin MSDataListLib.DataCombo CmbModule 
               DataField       =   "ReportGroup"
               Height          =   315
               Left            =   1320
               TabIndex        =   16
               Top             =   1410
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   714
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Judul"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   165
               TabIndex        =   56
               Top             =   735
               Width           =   375
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   1700
               X2              =   105
               Y1              =   990
               Y2              =   990
            End
            Begin VB.Label Label3 
               DataField       =   "noIdx"
               Height          =   285
               Left            =   5265
               TabIndex        =   54
               Top             =   345
               Width           =   1635
            End
            Begin VB.Label LblView 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "ViewObject"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   1320
               TabIndex        =   17
               Tag             =   "BAHAN"
               Top             =   1755
               Width           =   2775
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   1700
               X2              =   120
               Y1              =   1710
               Y2              =   1710
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   1700
               X2              =   120
               Y1              =   1365
               Y2              =   1365
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "View Object"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   165
               TabIndex        =   37
               Top             =   1830
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   165
               TabIndex        =   36
               Top             =   1110
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "File Laporan"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   165
               TabIndex        =   35
               Top             =   375
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Module"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   165
               TabIndex        =   34
               Top             =   1470
               Width           =   510
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   1700
               X2              =   120
               Y1              =   615
               Y2              =   615
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   1700
               X2              =   120
               Y1              =   2070
               Y2              =   2070
            End
         End
         Begin VB.Frame Frame1 
            Height          =   5025
            Index           =   2
            Left            =   -74880
            TabIndex        =   25
            Top             =   330
            Width           =   8670
            Begin VB.ComboBox Combo1 
               DataField       =   "ModulesName"
               Height          =   315
               ItemData        =   "frmReport.frx":42F50
               Left            =   1230
               List            =   "frmReport.frx":42F69
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Tag             =   "Design"
               Top             =   1260
               Width           =   3555
            End
            Begin VB.TextBox txtBox 
               DataField       =   "AliasReport"
               Height          =   315
               Index           =   2
               Left            =   1230
               MaxLength       =   200
               TabIndex        =   20
               Tag             =   "Design"
               Top             =   285
               Width           =   3555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alias"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   4
               Left            =   165
               TabIndex        =   29
               Top             =   345
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   3
               Left            =   165
               TabIndex        =   28
               Top             =   1005
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Report Name"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   2
               Left            =   165
               TabIndex        =   27
               Top             =   675
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Module"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   26
               Top             =   1320
               Width           =   510
            End
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   1860
            Top             =   2700
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   16777215
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   9
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":42FCB
                  Key             =   "Orang"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":43B9F
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":6D1F9
                  Key             =   "person1"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":6DAD5
                  Key             =   "person2"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":6E3B1
                  Key             =   "TOP"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":6F205
                  Key             =   "Dept"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":6FCD1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":76533
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmReport.frx":7CD95
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TVConfig 
            Height          =   2700
            Left            =   120
            TabIndex        =   1
            Top             =   855
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   4763
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            SingleSel       =   -1  'True
            ImageList       =   "ImageList1"
            BorderStyle     =   1
            Appearance      =   0
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
         Begin MSComctlLib.ListView ListFilter 
            Height          =   2715
            Left            =   2760
            TabIndex        =   2
            Top             =   840
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   4789
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
            Appearance      =   0
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
               Text            =   "="
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
         Begin VB.Label Label4 
            BackColor       =   &H000000FF&
            DataField       =   "noIDX"
            Height          =   255
            Left            =   7575
            TabIndex        =   55
            Top             =   6885
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000C&
            Caption         =   "MODULE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   2505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            Caption         =   "FILTER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   30
            Top             =   600
            Width           =   7065
         End
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblReportIndex 
      Caption         =   "0"
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents rsReport As ADODB.Recordset
Attribute rsReport.VB_VarHelpID = -1

Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Private Rc As New DBQuick

Private RcFilter As New DBQuick
Dim obj As Object
Dim AliasReport As String
Private RcFlt As New DBQuick

Private Midex As String
Dim rsTemp As New DBQuick

Private IdxOpt As Integer

Private RcProses As New DBQuick
Dim Posisi As String

Private StrLaporan As String
Dim CNN2 As ADODB.Connection
Dim ReportId As String
Dim sViewObject As String
Dim rsDisain As New DBQuick
Dim rsREC As New DBQuick

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

Private TipeFld As TypeFld
Dim mVarTmp As String
Dim rNode As Node
Dim pesan

Private Sub CmbOperator_Click()

    Select Case CmbOperator.Text

        Case "is between"

            If (Trim(ListFilter.SelectedItem.SubItems(4)) = "135" Or Trim(ListFilter.SelectedItem.SubItems(4)) = "0") Then
                ActivateObject oDTPicker, True
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "6") Then
                ActivateObject oTextBx, True
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "202") Then
                ActivateObject oComboBx, True
            End If

        Case "is not between"

            If (Trim(ListFilter.SelectedItem.SubItems(4)) = "135" Or Trim(ListFilter.SelectedItem.SubItems(4)) = "0") Then
                ActivateObject oDTPicker, True
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "6") Then
                ActivateObject oTextBx, True
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "202") Then
                ActivateObject oComboBx, True
            End If

        Case Else

            If (Trim(ListFilter.SelectedItem.SubItems(4)) = "135" Or Trim(ListFilter.SelectedItem.SubItems(4)) = "0") Then
                ActivateObject oDTPicker, False
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "202") Then
                ActivateObject oComboBx, False
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "6") Then
                ActivateObject oTextBx, False
            End If

    End Select

End Sub

Private Sub CmdCancel_Click(Index As Integer)
    PictMain.Visible = False
End Sub

Private Sub cmdLink_Click()
    On Error GoTo RepERR

    With dialog
        .InitDir = ReportPos  'App.Path & "\Report"
        .Filter = "*.rpt|*.rpt" '"Crystal Report"
        .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
        .ShowOpen

        If .Filename = "" Then
        Else
            txtBox(0) = .FileTitle
        End If

    End With
RepERR:


    If Err <> 0 Then
        MessageBox Err.Description & " - " & Err.Number, App.ProductName, msgOkOnly, msgExclamation
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

Private Sub cmdOk_Click(Index As Integer)

    If Index = 0 Then

        '    Debug.Print CmbOperator.ListIndex
        If CmbOperator.ListIndex <> -1 Then

            Select Case ListFilter.SelectedItem.SubItems(4)

                Case 0, 135 'DATE PICKER

                    If Trim(ListFilter.SelectedItem.SubItems(4)) = "0" Then
                        If DTPickFilter(0).Value = "" Or DTPickFilter(1).Value = "" Then
                            MessageBox "Data Belum Lengkap!!"
                            Exit Sub
                        End If

                        ActivateObject oDTPicker, False
                        ListFilter.SelectedItem.SubItems(5) = GetOperator(CmbOperator.ListIndex)
                        ListFilter.SelectedItem.SubItems(1) = DTPickFilter(0).Value
                        ListFilter.SelectedItem.SubItems(2) = DTPickFilter(1).Value
                    ElseIf Trim(ListFilter.SelectedItem.SubItems(4)) = "135" Then
                        ActivateObject oDTPicker, True
                        ListFilter.SelectedItem.SubItems(5) = GetOperator(CmbOperator.ListIndex)
                        ListFilter.SelectedItem.SubItems(1) = DTPickFilter(0).Value
                        ListFilter.SelectedItem.SubItems(2) = DTPickFilter(1).Value
                    End If

                Case 202 ' BERARTI COMBO

'                    If CmbFilter(0).ListIndex = -1 Or CmbFilter(1).ListIndex = -1 And CmbFilter(1).Visible = True Then
'                        MessageBox "Data Belum Lengkap!!"
'                        Exit Sub
'                    End If

                    SetKeListView GetOperator(CmbOperator.ListIndex)
                    ActivateObject oComboBx, False
                    ListFilter.SelectedItem.SubItems(5) = GetOperator(CmbOperator.ListIndex)
                Case 3, 6 ' Textbox iki
                    ActivateObject oTextBx, False
                    ListFilter.SelectedItem.SubItems(5) = GetOperator(CmbOperator.ListIndex)
                    SetKeListView GetOperator(CmbOperator.ListIndex)
            End Select
            PictMain.Visible = False
        Else
            MessageBox "Select Operator", App.ProductName, msgOkOnly, msgExclamation
        End If

    Else
        PictMain.Visible = False
        ListFilter.SelectedItem.Checked = False
    End If

End Sub

Function SetKeListView(operatorNya As String) As String

    Select Case operatorNya

        Case "BETWEEN"

            If (Trim(ListFilter.SelectedItem.SubItems(4)) = "202") Then
                ListFilter.SelectedItem.SubItems(1) = CmbFilter(0).Text
                ListFilter.SelectedItem.SubItems(2) = CmbFilter(1).Text
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "6") Then
                ListFilter.SelectedItem.SubItems(1) = TxtFilter(0).Text
                ListFilter.SelectedItem.SubItems(2) = TxtFilter(1).Text
            End If

        Case "NOT BETWEEN"

            If (Trim(ListFilter.SelectedItem.SubItems(4)) = "202") Then
                ListFilter.SelectedItem.SubItems(1) = CmbFilter(0).Text
                ListFilter.SelectedItem.SubItems(2) = CmbFilter(1).Text
            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "6") Then
                ListFilter.SelectedItem.SubItems(1) = TxtFilter(0).Text
                ListFilter.SelectedItem.SubItems(2) = TxtFilter(1).Text
            End If

        Case Else
            Select Case ListFilter.SelectedItem.SubItems(4)
                Case "202"
                    ListFilter.SelectedItem.SubItems(3) = CmbFilter(0).Text
                Case "3", "6"
                    ListFilter.SelectedItem.SubItems(3) = TxtFilter(0).Text
                Case Else
            End Select
'            If (Trim(ListFilter.SelectedItem.SubItems(4)) = "202") Then
'                ListFilter.SelectedItem.SubItems(3) = CmbFilter(0).Text
'            ElseIf (Trim(ListFilter.SelectedItem.SubItems(4)) = "6") Then
'                ListFilter.SelectedItem.SubItems(3) = TxtFilter(0).Text
'            End If

    End Select

End Function

Private Sub CmdTombol_Click(Index As Integer)
Dim bkMark As Variant
    Select Case Index

        Case 0
            Unload Me
            Exit Sub

        Case 1
            'PREVIEW
            Screen.MousePointer = vbHourglass
            CallReport
            Screen.MousePointer = 0

        Case 2
            'NEW
      
            SSTab1.Tab = 1
            'BindControlToData
            Posisi = "new"
            TombolEdit True
            txtBox(0) = ""
            txtBox(1) = ""
            LblView.Caption = ""
            CmbModule.Text = ""
            CmdTombol(2).Enabled = False
            CmdTombol(6).Enabled = False
            CmdTombol(3).Enabled = True
            'DgDesign.Enabled = False
        Case 3
            '        Save
            
            SSTab1.TabEnabled(0) = True
            DgDesign.Enabled = True
            TombolEdit False
            ExecSave
            display_data
            
            CmdTombol(3).Enabled = False
            CmdTombol(6).Enabled = True
            CmdTombol(2).Enabled = True
            SSTab1.Tab = 0
        Case 4  'DELETE
            If GridReport.Columns.Count > 1 Then
                If SSTab1.Tab = 1 Then
                    If MessageBox("Yakin data akan dihapus ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
                       SendDataToServer "delete from [report modules] where noIdx = '" & Label3.Caption & "'"
                       DelReportPermit
                    End If
                    display_data
                Else
                    If MessageBox("Yakin data akan dihapus ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
                        SendDataToServer "DELETE FROM [Report Permit] Where (NoIdx = '" & Label4.Caption & "')"
                        DelreportModules
                    End If
                End If
            End If
        
        
            If tvConfig.Nodes.Count > 0 Then OpenDB2 tvConfig.Nodes(1).Text
        Case 5  'BATAL
            SSTab1.TabEnabled(0) = True
            DgDesign.Enabled = True
            BindControlToData
            GenerateFld
            TombolEdit False
            CmdTombol(6).Enabled = True
            CmdTombol(3).Enabled = False
            CmdTombol(2).Enabled = True
            SSTab1.Tab = 0
        Case 6
            ' Edit Data
            If Not rNode Is Nothing Then
                bkMark = rsReport.Bookmark
                CmdTombol(2).Enabled = False
                SSTab1.TabEnabled(0) = False
                DgDesign.Enabled = False
                SSTab1.Tab = 1
                Posisi = "edit"
                TombolEdit True
                CmdTombol(2).Enabled = False
                CmdTombol(6).Enabled = False
                CmdTombol(3).Enabled = True
            Else
                MessageBox "Pilih Modul Laporan untuk diedit", "Report Control", msgOkOnly, msgExclamation
            End If
        Case 7
            If Not rNode Is Nothing Then
                OpenDB2 Mid(rNode.Key, 1, Len(rNode.Key) - 1)
                GridReport.ReBind
             Else
                tvConfig.SetFocus
                tvConfig.Nodes(rNode).Selected = True
                TVConfig_NodeClick tvConfig.SelectedItem
            End If
    End Select

End Sub

Private Sub TombolEdit(TombolEdit As Boolean)
    cmdLink.Enabled = TombolEdit
    txtBox(1).Enabled = TombolEdit
    txtBox(3).Enabled = TombolEdit
    CmbModule.Enabled = TombolEdit
    cmdView.Enabled = TombolEdit
End Sub

Private Sub ExecSave()
    Dim obj As Object

    For Each obj In Me

        If TypeOf obj Is OptionButton Then
            If obj.Value = True Then ReportId = obj.ToolTipText
        End If

    Next obj
    AliasReport = Mid(txtBox(0), 1, Len(txtBox(0)) - 4)
    If Posisi = "edit" Then
       If MessageBox("Apakah Data Ini akan Diubah ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
          SendDataToServer "update [report modules] set Description = '" & txtBox(1).Text & "', ReportGroup = '" & CmbModule.BoundText & "', FileNameReport = '" & txtBox(0).Text & "', ViewObject = '" & LblView.Caption & "' , IDreport = '" & ReportId & "' , [Alias Report] = '" & txtBox(3).Text & "' where noIdx = '" & Label3.Caption & "'"
         ' MessageBox "Ubah Module Laporan Sukses", "Konfirmasi", msgOkOnly, msgInfo
       End If
    Else
       If MessageBox("Apakah Data ini akan disimpan ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
          SendDataToServer "insert into [report modules] (Description, ReportGroup, FileNameReport, ViewObject, IDreport, [Alias Report]) values ('" & txtBox(1).Text & "', '" & CmbModule.BoundText & "', '" & txtBox(0).Text & "', '" & LblView.Caption & "', '" & ReportId & "', '" & txtBox(3).Text & "')"
          InsReportPermit
         ' MessageBox "Simpan Module Laporan Sukses", "Konfirmasi", msgOkOnly, msgInfo
       End If
    End If

End Sub

Private Sub cmdView_Click()
    OpenPartner 0
End Sub

Private Sub DgDesign_DblClick()
    'CallReport
End Sub

Private Sub DgFilter_Error(ByVal DataError As Integer, _
                           Response As Integer)
    DataError = 0
    Response = 0
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

Private Sub DgReport_DblClick()
    'Dim mVarWindState As Integer
    'mVarWindState = Me.WindowState
    CallReport
    'Me.WindowState = mVarWindState
End Sub

Private Sub DgDesign_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If SSTab1.Tab = 1 Then rsReport.AbsolutePosition = rsDisain.AbsolutePosition
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    'ScanKey KeyCode, Shift, MyDDE

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    
    
    '*** Setting Layout form
    HiasFormManTell Picture2, Me
    Me.Tag = "awal"
    CenterForm Picture2, Me
    SSTab1.BackColor = Picture2.BackColor
    SSTab1.Tab = 0
    SSTab1.TabVisible(1) = False
    IdxOpt = 0
    GridLayout
    OpenDBEdit ""
    
    
    
    '*** Load data group / module di treeview di tab report
    LoadTree
    
    
    
    '*** Load data group / module untuk data combo di tab design
    rsTemp.DBOpen "SELECT GroupID, GroupName FROM [Report Group]", CNN
    Set CmbModule.RowSource = rsTemp.DBRecordset
    CmbModule.ListField = rsTemp.DBRecordset.Fields(1).Name
    CmbModule.BoundColumn = rsTemp.DBRecordset.Fields(0).Name
    If tvConfig.Nodes.Count > 0 Then OpenDB2 tvConfig.Nodes(1).Text
    
    
    '*** Load data report untuk grid dgDesign
    display_data
    
    
    CmdTombol(3).Enabled = False
    Label2(0).FontBold = True
    Label2(1).FontBold = True
    Label2(0).BackColor = &H400000
    Label2(1).BackColor = &H400000
    Label2(0).ForeColor = &H80000005
    Label2(1).ForeColor = &H80000005

End Sub

Function display_data()
   rsREC.DBOpen "select * from [report modules] order by reportGroup,FileNameReport,noIDX desc", CNN, lckLockBatch
   Set DgDesign.DataSource = rsREC.DBRecordset
   Set txtBox(0).DataSource = rsREC.DBRecordset
   Set txtBox(1).DataSource = rsREC.DBRecordset
   Set txtBox(3).DataSource = rsREC.DBRecordset
   Set CmbModule.DataSource = rsREC.DBRecordset
   Set LblView.DataSource = rsREC.DBRecordset
   Set Label3.DataSource = rsREC.DBRecordset
   DgDesign.Columns(0).Visible = False
End Function

Private Sub LoadTree()
On Error GoTo 1
    Dim vNode As Node
    Dim rsForms As DBQuick
    Dim No  As Integer
    Dim strProvider As String
    Dim FirstNode As String
    tvConfig.Nodes.Clear
    
    Set rsForms = New DBQuick
    rsForms.DBOpen "SELECT GroupID, GroupName FROM [Report Group]", CNN, lckLockReadOnly
    No = 1

    If rsForms.Recordcount > 0 Then
        rsForms.DBRecordset.MoveFirst
        FirstNode = Trim(rsForms.DBRecordset.Fields(0))
        While Not rsForms.DBRecordset.EOF
            Set vNode = tvConfig.Nodes.Add(, , CStr(rsForms.DBRecordset.Fields(0)) & "A", Trim(IIf(IsNull(rsForms.DBRecordset.Fields("GroupName").Value), " ", _
            rsForms.DBRecordset.Fields("GroupName").Value)), 7)
            vNode.Tag = No
            vNode.Expanded = True
            No = No + 1
            rsForms.DBRecordset.MoveNext
        Wend
    End If
Exit Sub
1:
   MessageBox Err.Description, "frmReport : LoadTree", msgOkOnly, msgExclamation
End Sub

Private Sub ActivateObject(sObj As FieldControlType, _
                           bStatus As Boolean)
    Dim I As Integer
    Dim strSQL As String
    

    '    oTextBx
    '    oComboBx
    '    oMaskBx
    '    fcCheckBx
    '    oDTPicker
    '
    sViewObject = GridReport.Columns(1).Text

    For I = 0 To 1
        TxtFilter(I).Visible = False
        'TxtFilter(I).Text = ""
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
                MessageBox "Object Table Not Defined"
            Else
                strSQL = "select Distinct [" & Trim(ListFilter.SelectedItem.Text) & "] from [" & sViewObject & "] where [" & ListFilter.SelectedItem.Text & "] is not null "
'                Debug.Print strSQL
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

Private Sub GridReport_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim vLst As ListItem
'Set Rc = New DBQuick
If SSTab1.Tab = 0 Then
    ListFilter.ListItems.Clear
    
    rsREC.MakeFind "noIDX = '" & Label4.Caption & "'"
    rsDisain.AbsolutePosition = rsReport.AbsolutePosition
    If GridReport.Tag = "0" Then Exit Sub
    If rsReport.Recordcount < 1 Then Exit Sub
    If Rc.DBOpen("SELECT Report_Filter.REPORT_ID, Report_Filter.FIELD_NAME, Report_Filter.FIELD_TYPE, " & _
            " Report_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Report_Filter ON ([Report Modules].NoIdx = Report_Filter.REPORT_ID) " & _
            " WHERE [Report Modules].NoIdx = '" & GridReport.Columns(0).Text & "' AND [Report Modules].ViewObject = '" & GridReport.Columns(1).Text & "'", CNN, lckLockBatch) = True Then
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
End If
End Sub

Private Sub ListFilter_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked = False Then
        PictMain.Visible = False
        ListFilter.SelectedItem.SubItems(1) = ""
        ListFilter.SelectedItem.SubItems(2) = ""
        ListFilter.SelectedItem.SubItems(3) = ""
        ListFilter.SelectedItem.SubItems(5) = ""
        
    Else
        PictMain.Visible = True
        CmbFilter(0).Clear
        CmbOperator.ListIndex = -1
        Set ListFilter.SelectedItem = Item
        'LBLFilter = _
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
            Case 0, 135   'DATEPICKER
                ActivateObject oDTPicker, False
        End Select

    End If

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

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
   CmdTombol(1).Enabled = False
Else
   CmdTombol(1).Enabled = True
End If
End Sub

Private Sub TVConfig_NodeClick(ByVal Node As MSComctlLib.Node)
    Me.Tag = ""
    Set rNode = Node
    Select Case Node.Text
        Case Node.Key
            '*** Menampilkan data di gridReport
            OpenDB2 Mid(Node.Key, 1, Len(Node.Key) - 1)
            GridReport.ReBind

        Case Else
            OpenDB2 Mid(Node.Key, 1, Len(Node.Key) - 1)
            GridReport.ReBind
    End Select

End Sub

Private Sub OpenDB2(ByVal Nodetext As String)
    Dim I As Long
    Dim strSQL As String
  
   'strSQL = "SELECT * FROM [Report Modules] WHERE (ReportGroup = N'" & Nodetext & "')"
'    strSQL = " SELECT [report permit].[User ID],[report permit].noidx, [report modules].Description, [report modules].[Alias Report]," & _
'                     " [report modules].ReportGroup , [report modules].FileNameReport, [report modules].ViewObject, [report permit].Laporan " & _
'                     " FROM dbo.[report permit] INNER JOIN " & _
'                     " [report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx " & _
'                     " Where (dbo.[report permit].[User ID] =" & aksess.GetID & ") and (dbo.[report permit].laporan <> 0) and [report modules].ReportGroup= N'" & Nodetext & "'" & _
'                     " ORDER BY [report modules].[Alias Report]"
'    Debug.Print strSQL

    strSQL = "SELECT TOP (100) PERCENT dbo.[report permit].noidx, dbo.[report permit].Laporan, dbo.user_table_group.[Group Name], dbo.[report permit].IDGroup," & _
             "dbo.[report modules].ReportGroup , dbo.[report modules].[Alias Report],[report modules].Description,[report modules].ViewObject,[report modules].FileNameReport " & _
             "FROM dbo.[report permit] INNER JOIN " & _
             "dbo.[report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx RIGHT OUTER JOIN " & _
             "dbo.user_table_group ON dbo.[report permit].IDGroup = dbo.user_table_group.id " & _
             "WHERE (dbo.[report modules].ReportGroup = N'" & Nodetext & "')  and (dbo.[report permit].IDGroup=" & GroupID & ")" & _
             "ORDER BY dbo.[report modules].[Alias Report]"


    Set rsReport = New ADODB.Recordset
    rsReport.CursorLocation = adUseClient
    rsReport.Open strSQL, CNN, adOpenKeyset, adLockOptimistic
    Set GridReport.DataSource = rsReport
    Set Label4.DataSource = rsReport
    GridReport.ReBind
  
    If Me.Tag = "awal" Then
        If tvConfig.Nodes.Count > 0 Then
            'strSQL = "SELECT * FROM [REPORT MODULES] WHERE REPORTGROUP = '" & TVConfig.Nodes(1).Text & "'"
            
            strSQL = " SELECT [report permit].[User ID],[report permit].noidx, [report modules].Description, [report modules].[Alias Report]," & _
                     " [report modules].ReportGroup , [report modules].FileNameReport, [report modules].ViewObject, [report permit].Laporan " & _
                     " FROM dbo.[report permit] INNER JOIN " & _
                     " [report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx " & _
                     " Where (dbo.[report permit].[User ID] = " & aksess.GetID & ") and (dbo.[report permit].laporan <> 0) and [report modules].ReportGroup= '" & tvConfig.Nodes(1).Text & "'" & _
                     " ORDER BY [report modules].[Alias Report]"
            
        Else
            'strSQL = "SELECT * FROM [REPORT MODULES] WHERE REPORTGROUP = '" & "" & "'"
             strSQL = " SELECT [report permit].[User ID],[report permit].noidx, [report modules].Description, [report modules].[Alias Report]," & _
                     " [report modules].ReportGroup , [report modules].FileNameReport, [report modules].ViewObject, [report permit].Laporan " & _
                     " FROM dbo.[report permit] INNER JOIN " & _
                     " [report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx " & _
                     " Where (dbo.[report permit].[User ID] = " & aksess.GetID & ") and (dbo.[report permit].laporan <> 0) and [report modules].ReportGroup= '" & "" & "'" & _
                     " ORDER BY [report modules].[Alias Report]"
        End If

    Else
       'strSQL = "SELECT * FROM [REPORT MODULES] WHERE REPORTGROUP = '" & TVConfig.SelectedItem.Text & "'"
       
        strSQL = " SELECT [report permit].[User ID],[report permit].noidx, [report modules].Description, [report modules].[Alias Report]," & _
                     " [report modules].ReportGroup , [report modules].FileNameReport, [report modules].ViewObject, [report permit].Laporan " & _
                     " FROM dbo.[report permit] INNER JOIN " & _
                     " [report modules] ON dbo.[report permit].noidx = dbo.[report modules].NoIdx " & _
                     " Where (dbo.[report permit].[User ID] = " & aksess.GetID & ") and (dbo.[report permit].laporan <> 0) and [report modules].ReportGroup= '" & tvConfig.SelectedItem.Text & "'" & _
                     " ORDER BY [report modules].[Alias Report]"
    End If
'    Debug.Print strSQL
    Set rsDisain = New DBQuick
    rsDisain.DBOpen strSQL, CNN, lckLockBatch
    Set DgDesign.DataSource = rsDisain.DBRecordset
'    rsDisain.CloseDB
    If rsReport.Recordcount < 1 Then Exit Sub
    lblReportIndex.Caption = GridReport.Columns(0).Text
    BindControlToData
End Sub

Function BindControlToData()
    Set txtBox(0).DataSource = rsDisain.DataSource
    'txtBox(0).DataField = "Alias Report"
    
    Set txtBox(1).DataSource = rsDisain.DataSource
    'txtBox(0).DataField = "Description"
    
    Set LblView.DataSource = rsDisain.DataSource
    'LblView.DataField = "viewobject"

    Set CmbModule.DataSource = rsDisain.DataSource

    If Not rsDisain.DBRecordset.EOF Or Not rsDisain.DBRecordset.BOF Then CmbModule.Text = rsDisain.Fields("ReportGroup")
End Function

Private Sub OpenFlt()

    If Not rsReport Is Nothing Then
        If rsReport.State = 1 Then
            If rsReport.Recordcount <> 0 And Len(sViewObject) <> 0 Then
                RcFlt.DBOpen " select [" & RcFilter.Fields(0) & "] from [" & sViewObject & "] group By [" & RcFilter.Fields(0) & "] order by [" & RcFilter.Fields(0) & "]", CNN, lckLockReadOnly
               ' cboFilter.Text = ""
               ' cboFilter.ListField = RcFilter.Fields(0)
               ' Set cboFilter.RowSource = RcFlt.DBRecordset
            Else
               ' cboFilter.Text = ""
               ' Set cboFilter.RowSource = Nothing
            End If
        End If
    End If

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

Private Sub OpenDBEdit(ByVal Nodetext As String)
    Dim I As Long
    Dim rsReportEdit As ADODB.Recordset
    Nodetext = "Accounting"
    'strSQL = "SELECT * FROM [Report Modules] WHERE (ReportGroup = N'" & Nodetext & "')"
    Set rsReportEdit = New ADODB.Recordset
    rsReportEdit.CursorLocation = adUseClient
    ' rsReportEdit.Open strSQL, CNN, adOpenKeyset, adLockOptimistic
 
End Sub

Private Function CreateIdx() As String
    Dim RcIdx As New DBQuick
    Dim mVarNo As Integer
    RcIdx.DBOpen "SELECT MAX(RIGHT(IDReport, 5)) AS MaxNo FROM [Report Modules]", CNN, lckLockReadOnly

    With RcIdx

        If .Recordcount <> 0 Then
            mVarNo = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        Else
            mVarNo = 0
        End If

        mVarNo = mVarNo + 1

        Select Case Len(Trim(Str(mVarNo)))

            Case 1
                CreateIdx = "0000" & Trim(Str(mVarNo))

            Case 2
                CreateIdx = "000" & Trim(Str(mVarNo))

            Case 3
                CreateIdx = "00" & Trim(Str(mVarNo))

            Case 4
                CreateIdx = "0" & Trim(Str(mVarNo))

            Case 5
                CreateIdx = Trim(Str(mVarNo))
        End Select

    End With

    RcIdx.CloseDB
End Function
Private Function GetPeriodFilter() As Boolean
On Error GoTo GetErr
Dim I As Long
For I = 1 To ListFilter.ListItems.Count
    If ListFilter.ListItems(I).Text = "PERIODE" Then
        GetPeriodFilter = True
        Exit For
    Else
        GetPeriodFilter = False
    End If
Next
Exit Function
GetErr:
    MessageBox Err.Description, "frmReport : GetPeriodReport", msgOkOnly, msgCrtical
End Function
Private Function GetPeriodValue() As Boolean
On Error GoTo GetErr
Dim I As Long
For I = 1 To ListFilter.ListItems.Count
    If ListFilter.ListItems(I).Text = "PERIODE" Then
        If ListFilter.ListItems(I).Checked = True Then
            mVarTempPeriode = Val(ListFilter.ListItems(I).SubItems(3))
            GetPeriodValue = True
        Else
            GetPeriodValue = False
        End If
        Exit For
    Else
        GetPeriodValue = False
    End If
Next
Exit Function
GetErr:
    MessageBox Err.Description, "frmReport : GetPeriodReport", msgOkOnly, msgCrtical
End Function
Private Sub CallReport()

On Error GoTo Hell
Dim mPer As Integer
Dim Mprint As New frmReportView
Dim RcTes As New DBQuick

If SeekForm(GridReport.Columns(2)) = True Then Exit Sub
Select Case UCase(GridReport.Columns(4))
       Case "ACCNERACABULANAN.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = LocalPeriodeActive
                    IsiLabaRugi mVarTempPeriode
                    StrLaporan = " SELECT GLAccount.NoAccount AS [Kode Akun], GLAccount.AccountName, [Tabel Pembantu].CurrentDR" & mPer & " AS [Saldo Awal DR], [Tabel Pembantu].CurrentCR" & mPer & " AS [Saldo Awal CR], " & _
                                " SUM(ISNULL([Detail Journal].Debet, 0)) AS [Current DR], SUM(ISNULL([Detail Journal].Credit, 0))  AS [Current CR], GLAccount.GroupAccount AS [Level I], GLAccount_1.AccountName AS [Group I], " & _
                                " GLAccount_1.GroupAccount AS [Level II], GLAccount_2.AccountName AS [Group II], GLAccount_2.GroupAccount AS [Level III], GLAccount_3.AccountName AS [Group III], " & _
                                " GLAccount_3.GroupAccount AS [Level IV], GLAccount_4.AccountName AS [Group IV], [Table Journal].Periode, AccType.ID, [Tabel Pembantu].CurrentDR" & mVarTempPeriode & " AS [RL DR], " & _
                                " [Tabel Pembantu].CurrentCR" & mVarTempPeriode & " AS [RL CR], [Tabel Pembantu].[Seting Relasi] " & _
                                " FROM [Table Journal] INNER JOIN [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID RIGHT OUTER JOIN GLAccount LEFT OUTER JOIN GLAccount GLAccount_1 " & _
                                " INNER JOIN GLAccount GLAccount_2 ON GLAccount_1.GroupAccount = GLAccount_2.NoAccount INNER JOIN GLAccount GLAccount_3 ON GLAccount_2.GroupAccount = GLAccount_3.NoAccount " & _
                                " INNER JOIN GLAccount GLAccount_4 ON GLAccount_3.GroupAccount = GLAccount_4.NoAccount INNER JOIN AccType ON GLAccount_1.Type = AccType.Tipe ON " & _
                                " GLAccount.GroupAccount = GLAccount_1.NoAccount ON  [Detail Journal].NoAccount = GLAccount.NoAccount LEFT OUTER JOIN [Tabel Pembantu] ON " & _
                                " GLAccount.NoAccount = [Tabel Pembantu].NoAccount " & _
                                " WHERE (GLAccount.[Group] = N'Detail List Account') AND ([Tabel Pembantu].[Kelompok Perkiraan] = 1) GROUP BY GLAccount.NoAccount, GLAccount.AccountName, GLAccount.GroupAccount, " & _
                                " GLAccount_1.AccountName, GLAccount_1.GroupAccount, GLAccount_2.AccountName, GLAccount_2.GroupAccount, GLAccount_3.AccountName, GLAccount_3.GroupAccount, GLAccount_4.AccountName, " & _
                                " [Table Journal].Periode, AccType.ID, [Tabel Pembantu].CurrentDR" & mPer & ", [Tabel Pembantu].CurrentCR" & mPer & ", [Tabel Pembantu].CurrentDR" & mVarTempPeriode & ",  [Tabel Pembantu].CurrentCR" & mVarTempPeriode & ", [Tabel Pembantu].[Seting Relasi] " & _
                                " HAVING (NOT (GLAccount_1.AccountName IS NULL)) AND ([Table Journal].Periode = " & mVarTempPeriode & ") OR ([Table Journal].Periode IS NULL) ORDER BY GLAccount.NoAccount"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If
            
       Case "ACCRUGILABA.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = LocalPeriodeActive
                    StrLaporan = " SELECT GLAccount.NoAccount AS [Kode Akun], GLAccount.AccountName, [Tabel Pembantu].CurrentDR" & mPer & " AS [Saldo Awal DR],                       [Tabel Pembantu].CurrentCR" & mPer & " AS [Saldo Awal CR], SUM(ISNULL([Detail Journal].Debet, 0)) AS [Current DR], SUM(ISNULL([Detail Journal].Credit, 0))                       AS [Current CR], GLAccount.GroupAccount AS [Level I], GLAccount_1.AccountName AS [[Group]] I], GLAccount_1.GroupAccount AS [Level II],                        GLAccount_2.AccountName AS [Group II], GLAccount_2.GroupAccount AS [Level III], GLAccount_3.AccountName AS [Group III],                        GLAccount_3.GroupAccount AS [Level IV]]], GLAccount_4.AccountName AS [Group IV], [Table Journal].Periode, AccType.ID FROM         [Table Journal] INNER JOIN                      [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID RIGHT OUTER JOIN" & _
                                 " GLAccount LEFT OUTER JOIN GLAccount GLAccount_1 INNER JOIN GLAccount GLAccount_2 ON GLAccount_1.GroupAccount = GLAccount_2.NoAccount INNER JOIN                      GLAccount GLAccount_3 ON GLAccount_2.GroupAccount = GLAccount_3.NoAccount INNER JOIN                      GLAccount GLAccount_4 ON GLAccount_3.GroupAccount = GLAccount_4.NoAccount INNER JOIN                      AccType ON GLAccount_1.Type = AccType.Tipe ON GLAccount.GroupAccount = GLAccount_1.NoAccount ON                       [Detail Journal].NoAccount = GLAccount.NoAccount LEFT OUTER JOIN [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount" & _
                                 " WHERE (GLAccount.[Group] = N'Detail List Account') AND ([Tabel Pembantu].[Kelompok Perkiraan] = 0) GROUP BY GLAccount.NoAccount, GLAccount.AccountName, GLAccount.GroupAccount, GLAccount_1.AccountName, GLAccount_1.GroupAccount,                       GLAccount_2.AccountName, GLAccount_2.GroupAccount, GLAccount_3.AccountName, GLAccount_3.GroupAccount, GLAccount_4.AccountName,                       [Table Journal].Periode, AccType.ID, [Tabel Pembantu].CurrentDR" & mPer & ", [Tabel Pembantu].CurrentCR" & mPer & " HAVING      (NOT (GLAccount_1.AccountName IS NULL)) AND ([Table Journal].Periode = " & mVarTempPeriode & ") OR  ([Table Journal].Periode IS NULL) ORDER BY GLAccount.NoAccount"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If
       Case "ACCPERUBAHANMODAL.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = LocalPeriodeActive
                    StrLaporan = " SELECT    GLAccount.NoAccount AS [Kode Akun], GLAccount.AccountName, [Tabel Pembantu].CurrentDR" & mPer & " AS [Saldo Awal DR],                       [Tabel Pembantu].CurrentCR2 AS [Saldo Awal CR], SUM(ISNULL([Detail Journal].Debet, 0)) AS [Current DR], SUM(ISNULL([Detail Journal].Credit, 0))                       AS [Current CR], GLAccount.GroupAccount AS [Level I], GLAccount_1.AccountName AS [[Group]] I], GLAccount_1.GroupAccount AS [Level II],                       GLAccount_2.AccountName AS [Group II], GLAccount_2.GroupAccount AS [Level III], GLAccount_3.AccountName AS [Group III],                       GLAccount_3.GroupAccount AS [Level IV]]], GLAccount_4.AccountName AS [Group IV], [Table Journal].Periode, AccType.ID,                       [Tabel Pembantu].[Seting Relasi], [Tabel Pembantu].CurrentDR" & mVarTempPeriode & " AS [RL DR], [Tabel Pembantu].CurrentCR" & mVarTempPeriode & " AS [RL CR]" & _
                                 " FROM         [Table Journal] INNER JOIN                      [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID RIGHT OUTER JOIN                      GLAccount LEFT OUTER JOIN                      GLAccount GLAccount_1 INNER JOIN                      GLAccount GLAccount_2 ON GLAccount_1.GroupAccount = GLAccount_2.NoAccount INNER JOIN                      GLAccount GLAccount_3 ON GLAccount_2.GroupAccount = GLAccount_3.NoAccount INNER JOIN                      GLAccount GLAccount_4 ON GLAccount_3.GroupAccount = GLAccount_4.NoAccount INNER JOIN                      AccType ON GLAccount_1.Type = AccType.Tipe ON GLAccount.GroupAccount = GLAccount_1.NoAccount ON                       [Detail Journal].NoAccount = GLAccount.NoAccount LEFT OUTER JOIN                      [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount" & _
                                 " WHERE     (GLAccount.[Group] = N'Detail List Account') AND ([Tabel Pembantu].[Kelompok Modal] = 1) GROUP BY GLAccount.NoAccount, GLAccount.AccountName, GLAccount.GroupAccount, GLAccount_1.AccountName, GLAccount_1.GroupAccount,                        GLAccount_2.AccountName, GLAccount_2.GroupAccount, GLAccount_3.AccountName, GLAccount_3.GroupAccount, GLAccount_4.AccountName,                        [Table Journal].Periode, AccType.ID, [Tabel Pembantu].CurrentDR" & mPer & ", [Tabel Pembantu].CurrentCR" & mPer & ", [Tabel Pembantu].[Seting Relasi],                        [Tabel Pembantu].CurrentDR" & mVarTempPeriode & ", [Tabel Pembantu].CurrentCR" & mVarTempPeriode & " HAVING      (NOT (GLAccount_1.AccountName IS NULL)) AND ([Table Journal].Periode = " & mVarTempPeriode & ") OR                      ([Table Journal].Periode IS NULL) ORDER BY GLAccount.NoAccount"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If
'            Debug.Print StrLaporan
       Case "ACCNERACASALDO.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = LocalPeriodeActive
                    StrLaporan = " SELECT [Tabel Pembantu].NoAccount AS [Kode Akun], GLAccount.AccountName AS [Nama Akun], [Tabel Pembantu].CurrentDR" & mPer & " AS [Saldo Awal DR],[Tabel Pembantu].CurrentCR" & mPer & " AS [Saldo Awal CR], ISNULL(SUM([Detail Journal].Debet), 0) AS [Current DR], ISNULL(SUM([Detail Journal].Credit), 0) AS [Current CR], ISNULL([Table Journal].Periode, " & mVarTempPeriode & ") AS Periode, AccType.ID, [Tabel Pembantu].[Kelompok Perkiraan]" & _
                                 " FROM AccType INNER JOIN GLAccount ON AccType.Tipe = GLAccount.Type LEFT OUTER JOIN [Detail Journal] INNER JOIN [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID ON GLAccount.NoAccount = [Detail Journal].NoAccount LEFT OUTER JOIN [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     ([Tabel Pembantu].[Seting Relasi] = 0) AND (GLAccount.[Group] = N'Detail List Account')" & _
                                 " GROUP BY [Tabel Pembantu].NoAccount, GLAccount.AccountName, [Tabel Pembantu].CurrentDR" & mPer & ", [Tabel Pembantu].CurrentCR" & mPer & ", AccType.ID,[Tabel Pembantu].[Kelompok Perkiraan], ISNULL([Table Journal].Periode, " & mVarTempPeriode & ") HAVING      (ISNULL([Table Journal].Periode, " & mVarTempPeriode & ") = " & mVarTempPeriode & ") OR (ISNULL([Table Journal].Periode, " & mVarTempPeriode & ") IS NULL) ORDER BY [Tabel Pembantu].NoAccount"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If
            
       Case "ACCARUSKAS.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = mVarTempPeriode
                    StrLaporan = " SELECT     GLAccount.NoAccount AS [Kode Akun], GLAccount.AccountName, [Tabel Pembantu].CurrentDR" & mPer & " AS [Saldo Awal DR],                       [Tabel Pembantu].CurrentCR" & mPer & " AS [Saldo Awal CR], SUM(ISNULL([Detail Journal].Debet, 0)) AS [Current DR], SUM(ISNULL([Detail Journal].Credit, 0))                       AS [Current CR], GLAccount.GroupAccount AS [Level I], GLAccount_1.AccountName AS [[Group]] I], GLAccount_1.GroupAccount AS [Level II],                       GLAccount_2.AccountName AS [Group II], GLAccount_2.GroupAccount AS [Level III], GLAccount_3.AccountName AS [Group III],                       GLAccount_3.GroupAccount AS [Level IV]]], GLAccount_4.AccountName AS [Group IV], [Table Journal].Periode, AccType.ID,                       [Tabel Pembantu].[Seting Relasi], [Tabel Pembantu].CurrentDR" & mVarTempPeriode & " AS [RL DR], [Tabel Pembantu].CurrentCR" & mVarTempPeriode & " AS [RL CR],                       [Tabel Pembantu].[Label ArusKas]" & _
                                 " FROM         [Table Journal] INNER JOIN                       [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID RIGHT OUTER JOIN                       GLAccount LEFT OUTER JOIN                       GLAccount GLAccount_1 INNER JOIN                       GLAccount GLAccount_2 ON GLAccount_1.GroupAccount = GLAccount_2.NoAccount INNER JOIN                       GLAccount GLAccount_3 ON GLAccount_2.GroupAccount = GLAccount_3.NoAccount INNER JOIN                       GLAccount GLAccount_4 ON GLAccount_3.GroupAccount = GLAccount_4.NoAccount INNER JOIN                       AccType ON GLAccount_1.Type = AccType.Tipe ON GLAccount.GroupAccount = GLAccount_1.NoAccount ON                        [Detail Journal].NoAccount = GLAccount.NoAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount" & _
                                 " WHERE     (GLAccount.[Group] = N'Detail List Account') AND ([Tabel Pembantu].[Seting ArusKas] = 1) GROUP BY GLAccount.NoAccount, GLAccount.AccountName, GLAccount.GroupAccount, GLAccount_1.AccountName, GLAccount_1.GroupAccount,                        GLAccount_2.AccountName, GLAccount_2.GroupAccount, GLAccount_3.AccountName, GLAccount_3.GroupAccount, GLAccount_4.AccountName,                        [Table Journal].Periode, AccType.ID, [Tabel Pembantu].CurrentDR" & mPer & ", [Tabel Pembantu].CurrentCR" & mPer & ", [Tabel Pembantu].[Seting Relasi],                        [Tabel Pembantu].CurrentDR" & mVarTempPeriode & ", [Tabel Pembantu].CurrentCR" & mVarTempPeriode & ", [Tabel Pembantu].[Label ArusKas] HAVING      (NOT (GLAccount_1.AccountName IS NULL)) AND ([Table Journal].Periode = " & mVarTempPeriode & ") OR                       ([Table Journal].Periode IS NULL) ORDER BY GLAccount.NoAccount"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If

       Case "ACCJOURNAL.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = mVarTempPeriode
                    StrLaporan = " SELECT * from [Accjournal] where  (Periode =" & mVarTempPeriode & ") ORDER BY [Kode Journal]"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If
       Case "ACCBUKUBESAR.RPT":
            If GetPeriodFilter Then
                If GetPeriodValue Then
                    mPer = LocalPeriodeActive
                    StrLaporan = " SELECT     [Table Journal].JournalID AS [No Bukti], [Table Journal].DateTrans AS Tanggal, [Detail Journal].NoAccount AS [Kode akun], " & _
                                 " GLAccount.AccountName AS [Nama Akun], LEFT([Table Journal].RefNotes, 244) AS Keterangan, ISNULL([Tabel Pembantu].CurrentDR" & mPer & ", 0)" & _
                                 " AS [Saldo Awal DR], ISNULL([Tabel Pembantu].CurrentCR" & mPer & ", 0) AS [Saldo Awal CR], [Detail Journal].Debet, [Detail Journal].Credit AS Kredit," & _
                                 " [Table Journal].Periode, [Table Journal].TypeTrans AS [Tipe Trans], [Detail Journal].[No], ISNULL([Table Journal].TransID, '-') AS [Bukti Trans]," & _
                                 " ISNULL([Table Journal].PurchaseID, '-') AS [PO/SC Reff], [Detail Journal].[Doc Reff] AS [Dok Ref], [Table Journal].NoUrut," & _
                                 " [Detail Journal].Keterangan AS [Detail Journal], GLAccount.[Default] AS Posisi" & _
                                 "  FROM GLAccount INNER JOIN [Detail Journal] ON GLAccount.NoAccount = [Detail Journal].NoAccount " & _
                                 " INNER JOIN [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID LEFT OUTER JOIN " & _
                                 " [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount " & _
                                 " WHERE ([Table Journal].periode = " & mVarTempPeriode & ") ORDER BY [Table Journal].JournalID, [Detail Journal].[No], [Detail Journal].NoAccount"
                Else
                    MessageBox "Parameter Period Laporan harus diisi..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                    Exit Sub
                End If
            Else
                MessageBox "Parameter Period Laporan harus diset..", "Kontrol Filter Laporan", msgOkOnly, msgExclamation
                Exit Sub
            End If
       Case Else
'            StrLaporan = " Select * from [" & Replace(UCase(GridReport.Columns(2)), ".RPT", "") & "] " & ScanFilter2

            'FILTER QUERYREPORT BASED ON SELECTED VIEWOBJECT
            StrLaporan = " Select * from [" & UCase(rsReport.Fields(6).Value) & "] " & ScanFilter2
End Select
If mVarTempPeriode > mVarPeriode Then
   MessageBox "Transaksi periode " & mVarTempPeriode & " belum ada.Laporan Batal Ditampilkan.", "Laporan", msgOkOnly
   Exit Sub
End If
'MessageBox StrLaporan
'Debug.Print StrLaporan
RcTes.DBOpen StrLaporan, CNN
If RcTes.Recordcount <> 0 Then
    Select Case UCase(GridReport.Columns(4))
          Case "JOB ISSUES REPORT.RPT":
                CallRPTReport "Job Issues Report.Rpt", "Select * from [Job Issues Report] order by [OrderID]", , , "Select * from [Detail Job Issues Report]", "Detail Job Issues Report"
          Case "SHOP FLOOR.RPT":
                CallRPTReport "Shop Floor.RPT", "Select * from [Shop Floor] ", , , "Select * from [STAGE DETAIL]", "STAGE"

          Case Else:
                With Mprint
                     .QuerySource = StrLaporan
                     '.ReportName = MyDDE.GetFieldByName("FileNameReport")
                     '.ReportTitle = MyDDE.GetFieldByName("AliasReport")
                     '.ReportName = RcTes.Fields("FileNameReport")
                      .ReportName = GridReport.Columns(4)
                     '.ReportTitle = RcTes.Fields("AliasReport")
                     .ReportTitle = GridReport.Columns(2)
                     .SetFocus
                End With
  End Select
Else
'    Debug.Print StrLaporan
   MessageBox "Laporan Belum Ada Datanya. Harap Diperiksa Filter Kriterianya", "Peringatan", msgOkOnly
End If
RcTes.CloseDB
Exit Sub
Hell:
    MessageBox Err.Description, "Laporan-CallReport", msgOkOnly, msgExclamation
    Err.Clear
'====================================================================================




End Sub
Private Sub Preview()
    On Error GoTo Hell
    Dim RcTes As New DBQuick
    Dim strSQL As String
    If cekListKosong = False Then
        'strSQL = " SELECT * FROM [" & rsReport.Fields("ViewObject").Value & "]" & ScanFilter2
        strSQL = " SELECT * FROM [" & GridReport.Columns(1).Value & "]" & ScanFilter2
        strSQL = strSQL & mVarTmp
'        Debug.Print strSQL
    Else
        strSQL = " SELECT * FROM [" & GridReport.Columns(1).Value & "]"
    
    End If
    StrLaporan = strSQL
Hell:
    ' MessageBox Err.Description, vbCritical, App.ProductName
    Err.Clear
End Sub

Private Function ScanFilter2() As String
Dim sWhere, sData As String
Dim sFrom, sTo As String
Dim I As Integer
For I = 1 To ListFilter.ListItems.Count
    If ListFilter.ListItems(I).Checked = True Then
        
        If ListFilter.ListItems(I).SubItems(5) <> "BETWEEN" Then
            sFrom = FormatType(ListFilter.ListItems(I).SubItems(4), ListFilter.ListItems(I).SubItems(3))
            sData = FormedParam(ListFilter.ListItems(I).SubItems(5), sFrom)
            Select Case ListFilter.ListItems(I).SubItems(5)
                Case "LIKE"
                    sFrom = Mid$(sFrom, 1, Len(sFrom) - 1)
                    sWhere = sWhere & "(" & ListFilter.ListItems(I).Text & " LIKE " & sFrom & "%') AND "
                Case "NOT LIKE"
                    sWhere = sWhere & "NOT (" & ListFilter.ListItems(I).Text & " LIKE " & sFrom & "%') AND "
                Case Else
                    sWhere = sWhere & "(" & ListFilter.ListItems(I).Text & sData
            End Select
        Else
            sFrom = FormatType(ListFilter.ListItems(I).SubItems(4), ListFilter.ListItems(I).SubItems(1))
            sTo = FormatType(ListFilter.ListItems(I).SubItems(4), ListFilter.ListItems(I).SubItems(2))
            Select Case ListFilter.ListItems(I).SubItems(5)
                Case "BETWEEN"
                    sWhere = sWhere & "(" & ListFilter.ListItems(I).Text & " >= " & sFrom & " AND " & _
                             ListFilter.ListItems(I).Text & " <= " & sTo & ") AND "
                Case "NOT BETWEEN"
                    sWhere = sWhere & "(NOT (" & ListFilter.ListItems(I).Text & " >= " & sFrom & " AND " & _
                             ListFilter.ListItems(I).Text & " <= " & sTo & ")) AND "
            End Select
            
            
        End If
    End If
Next

If Trim(sWhere) <> "" Then
   sWhere = Mid$(sWhere, 1, Len(sWhere) - 4)
   sWhere = " Where " & sWhere
End If
ScanFilter2 = sWhere
End Function
Private Function FormatType(sType As Integer, sText As String) As String
Select Case sType
    Case 202, 203, 200  'TEXT
        FormatType = FNumText(sText)
    Case 3, 5, 6    'NUMERIC
        FormatType = FQty(sText)
    Case 135        'DATE
'        sText = CDate(sText)
        FormatType = FDatePicker(DateData, CDate(sText))
End Select
End Function

Private Function FormedParam(ByVal sOperator As String, ByVal sData1 As String, Optional ByVal sData2 As String) As String
Select Case sOperator
    Case "="
        FormedParam = sOperator & " " & sData1 & ") AND "
    Case "<>"
        FormedParam = sOperator & " " & sData1 & ") AND "
    Case "<"
        FormedParam = sOperator & " " & sData1 & ") AND "
    Case "<="
        FormedParam = sOperator & " " & sData1 & ") AND "
    Case ">"
        FormedParam = sOperator & " " & sData1 & ") AND "
    Case ">="
        FormedParam = sOperator & " " & sData1 & ") AND "
'    Case "BETWEEN"
'        FormedParam = sOperator & " " & sData1 & ") AND "
'    Case "NOT BETWEEN"
'        FormedParam = sOperator & " " & sData1 & ") AND "
'    Case "LIKE"
'        FormedParam = sOperator & " " & sData1 & ") AND "
'    Case "NOT LIKE"
'        FormedParam = sOperator & " " & sData1 & ") AND "
End Select
End Function
Function AdaSpasi(texNya As String) As Boolean
    Dim I As Integer

    For I = 1 To Len(texNya)

        If Asc(Mid(texNya, I, 1)) = 32 Then
            AdaSpasi = True
            Exit Function
        End If

    Next I

End Function

Private Function cekListKosong() As Boolean
    Dim ncount As Integer
    cekListKosong = True

    For ncount = 1 To ListFilter.ListItems.Count

        If ListFilter.ListItems(ncount).Checked = True Then cekListKosong = False
    Next

End Function

Private Sub GridReport_DblClick()
    Dim mVarWindState As Integer
    mVarWindState = Me.WindowState
    Preview
    Me.WindowState = mVarWindState
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

    If Rc.DBOpen("SELECT Report_Filter.REPORT_ID, Report_Filter.FIELD_NAME, Report_Filter.FIELD_TYPE, Report_Filter.OBJECT_TYPE From [Report Modules]  INNER JOIN Report_Filter ON ([Report Modules].NoIdx = Report_Filter.REPORT_ID) Where [Report Modules].NoIdx = '" & CInt(lblReportIndex) & "' AND [Report Modules].ViewObject = '" & sViewObject & "'", CNN, lckLockBatch) = True Then

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

Private Function ScanFilter() As String
    Dim mVarI As Integer
    Dim mVarTmp As String
    Dim RcFlt As New Recordset
    ScanFilter = ""

    If Not RcFilter Is Nothing Then
        If RcFilter.Recordcount <> 0 Then
            Set RcFlt = RcFilter.DBRecordset.Clone(adLockReadOnly)

            With RcFlt

                If .Recordcount <> 0 Then
                    .Filter = "Isi <> ''"

                    If .Recordcount <> 0 Then
                        .MoveFirst

                        Do

                            If .EOF = True Then Exit Do

                            Select Case ScanTypeFld(.Fields(0).Value)

                                Case fldString
                                    mVarTmp = Trim(mVarTmp & "[" & .Fields(0) & "] Like N'" & .Fields(1) & "%'") & " AND "

                                Case fldNumeric
                                    mVarTmp = Trim(mVarTmp & "[" & .Fields(0) & "] = " & .Fields(1)) & " AND "

                                Case fldTanggal
                                    mVarTmp = mVarTmp & "[" & .Fields(0) & "] >= Convert(datetime,'" & Format(.Fields(1), "dd/mm/yy") & "',3)" & " AND "
                            End Select

                            .MoveNext
                        Loop

                        .MoveFirst
                        ScanFilter = " WHERE " & Left(mVarTmp, Len(mVarTmp) - 5)
                    End If
                End If

            End With

        End If
    End If

    CloseDB RcFlt
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

Private Sub IsiLabaRugi(ByVal PeriodeData As Integer)
    Dim Rc As New DBQuick
    Dim mVarJrl As New clsJournal
    Rc.DBOpen "SELECT ABS(SUM([Detail Journal].Debet - [Detail Journal].Credit)) AS Debet FROM  [Detail Journal] INNER JOIN GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID WHERE     ([Tabel Pembantu].[Kelompok Perkiraan] = 0) AND ([Table Journal].Periode = " & PeriodeData & ")", CNN, lckLockReadOnly

    With Rc

        If .Recordcount <> 0 Then
            SendDataToServer ("Update [Tabel Pembantu] set CurrentCR" & mVarTempPeriode & " = " & -IIf(Not IsNull(.Fields(0)), .Fields(0), 0) & " Where NoAccount=N'" & AccountLink & "'")
        Else
        End If

    End With

End Sub

Private Function AccountLink() As String
    Dim Rc As New DBQuick
    Rc.DBOpen "SELECT     NoAccount FROM         [Tabel Pembantu] WHERE     ([Seting Relasi] = 1)", CNN, lckLockReadOnly
    AccountLink = "xxx"

    With Rc

        If .Recordcount <> 0 Then
            AccountLink = IIf(Not IsNull(.Fields(0)), .Fields(0), "xxx")
        End If

    End With

End Function

Private Sub GridLayout()
'    Dim nRow As Integer
'
'    For nRow = 0 To GridReport.Columns.Count - 1
'
'        If DgDesign.Columns(nRow).Caption = "REPORT_ID" Then
'            DgDesign.Columns(nRow).width = 0
'        ElseIf DgDesign.Columns(nRow).Caption = "NoIdx" Then
'            DgDesign.Columns(nRow).width = 0
'        End If
'
'        If GridReport.Columns(nRow).Caption = "REPORT_ID" Then
'            GridReport.Columns(nRow).width = 0
'        ElseIf GridReport.Columns(nRow).Caption = "NoIdx" Then
'            GridReport.Columns(nRow).width = 0
'        End If
'
'    Next
With GridReport
    .Height = 3350
    .Columns(0).Visible = False
    .Columns(1).Visible = False
    
    .Columns(3).Visible = False
    .Columns(4).Visible = False
    .Columns(2).width = 4000
    .Columns(5).width = 5125
    .HoldFields
End With
With DgDesign
    .Height = 4300
    .Columns(0).Visible = False
    .Columns(1).Visible = False
    .Columns(2).Visible = False
    .Columns(3).width = 2500
    .Columns(4).width = .Columns(3).width
    .Columns(5).width = 3925
    .Columns(3).Caption = "Judul Laporan"

    .HoldFields
End With
End Sub

Private Sub InsReportPermit()
'In used for insert to report permit for report access
rsREC.DBOpen "select * from [report modules] order by noIDX", CNN, lckLockBatch
rsREC.DBRecordset.MoveLast
SendDataToServer "insert into [report permit] ([User ID], noidx, laporan) values (" & aksess.GetID & ", '" & rsREC.Fields("noidx") & " ',0)"
End Sub

Private Sub DelReportPermit()
'in used for remove to report permit for remove report access
SendDataToServer "delete from [report permit] where [user id]=" & aksess.GetID & " and noidx=" & DgDesign.Columns(0).Value & ""
End Sub

Private Sub DelreportModules()
'In Used for remove to report modules
SendDataToServer "delete from [report modules] where noIdx = '" & Label3.Caption & "'"
End Sub

