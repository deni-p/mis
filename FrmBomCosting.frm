VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmBOMCosting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Konfigurasi Harga Pokok"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBomCosting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Tag             =   "Product Costing"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   8325
      Left            =   0
      ScaleHeight     =   8325
      ScaleWidth      =   10305
      TabIndex        =   8
      Top             =   0
      Width           =   10305
      Begin TabDlg.SSTab SSTab1 
         Height          =   8025
         Left            =   150
         TabIndex        =   1
         Top             =   150
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   14155
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
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
         TabCaption(0)   =   "List Barang"
         TabPicture(0)   =   "FrmBomCosting.frx":6852
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Picture3"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail Biaya"
         TabPicture(1)   =   "FrmBomCosting.frx":686E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Picture4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture3 
            Height          =   7545
            Left            =   -74925
            ScaleHeight     =   7485
            ScaleWidth      =   9735
            TabIndex        =   14
            Top             =   375
            Width           =   9800
            Begin MSComctlLib.ListView ListView1 
               Height          =   7485
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   9780
               _ExtentX        =   17251
               _ExtentY        =   13203
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15380335
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
                  Text            =   "Kode Barang"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Nama Barang"
                  Object.Width           =   5644
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "UOM"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Fixed Cost"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Last Cost"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Average Cost"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00EAAF6F&
            Height          =   7575
            Left            =   75
            ScaleHeight     =   7515
            ScaleWidth      =   9735
            TabIndex        =   9
            Top             =   360
            Width           =   9800
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Kode Barang"
               Height          =   315
               Index           =   0
               Left            =   1410
               MaxLength       =   16
               TabIndex        =   2
               Tag             =   "Partner"
               Text            =   " - Kode Barang -"
               Top             =   75
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "UOM"
               Height          =   315
               Index           =   1
               Left            =   6645
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Tag             =   "Partner"
               Text            =   " - Satuan - "
               Top             =   405
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Nama Barang"
               Height          =   315
               Index           =   2
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   3
               Tag             =   "Partner"
               Text            =   " - Nama Barang  -"
               Top             =   405
               Width           =   3945
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Fixed Cost"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#.##0;(#.##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   1410
               Locked          =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Tag             =   "Partner"
               Text            =   " - Fixed Cost -"
               Top             =   735
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Average Cost"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   6645
               Locked          =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Tag             =   "Partner"
               Text            =   " - Average Cost - "
               Top             =   735
               Width           =   2250
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Last Cost"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#.##0;(#.##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   1410
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Tag             =   "Partner"
               Text            =   " - Last Cost -"
               Top             =   1065
               Width           =   2250
            End
            Begin TabDlg.SSTab TabDetail 
               Height          =   5850
               Left            =   225
               TabIndex        =   18
               Top             =   1545
               Width           =   9330
               _ExtentX        =   16457
               _ExtentY        =   10319
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Group Biaya"
               TabPicture(0)   =   "FrmBomCosting.frx":688A
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "LBLDetil"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Picture1"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Picture5"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).ControlCount=   3
               TabCaption(1)   =   "Detail Akun"
               TabPicture(1)   =   "FrmBomCosting.frx":68A6
               Tab(1).ControlEnabled=   0   'False
               Tab(1).ControlCount=   0
               Begin VB.PictureBox Picture5 
                  Height          =   2430
                  Left            =   60
                  ScaleHeight     =   2370
                  ScaleWidth      =   9120
                  TabIndex        =   23
                  Top             =   3330
                  Width           =   9180
                  Begin VB.CommandButton CmdFresh 
                     Height          =   330
                     Index           =   1
                     Left            =   405
                     Picture         =   "FrmBomCosting.frx":68C2
                     Style           =   1  'Graphical
                     TabIndex        =   33
                     Tag             =   "True"
                     ToolTipText     =   "Go"
                     Top             =   2025
                     Width           =   345
                  End
                  Begin VB.TextBox TxtCarik 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     Height          =   315
                     Left            =   750
                     TabIndex        =   32
                     Tag             =   "True"
                     Top             =   2040
                     Width           =   3390
                  End
                  Begin VB.CommandButton CmdFresh 
                     Height          =   330
                     Index           =   0
                     Left            =   60
                     Picture         =   "FrmBomCosting.frx":D114
                     Style           =   1  'Graphical
                     TabIndex        =   31
                     Tag             =   "True"
                     ToolTipText     =   "Refresh"
                     Top             =   2025
                     Width           =   345
                  End
                  Begin VB.CommandButton CmdPanahHPP 
                     Enabled         =   0   'False
                     Height          =   360
                     Index           =   1
                     Left            =   4305
                     Picture         =   "FrmBomCosting.frx":13966
                     Style           =   1  'Graphical
                     TabIndex        =   30
                     Top             =   855
                     Width           =   495
                  End
                  Begin VB.CommandButton CmdPanahHPP 
                     Enabled         =   0   'False
                     Height          =   360
                     Index           =   2
                     Left            =   4305
                     Picture         =   "FrmBomCosting.frx":13A5E
                     Style           =   1  'Graphical
                     TabIndex        =   29
                     Top             =   1215
                     Width           =   495
                  End
                  Begin VB.CommandButton CmdPanahHPP 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     Height          =   360
                     Index           =   0
                     Left            =   4305
                     Picture         =   "FrmBomCosting.frx":13B50
                     Style           =   1  'Graphical
                     TabIndex        =   28
                     Top             =   495
                     Width           =   495
                  End
                  Begin VB.CommandButton CmdPanahHPP 
                     Enabled         =   0   'False
                     Height          =   360
                     Index           =   3
                     Left            =   4305
                     Picture         =   "FrmBomCosting.frx":13C42
                     Style           =   1  'Graphical
                     TabIndex        =   27
                     Top             =   1575
                     Width           =   495
                  End
                  Begin MSDataGridLib.DataGrid GridHPP 
                     Height          =   2370
                     Index           =   1
                     Left            =   4905
                     TabIndex        =   24
                     Top             =   0
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   4180
                     _Version        =   393216
                     AllowUpdate     =   -1  'True
                     BorderStyle     =   0
                     HeadLines       =   1
                     RowHeight       =   15
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
                     ColumnCount     =   2
                     BeginProperty Column00 
                        DataField       =   "NoAccount"
                        Caption         =   "Kode"
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
                        DataField       =   "AccountName"
                        Caption         =   "Nama Rekening"
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
                        ScrollBars      =   2
                        BeginProperty Column00 
                        EndProperty
                        BeginProperty Column01 
                        EndProperty
                     EndProperty
                  End
                  Begin MSDataGridLib.DataGrid GridHPP 
                     Height          =   1935
                     Index           =   0
                     Left            =   -30
                     TabIndex        =   25
                     Top             =   0
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   3413
                     _Version        =   393216
                     AllowUpdate     =   -1  'True
                     BorderStyle     =   0
                     HeadLines       =   1
                     RowHeight       =   15
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
                     ColumnCount     =   2
                     BeginProperty Column00 
                        DataField       =   "NoAccount"
                        Caption         =   "Kode"
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
                        DataField       =   "AccountName"
                        Caption         =   "Nama Rekening"
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
                        ScrollBars      =   2
                        BeginProperty Column00 
                        EndProperty
                        BeginProperty Column01 
                        EndProperty
                     EndProperty
                  End
               End
               Begin VB.PictureBox Picture1 
                  Height          =   2430
                  Left            =   60
                  ScaleHeight     =   2370
                  ScaleWidth      =   9120
                  TabIndex        =   19
                  Top             =   375
                  Width           =   9180
                  Begin MSDataGridLib.DataGrid DataGrid1 
                     Height          =   2400
                     Index           =   0
                     Left            =   -15
                     TabIndex        =   20
                     Top             =   -15
                     Width           =   9150
                     _ExtentX        =   16140
                     _ExtentY        =   4233
                     _Version        =   393216
                     AllowUpdate     =   0   'False
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
                     ColumnCount     =   3
                     BeginProperty Column00 
                        DataField       =   "Cost Element"
                        Caption         =   "Jenis Biaya"
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
                        DataField       =   "Keterangan"
                        Caption         =   "Keterangan"
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
                        DataField       =   "Cost"
                        Caption         =   "Cost"
                        BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                           Type            =   1
                           Format          =   "#,##0.00;(#,##0.00)"
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
                           Alignment       =   1
                        EndProperty
                     EndProperty
                  End
               End
               Begin VB.Label LBLDetil 
                  BackColor       =   &H00000000&
                  Caption         =   " DETIL REKENING "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   90
                  TabIndex        =   26
                  Top             =   2955
                  Width           =   9120
               End
            End
            Begin VB.Label LblAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " - Cost Rollup -"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00;(#,##0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   6645
               TabIndex        =   21
               Top             =   1065
               Width           =   2250
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   5475
               X2              =   7020
               Y1              =   1365
               Y2              =   1365
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cost Rollup"
               Height          =   195
               Index           =   3
               Left            =   5490
               TabIndex        =   22
               Top             =   1125
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Average Cost"
               Height          =   195
               Index           =   6
               Left            =   5490
               TabIndex        =   17
               Top             =   795
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               Height          =   195
               Index           =   5
               Left            =   5490
               TabIndex        =   16
               Top             =   465
               Width           =   345
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   285
               X2              =   1710
               Y1              =   1035
               Y2              =   1035
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   285
               X2              =   1710
               Y1              =   1365
               Y2              =   1365
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last Cost"
               Height          =   195
               Index           =   2
               Left            =   285
               TabIndex        =   13
               Top             =   1110
               Width           =   675
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   5475
               X2              =   6900
               Y1              =   1035
               Y2              =   1035
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   5475
               X2              =   6900
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fixed Cost"
               Height          =   195
               Index           =   4
               Left            =   285
               TabIndex        =   12
               Top             =   780
               Width           =   765
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   285
               X2              =   1710
               Y1              =   375
               Y2              =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Barang"
               Height          =   195
               Index           =   1
               Left            =   285
               TabIndex        =   11
               Top             =   120
               Width           =   915
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   285
               X2              =   1710
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Barang"
               Height          =   195
               Index           =   0
               Left            =   285
               TabIndex        =   10
               Top             =   450
               Width           =   960
            End
         End
      End
   End
End
Attribute VB_Name = "frmBOMCosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAdd As Boolean
Private RcPart As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mFirstCaller, mKeyLoad As Boolean
Private WithEvents rsCost As DBQuick
Attribute rsCost.VB_VarHelpID = -1
Private WithEvents rsCostAcc As DBQuick
Attribute rsCostAcc.VB_VarHelpID = -1
Private rsGL As New DBQuick
Dim strSQL As String

Private Sub CmdFresh_Click(Index As Integer)
   Select Case Index
      Case 0:
         TxtCarik.Text = ""
         rsGL.DBRecordset.Filter = adFilterNone
         rsGL.DBRecordset.Requery
         Set GridHPP(0).DataSource = rsGL.DBRecordset
         
      Case 1: If TxtCarik <> "" Then rsGL.DBRecordset.Filter = "AccountName like '%" & TxtCarik & "%'"
   End Select
End Sub

Private Sub CmdPanahHPP_Click(Index As Integer)
   Select Case Index
      Case 0
         SendToCostAcc rsGL.DBRecordset.Fields("NoAccount"), rsGL.DBRecordset.Fields("AccountName")
      Case 1
         With rsGL.DBRecordset
            .MoveFirst
            While Not .EOF
               SendToCostAcc .Fields("NoAccount"), .Fields("AccountName")
               .MoveNext
            Wend
         End With
      Case 2
         rsCostAcc.DBRecordset.Delete
      Case 3
         With rsCostAcc.DBRecordset
            .MoveFirst
            While Not .EOF
               .Delete
               .MoveNext
            Wend
         End With
   End Select
End Sub

Private Sub SendToCostAcc(noAccount As String, AccountName As String)
   With rsCostAcc.DBRecordset
      If .Recordcount = 1 Then
         If noAccount <> .Fields("NoAccount") Then
            AddCostAcc noAccount, AccountName
         End If
      ElseIf .Recordcount > 1 Then
         .MoveFirst
         .Find "noAccount ='" & noAccount & "'", 1, adSearchForward, 0
         If Not .EOF Then
            If (noAccount <> .Fields("NoAccount")) Then
               AddCostAcc noAccount, AccountName
            End If
         Else
            AddCostAcc noAccount, AccountName
         End If
      Else
         AddCostAcc noAccount, AccountName
      End If
   End With
End Sub

Private Sub AddCostAcc(noAccount As String, AccountName As String)
   With rsCostAcc.DBRecordset
      .AddNew
      .Fields("NoAccount") = noAccount
      .Fields("AccountName") = AccountName
      .Fields("ID") = rsCost.DBRecordset.Fields("ID")
      .Filter = "id='" & rsCost.DBRecordset.Fields("ID") & "'"
   End With
End Sub



Private Sub DataGrid1_AfterColEdit(Index As Integer, ByVal ColIndex As Integer)
If DataGrid1(0).col = 2 And mAdd = True Then TotalTrans
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
'Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If mAdd = True Then
    DataGrid1(0).MarqueeStyle = dbgFloatingEditor
    Select Case DataGrid1(0).col
        Case 2:
           DataGrid1(Index).AllowUpdate = mAdd
        Case Else
           'DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
           DataGrid1(Index).AllowUpdate = False
    End Select
Else
   DataGrid1(0).AllowUpdate = mAdd
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mKeyLoad = False Then mKeyLoad = True Else mKeyLoad = False
If mKeyLoad = False Then ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
SSTab1.Tab = 0
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmBOMCosting
    .SetPermissions = UserAddnewDenied
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], UOM, FixCost AS [Fixed Cost], AvgCost AS [Average Cost], LastCost AS [Last Cost] FROM         Inventory WHERE     (Manufacture = 1)"
End With
OpenHeader
OpenGL
Set mCall = New frmCaller
DataGrid1(0).AllowUpdate = True
LBLDetil.BackColor = &H0&
LBLDetil.ForeColor = &HFFFFFF
LBLDetil.FontBold = True
End Sub
Private Sub OpenGL()
strSQL = "SELECT NoAccount, AccountName From GLAccount WHERE ([Group] = N'Detail List Account') ORDER BY NoAccount"
rsGL.DBOpen strSQL, CNN, lckLockReadOnly, lckLockSync

Set GridHPP(0).DataSource = rsGL.DBRecordset
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set frmBOMCosting = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set frmBOMCosting = Nothing
End If
End Sub

Private Sub Form_Resize()
GridLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mCall = Nothing
Set frmBOMCosting = Nothing
End Sub

Private Sub GridHPP_HeadClick(Index As Integer, ByVal ColIndex As Integer)
Select Case Index
    Case 0
        rsGL.DBRecordset.Sort = GridHPP(0).Columns(ColIndex).DataField
End Select
End Sub

Private Sub ListView1_DblClick()
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   SSTab1.Tab = 1
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   MyDDE.FindStringData "[Kode Barang]='" & Item.Text & "'"
End If
End Sub

Private Sub mCall_BeforeUnload()
On Error GoTo 1
If FindOwnRecordset(MyDDE.ChildRecordset, "[Cost Element] = '" & MyDDE.ChildRecordset.Fields("Cost Element") & "'") = True Then
   MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Cost Element") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
   MyDDE.ChildRecordset.CancelBatch adAffectCurrent
   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
Else
   If Not IsNull(MyDDE.ChildRecordset.Fields("Cost Element")) = True Then
      If MyDDE.ChildRecordset.Fields("Cost Element") = "" Then
         MyDDE.ChildRecordset.CancelBatch adAffectCurrent
         If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
      Else
         Dim rsID As New DBQuick
         rsID.DBOpen "select newID()", CNN
         MyDDE.ChildRecordset.Fields("ID") = rsID.DBRecordset.Fields(0)
         rsID.CloseDB
      End If
   End If
End If
mAdd = txtBox(3).Enabled
Exit Sub
1:
MessageBox Err.Description, "frmbomcosting:mcall_beforeunload" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 2
Select Case mCall.FromTagActive
       Case "COST ELEMENT":
            With MyDDE.ChildRecordset
                 .Fields("Cost Element") = mCall.GetFieldByName(0)
                 .Fields("Keterangan") = mCall.GetFieldByName(1)
                 .Fields("Cost") = 0
            End With
End Select
Exit Sub
2:
MessageBox Err.Description, "frmbomcosting:mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo xErr
Dim x As Integer
For x = 0 To 3
   CmdPanahHPP(x).Enabled = False
Next
Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
            txtBox(0).SetFocus
            'Label2 = IndexAuto

            SSTab1.Tab = 1
            For x = 0 To 3
               CmdPanahHPP(x).Enabled = True
            Next
       Case tmbEdit:
            txtBox(0).Enabled = False
            txtBox(1).Enabled = False
            txtBox(2).Enabled = False
            mAdd = True
            txtBox(3).SetFocus
            SSTab1.Tab = 1
            TotalTrans
            For x = 0 To 3
               CmdPanahHPP(x).Enabled = True
            Next
            
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
            Select Case MyDDE.ChildRecordset.status
                   Case 8: mAdd = False
                   Case Else
                        If MyDDE.ChildRecordset.Recordcount <> 0 Then
                           mAdd = True
                        Else
                           mAdd = False
                        End If
            End Select
            Else
               mAdd = False
            End If
            
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then OpenDetailPartner SSTab1.Tab
            
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  With MyDDE.ChildRecordset
                       .MoveFirst
                       If SendDataToServer("Delete From [BOM Costing Detail] WHERE  (NoItem = N'" & txtBox(0) & "')") = True Then
                          Do
                            If MyDDE.ChildRecordset.EOF Then Exit Do
                            SendDataToServer "delete from bom_cogs_detail where id='" & .Fields("ID") & "'"
                            SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                                             " ( ID,NoItem, [Cost Element Type], CostValue)" & _
                                             " VALUES ('" & .Fields("ID") & "',N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & ")"
                            .MoveNext
                          Loop
                       End If
                       .MoveLast
                  End With
                  With rsCostAcc.DBRecordset
                     .Filter = adFilterNone
                     If .Recordcount > 0 Then
                        .MoveFirst
                        While Not .EOF
                           SendDataToServer "insert into bom_cogs_detail (id,no_account) values ('" & .Fields("ID") & "','" & .Fields("noAccount") & "')"
                           .MoveNext
                        Wend
                     End If
                  End With
               End If
               mAdd = False
            End If
            txtBox(4).Text = Format((CDbl(txtBox(3).Text) + CDbl(txtBox(5).Text)) / 2, QtyFormFloat)
       Case tmbPrint:
            CallRPTReport "BOM Costing List.rpt", "sELECT * FROM [BOM Costing List] Where [BOM ID] =N'" & txtBox(0) & "'"
       Case Else: 'mVarDataDc = False
End Select
mAdd = txtBox(3).Enabled
SSTab1.TabEnabled(0) = Not mAdd
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   OpenDetail MyDDE.GetFieldByName("Kode Barang")
   TotalTrans
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Dim mDel As New clsDelete
Dim CascaDel As String
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterBarang) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  CascaDel = mDel.CekCascadeDeleteTable(txtBox(0), reDelMasterBarang)
                  MessageBox "Data ' " & txtBox(2).Text & " (" & txtBox(0) & ") '  sedang digunakan pada transaksi " & _
                  UCase(CascaDel) & vbCrLf & "Record tidak bisa dihapus..", "Kontrol Penghapusan Data", msgOkOnly, msgExclamation
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
        Else
            MessageBox "Data Barang tidak dapat dihapus", "Inventory Kontrol", msgOkOnly, msgCrtical
            MyDDE.CancelTrans = True
            MyDDE.IsChildMemberReady = False
        End If
        Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  If mAdd = True Then txtBox(3) = CDbl(LblAmount)
                  PrepareQuery
               Else
                  'MessageBox "Date detail calendar belum ada.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
                  PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDetail:
            MyDDE.CancelTrans = mFirstCaller
            If MyDDE.CancelTrans = True Then Exit Sub
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
               If MyDDE.ChildRecordset.Fields(2) = 0 Then
                  MyDDE.IsChildMemberReady = False
                  MyDDE.CancelTrans = True
                  MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly, msgCrtical
               Else
                  MyDDE.IsChildMemberReady = True
                  MyDDE.CancelTrans = False
               End If
            Else
               MyDDE.IsChildMemberReady = True
               MyDDE.CancelTrans = False
            End If
End Select
Set mDel = Nothing
Exit Sub
1:
MessageBox Err.Description, "frmbomcosting:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
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
'MessageBox .PrepareAppend
    .PrepareUpdate = " UPDATE [Inventory] Set FixCost=" & CDbl(txtBox(3)) & ",AvgCost=" & CDbl(txtBox(4)) & ",LastCost=" & CDbl(txtBox(5)) & " WHERE     ([NoItem] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Inventory] WHERE   ([NoItem] = N'" & txtBox(0) & "') "
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub GridLayout()
'DataGrid1(0).Height = 2160
'DataGrid1(0).width = 8910
DataGrid1(0).Columns(0).width = 2340.284
DataGrid1(0).Columns(1).width = 4169.764
DataGrid1(0).Columns(2).width = 1830.047
End Sub
Private Sub rsCost_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo xErr
If pRecordset.Recordcount <> 0 Then
    LBLDetil.Caption = "DETIL REKENING ITEM HPP - " & UCase(pRecordset.Fields(1).Value)
    rsCostAcc.DBRecordset.Filter = "id='" & rsCost.DBRecordset.Fields("ID") & "'"
End If
Exit Sub
xErr:
   Err.Clear
End Sub
Private Sub OpenDetail(ByVal Param As String)
Set rsCost = New DBQuick
Set rsCostAcc = New DBQuick

rsCost.DBOpen "SELECT [BOM Costing Detail].[Cost Element Type] AS [Cost Element], " & _
        " [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost,[BOM Costing Detail].ID " & _
        " FROM [BOM Costing Detail] INNER JOIN [Cost Element] ON " & _
        " [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] " & _
        " WHERE ([BOM Costing Detail].NoItem = N'" & Param & "') ORDER BY " & _
        " [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
Set MyDDE.ChildRecordset = rsCost.DBRecordset.Clone(adLockBatchOptimistic)
'Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
Set DataGrid1(0).DataSource = rsCost.DBRecordset


strSQL = "SELECT bom_cogs_detail.id, GLAccount.noAccount, GLAccount.AccountName " & _
         "FROM [BOM Costing Detail] INNER JOIN " & _
               "bom_cogs_detail ON [BOM Costing Detail].ID = bom_cogs_detail.id INNER JOIN " & _
               "GLAccount ON bom_cogs_detail.no_account = GLAccount.NoAccount " & _
         "WHERE ([BOM Costing Detail].NoItem = '" & Param & "')"
    
rsCostAcc.DBOpen strSQL, CNN, lckLockBatch
Set GridHPP(1).DataSource = rsCostAcc.DBRecordset

If Not (rsCostAcc.DBRecordset.EOF Or rsCostAcc.DBRecordset.BOF) Then
    rsCostAcc.DBRecordset.Filter = "id='" & rsCost.DBRecordset.Fields("ID") & "'"
End If
End Sub

Private Sub OpenDetailPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1: RcPartner.DBOpen "SELECT     [Cost Element Type] AS [Cost Element], Description AS Keterangan FROM         [Cost Element] ORDER BY [Cost Element Type]", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1: mCall.FromTagActive = "COST ELEMENT"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
'    messagebox Err.Description
    Err.Clear
End Sub

Private Sub OpenHeader()
On Error GoTo 1
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Set Rc.DBRecordset = MyDDE.ActiveRecordset.Clone(adLockReadOnly)
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            With ListView1.ListItems.Add(, , Avdata(0, I))
                 .SubItems(1) = Avdata(1, I)
                 .SubItems(2) = Avdata(2, I)
                 .SubItems(3) = Format(Avdata(3, I), QtyFormFloat)
                 .SubItems(4) = Format(Avdata(5, I), QtyFormFloat)
                 .SubItems(5) = Format(Avdata(4, I), QtyFormFloat)
            End With
        Next I
     Else
     End If
End With
Exit Sub
1:
MessageBox Err.Description, "frmbomcosting:openheader" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Sub TotalTrans()
On Error GoTo 2
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Set Rc.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
LblAmount = 0
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            LblAmount = LblAmount + IIf(Not IsNull(Avdata(2, I)), Avdata(2, I), 0)
        Next I
        LblAmount = FormatNumber(LblAmount, 0)
     End If
End With
Set Avdata = Nothing
Exit Sub
2:
MessageBox Err.Description, "frmbomcosting:totaltrans" & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub txtBox_Validate(Index As Integer, Cancel As Boolean)
txtBox(4).Text = Format((CDbl(txtBox(3).Text) + CDbl(txtBox(5).Text)) / 2, QtyFormFloat)
End Sub

Private Sub TxtCarik_Change()
On Error Resume Next
Dim strcari As String
If Len(TxtCarik.Text) <> 0 Then
    strcari = "[" & GridHPP(0).Columns(1).DataField & "]" & " Like '" & TxtCarik.Text & "%'"
    rsGL.DBRecordset.Filter = strcari  ', 0, adSearchForward, adBookmarkFirst
    If rsGL.DBRecordset.Recordcount = 0 Then MessageBox "Kriteria Yang Dicari Tidak Ada..............!", vbCritical
Else
    CmdFresh(0).Value = True
End If
End Sub
