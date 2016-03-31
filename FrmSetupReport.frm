VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSetupReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7185
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetupReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11610
   Tag             =   "Configuration Account"
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000018&
      Height          =   570
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   11550
      TabIndex        =   27
      Top             =   6615
      Width           =   11610
      Begin VB.CommandButton CmdKeluar 
         Caption         =   "Keluar"
         Height          =   405
         Left            =   9990
         TabIndex        =   28
         Top             =   45
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6585
      Left            =   60
      ScaleHeight     =   6555
      ScaleWidth      =   11430
      TabIndex        =   7
      Top             =   165
      Width           =   11460
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5700
         Left            =   75
         ScaleHeight     =   5670
         ScaleWidth      =   11220
         TabIndex        =   8
         Top             =   645
         Width           =   11250
         Begin TabDlg.SSTab SSTab1 
            Height          =   5580
            Left            =   45
            TabIndex        =   0
            Top             =   45
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   9843
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            Tab             =   1
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
            TabCaption(0)   =   "Account Setting"
            TabPicture(0)   =   "FrmSetupReport.frx":6852
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "CmdGrp(1)"
            Tab(0).Control(1)=   "CmdGrp(0)"
            Tab(0).Control(2)=   "Frame1"
            Tab(0).Control(3)=   "DgSeting"
            Tab(0).Control(4)=   "Label2(1)"
            Tab(0).Control(5)=   "Label2(0)"
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "CashFlow Configuration"
            TabPicture(1)   =   "FrmSetupReport.frx":686E
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "SSTab2"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "CmdKonfig(0)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "CmdKonfig(1)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "DgAccount"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Intelligent Account"
            TabPicture(2)   =   "FrmSetupReport.frx":688A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "CmdControl(1)"
            Tab(2).Control(1)=   "CmdControl(0)"
            Tab(2).Control(2)=   "DgControl"
            Tab(2).ControlCount=   3
            TabCaption(3)   =   "Account Position"
            TabPicture(3)   =   "FrmSetupReport.frx":68A6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Frame3"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Account Level"
            TabPicture(4)   =   "FrmSetupReport.frx":68C2
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "Frame4"
            Tab(4).ControlCount=   1
            Begin MSDataGridLib.DataGrid DgAccount 
               Height          =   5055
               Left            =   90
               TabIndex        =   2
               Top             =   450
               Width           =   5325
               _ExtentX        =   9393
               _ExtentY        =   8916
               _Version        =   393216
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               HeadLines       =   2
               RowHeight       =   15
               RowDividerStyle =   3
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
                  DataField       =   "No Akun"
                  Caption         =   "Kode Account"
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
                  DataField       =   "Nama Akun"
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  BeginProperty Column00 
                     DividerStyle    =   3
                     ColumnWidth     =   1679.811
                  EndProperty
                  BeginProperty Column01 
                     DividerStyle    =   3
                     ColumnWidth     =   3075.024
                  EndProperty
               EndProperty
            End
            Begin VB.Frame Frame4 
               Height          =   5070
               Left            =   -74895
               TabIndex        =   37
               Top             =   360
               Width           =   10920
               Begin VB.TextBox txtAccountLen 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   0
                  Left            =   5565
                  Locked          =   -1  'True
                  TabIndex        =   42
                  Top             =   4635
                  Width           =   2100
               End
               Begin VB.PictureBox Picture4 
                  Height          =   4320
                  Left            =   75
                  ScaleHeight     =   4260
                  ScaleWidth      =   10695
                  TabIndex        =   40
                  Top             =   180
                  Width           =   10755
                  Begin MSDataGridLib.DataGrid DgPrefix 
                     Bindings        =   "FrmSetupReport.frx":68DE
                     Height          =   4230
                     Left            =   15
                     TabIndex        =   41
                     Top             =   15
                     Width           =   10650
                     _ExtentX        =   18785
                     _ExtentY        =   7461
                     _Version        =   393216
                     AllowUpdate     =   0   'False
                     BorderStyle     =   0
                     HeadLines       =   1
                     RowHeight       =   15
                     FormatLocked    =   -1  'True
                     BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
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
                     ColumnCount     =   4
                     BeginProperty Column00 
                        DataField       =   "No"
                        Caption         =   "No"
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
                        DataField       =   "Group Name"
                        Caption         =   "Level"
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
                        DataField       =   "Length Level"
                        Caption         =   "Length Level"
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
                        DataField       =   "Prefix"
                        Caption         =   "Prefix"
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
                           Alignment       =   2
                        EndProperty
                        BeginProperty Column01 
                        EndProperty
                        BeginProperty Column02 
                           Alignment       =   1
                        EndProperty
                        BeginProperty Column03 
                        EndProperty
                     EndProperty
                  End
               End
               Begin VB.CommandButton CmdLevel 
                  Caption         =   "Lock Grid"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   1
                  Left            =   1725
                  TabIndex        =   39
                  Top             =   4605
                  Width           =   1485
               End
               Begin VB.CommandButton CmdLevel 
                  Caption         =   "Edit"
                  Height          =   390
                  Index           =   0
                  Left            =   90
                  TabIndex        =   38
                  Top             =   4605
                  Width           =   1485
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Chart Of Account Result"
                  Height          =   210
                  Index           =   2
                  Left            =   3405
                  TabIndex        =   43
                  Top             =   4680
                  Width           =   1995
               End
               Begin VB.Line Line1 
                  Index           =   4
                  X1              =   3405
                  X2              =   7125
                  Y1              =   4935
                  Y2              =   4935
               End
            End
            Begin VB.Frame Frame3 
               Height          =   5070
               Left            =   -74910
               TabIndex        =   31
               Top             =   360
               Width           =   10920
               Begin VB.CommandButton CmdPos 
                  Caption         =   "Edit"
                  Height          =   390
                  Index           =   0
                  Left            =   105
                  TabIndex        =   36
                  Top             =   4605
                  Width           =   1485
               End
               Begin VB.CommandButton CmdPos 
                  Caption         =   "Lock Grid"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   1
                  Left            =   1650
                  TabIndex        =   35
                  Top             =   4605
                  Width           =   1485
               End
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   330
                  Left            =   1695
                  TabIndex        =   32
                  Top             =   210
                  Width           =   4140
                  _ExtentX        =   7303
                  _ExtentY        =   582
                  _Version        =   393216
                  Style           =   2
                  ListField       =   "AccountName"
                  BoundColumn     =   "NoAccount"
                  Text            =   ""
               End
               Begin MSDataGridLib.DataGrid DgPosition 
                  Bindings        =   "FrmSetupReport.frx":68F3
                  Height          =   3915
                  Left            =   105
                  TabIndex        =   34
                  Top             =   630
                  Width           =   10680
                  _ExtentX        =   18838
                  _ExtentY        =   6906
                  _Version        =   393216
                  AllowUpdate     =   0   'False
                  BorderStyle     =   0
                  HeadLines       =   1
                  RowHeight       =   15
                  FormatLocked    =   -1  'True
                  BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
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
                  ColumnCount     =   3
                  BeginProperty Column00 
                     DataField       =   "Code"
                     Caption         =   "Code"
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
                     DataField       =   "Account Name"
                     Caption         =   "Account Name"
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
                     DataField       =   "Position"
                     Caption         =   "Position"
                     BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                        Type            =   5
                        Format          =   ""
                        HaveTrueFalseNull=   1
                        TrueValue       =   "Debet"
                        FalseValue      =   "Credit"
                        NullValue       =   "Credit"
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   1057
                        SubFormatType   =   7
                     EndProperty
                  EndProperty
                  SplitCount      =   1
                  BeginProperty Split0 
                     BeginProperty Column00 
                        Alignment       =   2
                     EndProperty
                     BeginProperty Column01 
                     EndProperty
                     BeginProperty Column02 
                        Alignment       =   2
                     EndProperty
                  EndProperty
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Group Account"
                  Height          =   210
                  Index           =   1
                  Left            =   180
                  TabIndex        =   33
                  Top             =   270
                  Width           =   1245
               End
            End
            Begin VB.CommandButton CmdControl 
               Caption         =   "Lock Grid"
               Enabled         =   0   'False
               Height          =   390
               Index           =   1
               Left            =   -73350
               TabIndex        =   26
               Top             =   5085
               Width           =   1485
            End
            Begin VB.CommandButton CmdControl 
               Caption         =   "Edit"
               Height          =   390
               Index           =   0
               Left            =   -74895
               TabIndex        =   25
               Top             =   5085
               Width           =   1485
            End
            Begin VB.CommandButton CmdKonfig 
               Caption         =   "Simpan"
               Enabled         =   0   'False
               Height          =   390
               Index           =   1
               Left            =   7095
               TabIndex        =   23
               Top             =   5100
               Width           =   1485
            End
            Begin VB.CommandButton CmdKonfig 
               Caption         =   "Mulai"
               Height          =   390
               Index           =   0
               Left            =   5520
               TabIndex        =   22
               Top             =   5100
               Width           =   1485
            End
            Begin VB.CommandButton CmdGrp 
               Caption         =   "Simpan"
               Enabled         =   0   'False
               Height          =   435
               Index           =   1
               Left            =   -72780
               TabIndex        =   19
               Top             =   2640
               Width           =   2085
            End
            Begin VB.CommandButton CmdGrp 
               Caption         =   "Seting Group"
               Height          =   435
               Index           =   0
               Left            =   -74895
               TabIndex        =   18
               Top             =   2640
               Width           =   2085
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   4560
               Left            =   5520
               TabIndex        =   3
               Top             =   465
               Width           =   5490
               _ExtentX        =   9684
               _ExtentY        =   8043
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Arus Kas"
               TabPicture(0)   =   "FrmSetupReport.frx":6908
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Frame5"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Perubahan Modal"
               TabPicture(1)   =   "FrmSetupReport.frx":6924
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Frame2"
               Tab(1).ControlCount=   1
               Begin VB.Frame Frame2 
                  Height          =   4170
                  Left            =   -74925
                  TabIndex        =   16
                  Top             =   315
                  Width           =   5340
                  Begin MSDataGridLib.DataGrid GrdArusKas 
                     Height          =   3615
                     Index           =   1
                     Left            =   75
                     TabIndex        =   6
                     Top             =   465
                     Width           =   5085
                     _ExtentX        =   8969
                     _ExtentY        =   6376
                     _Version        =   393216
                     AllowUpdate     =   0   'False
                     BorderStyle     =   0
                     HeadLines       =   2
                     RowHeight       =   15
                     RowDividerStyle =   3
                     FormatLocked    =   -1  'True
                     BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
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
                     ColumnCount     =   3
                     BeginProperty Column00 
                        DataField       =   "No Akun"
                        Caption         =   "Kode Account"
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
                        DataField       =   "Nama Akun"
                        Caption         =   "Deskripsi"
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
                        DataField       =   "TemplateGroup"
                        Caption         =   "TemplateGroup"
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
                           DividerStyle    =   3
                           ColumnWidth     =   1604.976
                        EndProperty
                        BeginProperty Column01 
                           DividerStyle    =   3
                           ColumnWidth     =   3000.189
                        EndProperty
                        BeginProperty Column02 
                           Object.Visible         =   0   'False
                        EndProperty
                     EndProperty
                  End
                  Begin VB.Label Label2 
                     Alignment       =   2  'Center
                     BackColor       =   &H8000000C&
                     Caption         =   "T E M P L A T E   P E R U B A H A N   M O D A L"
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
                     Index           =   2
                     Left            =   60
                     TabIndex        =   17
                     Top             =   165
                     Width           =   5085
                  End
               End
               Begin VB.Frame Frame5 
                  Height          =   4170
                  Left            =   75
                  TabIndex        =   13
                  Top             =   315
                  Width           =   5280
                  Begin MSDataGridLib.DataGrid GrdArusKas 
                     Height          =   3150
                     Index           =   0
                     Left            =   90
                     TabIndex        =   5
                     Top             =   945
                     Width           =   5085
                     _ExtentX        =   8969
                     _ExtentY        =   5556
                     _Version        =   393216
                     AllowUpdate     =   0   'False
                     BorderStyle     =   0
                     HeadLines       =   2
                     RowHeight       =   15
                     RowDividerStyle =   3
                     FormatLocked    =   -1  'True
                     BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
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
                     ColumnCount     =   3
                     BeginProperty Column00 
                        DataField       =   "No Akun"
                        Caption         =   "Kode Account"
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
                        DataField       =   "Nama Akun"
                        Caption         =   "Deskripsi"
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
                        DataField       =   "TemplateGroup"
                        Caption         =   "TemplateGroup"
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
                           DividerStyle    =   3
                           ColumnWidth     =   1604.976
                        EndProperty
                        BeginProperty Column01 
                           DividerStyle    =   3
                           ColumnWidth     =   3000.189
                        EndProperty
                        BeginProperty Column02 
                           Object.Visible         =   0   'False
                        EndProperty
                     EndProperty
                  End
                  Begin VB.ComboBox Combo1 
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
                     ItemData        =   "FrmSetupReport.frx":6940
                     Left            =   1605
                     List            =   "FrmSetupReport.frx":694D
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   4
                     Top             =   525
                     Width           =   3570
                  End
                  Begin VB.Label Label2 
                     Alignment       =   2  'Center
                     BackColor       =   &H8000000C&
                     Caption         =   "T E M P L A T E   A R U S   K A S"
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
                     Index           =   3
                     Left            =   90
                     TabIndex        =   15
                     Top             =   165
                     Width           =   5085
                  End
                  Begin VB.Label Label3 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Kelompok Arus Kas"
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
                     Left            =   90
                     TabIndex        =   14
                     Top             =   585
                     Width           =   1350
                  End
               End
            End
            Begin VB.Frame Frame1 
               Height          =   1965
               Left            =   -74895
               TabIndex        =   9
               Top             =   3405
               Width           =   10170
               Begin VB.CommandButton CmdSet 
                  Caption         =   "Simpan"
                  Enabled         =   0   'False
                  Height          =   450
                  Index           =   1
                  Left            =   2280
                  TabIndex        =   21
                  Top             =   1440
                  Width           =   2085
               End
               Begin VB.CommandButton CmdSet 
                  Caption         =   "Seting Relasi"
                  Height          =   450
                  Index           =   0
                  Left            =   165
                  TabIndex        =   20
                  Top             =   1440
                  Width           =   2085
               End
               Begin MSDataListLib.DataCombo DataCombo1 
                  DataField       =   "No Akun"
                  Height          =   330
                  Left            =   135
                  TabIndex        =   1
                  Top             =   525
                  Width           =   4125
                  _ExtentX        =   7276
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  Style           =   2
                  ListField       =   "Nama Akun"
                  BoundColumn     =   "No Akun"
                  Text            =   "DataCombo1"
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No Akun Relasi"
                  Height          =   210
                  Index           =   0
                  Left            =   135
                  TabIndex        =   11
                  Top             =   240
                  Width           =   1200
               End
               Begin VB.Label lblAkun 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No Akun"
                  Height          =   210
                  Left            =   135
                  TabIndex        =   10
                  Top             =   960
                  Width           =   705
               End
            End
            Begin MSDataGridLib.DataGrid DgControl 
               Height          =   4575
               Left            =   -74925
               TabIndex        =   24
               Top             =   435
               Width           =   10890
               _ExtentX        =   19209
               _ExtentY        =   8070
               _Version        =   393216
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
                  DataField       =   "Job No"
                  Caption         =   "Job No"
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
                  DataField       =   "Description"
                  Caption         =   "Description"
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
                     Alignment       =   2
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin MSDataGridLib.DataGrid DgSeting 
               Height          =   1890
               Left            =   -74895
               TabIndex        =   29
               Top             =   705
               Width           =   10920
               _ExtentX        =   19262
               _ExtentY        =   3334
               _Version        =   393216
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               HeadLines       =   2
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
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "No Akun"
                  Caption         =   "Kode Account"
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
                  DataField       =   "Nama Akun"
                  Caption         =   "Deskripsi"
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
                  DataField       =   "Kelompok Perkiraan"
                  Caption         =   "Neraca/Rugi Laba"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "Balance Sheet"
                     FalseValue      =   "Profit & Loss"
                     NullValue       =   "Profit & Loss"
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  BeginProperty Column00 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column01 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   2
                     DividerStyle    =   6
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H8000000C&
               Caption         =   "S E T I N G  G R O U P  N E R A C A  D A N  R U G I / L A B A"
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
               Left            =   -74895
               TabIndex        =   30
               Top             =   405
               Width           =   10920
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H8000000C&
               Caption         =   "Seting Relasi N E R A C A Terhadap  R U G I  L A B A"
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
               Height          =   270
               Index           =   0
               Left            =   -74895
               TabIndex        =   12
               Top             =   3120
               Width           =   10920
            End
         End
      End
   End
End
Attribute VB_Name = "FrmSetupReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcAccount As New DBQuick
Private RcSeting As New DBQuick
Private RcRelasi As New DBQuick
Private RcRls As New DBQuick
Private RcSetup As New DBQuick
Private RcControl As New DBQuick
Private RcFilter As New DBQuick
Private RcPosition As New DBQuick
Private RcPrefix As New DBQuick

Private OldState As String
Private Sub CmdControl_Click(Index As Integer)
Select Case Index
       Case 0:
            CmdControl(0).Enabled = False
            CmdControl(1).Enabled = True
       Case 1:
            CmdControl(0).Enabled = True
            CmdControl(1).Enabled = False
            
End Select
End Sub

Private Sub CmdGrp_Click(Index As Integer)
If Index = 0 Then
   CmdGrp(0).Enabled = False
   CmdGrp(1).Enabled = True
   DgSeting.Columns(2).Button = True
   'DgSeting.Columns(3).Button = True
Else
   CmdGrp(0).Enabled = True
   CmdGrp(1).Enabled = False
   DgSeting.Columns(2).Button = False
   'DgSeting.Columns(3).Button = False
End If
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

Private Sub CmdKonfig_Click(Index As Integer)
Dim I As Integer
If Index = 0 Then
   CmdKonfig(0).Enabled = False
   CmdKonfig(1).Enabled = True
   DgAccount.Columns(0).Button = True
   GrdArusKas(SSTab2.Tab).Columns(0).Button = True
Else
   CmdKonfig(1).Enabled = False
   CmdKonfig(0).Enabled = True
   DgAccount.Columns(0).Button = False
   GrdArusKas(SSTab2.Tab).Columns(0).Button = False
   If SSTab2.Tab = 0 Then
      With RcSetup.DBRecordset
           .MoveFirst
           Do
             If .EOF Then Exit Do
             SendDataToServer ("Update [Tabel Pembantu] Set [Seting ArusKas] = 0,[Label ArusKas] = null where left(NoAccount,6)=N'" & Left(.Fields(0), 6) & "'")
             SendDataToServer ("Update [Tabel Pembantu] Set [Seting ArusKas] = 1,[Label ArusKas] = N'" & Combo1.Text & "' where left(NoAccount,6)=N'" & Left(.Fields(0), 6) & "'")
             .MoveNext
           Loop
           .MoveFirst
      End With
   Else
      With RcSetup.DBRecordset
           .MoveFirst
           Do
             If .EOF Then Exit Do
             SendDataToServer ("Update [Tabel Pembantu] Set [Kelompok Modal] = 0 where left(NoAccount,6)=N'" & Left(.Fields(0), 6) & "'")
             SendDataToServer ("Update [Tabel Pembantu] Set [Kelompok Modal] = 1 where left(NoAccount,6)=N'" & Left(.Fields(0), 6) & "'")
             .MoveNext
           Loop
           .MoveFirst
      End With
   End If
End If
End Sub

Private Sub CmdLevel_Click(Index As Integer)
Select Case Index
       Case 0:
            If RcPrefix.DBRecordset.Recordcount <> 0 Then
               CmdLevel(0).Enabled = False
               CmdLevel(1).Enabled = True
            End If
       Case 1:
            CmdLevel(0).Enabled = True
            CmdLevel(1).Enabled = False
End Select
DgPrefix.Refresh
End Sub

Private Sub CmdPos_Click(Index As Integer)
Select Case Index
       Case 0:
            If RcPosition.DBRecordset.Recordcount <> 0 Then
               CmdPos(0).Enabled = False
               CmdPos(1).Enabled = True
               DgPosition.Columns(2).Button = True
            End If
       Case 1:
            CmdPos(0).Enabled = True
            CmdPos(1).Enabled = False
            DgPosition.Columns(2).Button = False
End Select
DgPosition.Refresh
End Sub

Private Sub CmdSet_Click(Index As Integer)
If Index = 0 Then
   CmdSet(0).Enabled = False
   CmdSet(1).Enabled = True
   DataCombo1.Enabled = True
Else
   CmdSet(0).Enabled = True
   CmdSet(1).Enabled = False
   DataCombo1.Enabled = False
   SendDataToServer ("  Update [Tabel Pembantu] Set [Seting Relasi] = 0  ")
   SendDataToServer ("  Update [Tabel Pembantu] Set [Seting Relasi] = 1  Where NoAccount=N'" & DataCombo1.BoundText & "'")
End If
End Sub

Private Sub Combo1_Click()
If Combo1 <> "" Then
   RcSetup.DBOpen "SELECT     [Tabel Pembantu].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun] FROM         [Tabel Pembantu] INNER JOIN                       GlAccount ON [Tabel Pembantu].NoAccount = GlAccount.NoAccount WHERE     ([Tabel Pembantu].[Seting ArusKas] = 1) AND (GlAccount.[Group] = N'Sub Account') AND ([Tabel Pembantu].[Label ArusKas] = N'" & Combo1 & "') ORDER BY [Tabel Pembantu].NoAccount", Cnn, lckLockBatch
Else
   RcSetup.DBOpen " SELECT [Tabel Pembantu].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun] FROM [Tabel Pembantu] INNER JOIN GlAccount ON [Tabel Pembantu].NoAccount = GlAccount.NoAccount WHERE     ([Tabel Pembantu].[Seting ArusKas] = 1) AND (GlAccount.[Group] = N'Sub Account') ORDER BY [Tabel Pembantu].NoAccount", Cnn, lckLockBatch
End If
Set GrdArusKas(SSTab2.Tab).DataSource = RcSetup.DBRecordset
End Sub

Private Sub DataCombo1_Click(Area As Integer)
lblAkun = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
OpenPOS Left(DataCombo2.BoundText, 2)
End Sub

Private Sub DgAccount_ButtonClick(ByVal ColIndex As Integer)
Select Case SSTab2.Tab
       Case 0:
            If RcAccount.DBRecordset.Recordcount <> 0 And CmdKonfig(1).Enabled = True Then
               If FindOwnRecordset(RcSetup.DBRecordset, "[No Akun] = '" & RcAccount.DBRecordset.Fields(0) & "'") = False Then
                  RcSetup.DBRecordset.AddNew 0, RcAccount.Fields(0)
                  RcSetup.DBRecordset.Fields(1) = RcAccount.Fields(1)
                  'RcSetup.DBRecordset.Fields(2) = True
                  'SendDataToServer ("Update [Tabel Pembantu] Set [Seting Neraca] =1 where NoAccount=N'" & RcAccount.DBRecordset.Fields(0) & "' ")
               Else
                  MessageBox "Data Sudah Ada.", "Peringatan", msgOkOnly
               End If
            End If
       Case 1:
            If RcAccount.DBRecordset.Recordcount <> 0 And CmdKonfig(1).Enabled = True Then
               RcSetup.DBRecordset.AddNew 0, RcAccount.Fields(0)
               RcSetup.DBRecordset.Fields(1) = RcAccount.Fields(1)
            End If
End Select
End Sub

Private Sub DgAccount_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DgAccount.Col = 0 And CmdKonfig(1).Enabled = True Then
   DgAccount.MarqueeStyle = dbgFloatingEditor
Else
   DgAccount.MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub DgControl_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 1 And CmdControl(0).Enabled = False Then
   If RcControl.DBRecordset.Recordcount <> 0 Then
      SendDataToServer " UPDATE AccType" & _
                       " SET Tipe =N'" & ValidString(Left(RcControl.DBRecordset.Fields("Description"), 50)) & "'" & _
                       " WHERE  (ID = " & RcControl.DBRecordset.Fields("Job No") & ")"
      If OldState <> "" Then
         SendDataToServer " UPDATE GlAccount" & _
                          " SET Type =N'" & ValidString(Left(RcControl.DBRecordset.Fields("Description"), 50)) & "'" & _
                          " WHERE  (Type = N'" & OldState & "')"
         OldState = ""
      End If
   End If
End If
End Sub

Private Sub DgControl_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 1 And CmdControl(0).Enabled = False Then OldState = RcControl.DBRecordset.Fields("Description")
End Sub

Private Sub DgControl_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DgControl_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If CmdControl(0).Enabled = False And DgControl.Col >= 1 Then
   DgControl.AllowUpdate = True
Else
   DgControl.AllowUpdate = False
End If
End Sub

Private Sub DgPosition_ButtonClick(ByVal ColIndex As Integer)
If ColIndex = 2 And CmdPos(0).Enabled = False Then
   If RcPosition.DBRecordset.Recordcount <> 0 Then
      If CBool(DgPosition.Columns(ColIndex).Value) = True Then
         DgPosition.Columns(ColIndex).Value = False
      Else
         DgPosition.Columns(ColIndex).Value = True
      End If
      RcPosition.DBRecordset.Fields(ColIndex) = CBool(DgPosition.Columns(ColIndex).Value)
      SendDataToServer (" UPDATE    GlAccount " & _
                        " Set [Default] = " & BoolToInt(CBool(RcPosition.DBRecordset.Fields(ColIndex))) & _
                        " WHERE     (LEFT(NoAccount, 4) = N'" & Left(RcPosition.DBRecordset.Fields(0), 4) & "') ")
   End If
End If
End Sub

Private Sub DgPosition_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DgPrefix_AfterColEdit(ByVal ColIndex As Integer)
If (CmdLevel(0).Enabled = False) And (ColIndex = 1 Or ColIndex = 2 Or ColIndex = 3) Then
Select Case ColIndex
       Case 1, 2, 3:

            txtAccountLen(0) = LinkAccount
            Call txtAccountLen_Change(0)
End Select
End If
End Sub

Private Sub DgPrefix_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'If (CmdLevel(0).Enabled = False) And (ColIndex = 1 Or ColIndex = 2 Or ColIndex = 3) Then
'Select Case ColIndex
'       Case 2, 3:
'            If ColIndex > 1 Then
'               If IsConfigReady = False Then
'                  MessageBox "Master Perkiraan Belum komplet.", "Peringatan", msgOkOnly
'                  Exit Sub
'               End If
'            End If
'End Select
'End If
End Sub

Private Sub DgPrefix_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If CmdLevel(0).Enabled = False And DgPrefix.Col > 0 Then
   If IsConfigReady = False And DgPrefix.Col > 1 Then
      MessageBox "Configuration account have been set.", "Account Configuration", msgOkOnly
      DgPrefix.AllowUpdate = False
   Else
      DgPrefix.AllowUpdate = True
   End If
Else
   DgPrefix.AllowUpdate = False
End If
End Sub

Private Sub DgSeting_ButtonClick(ByVal ColIndex As Integer)
Dim iVar As Integer
If (ColIndex = 2 Or ColIndex = 3) And CmdGrp(1).Enabled = True Then
   If RcSeting.DBRecordset.Fields(ColIndex) = True Then
      RcSeting.DBRecordset.Fields(ColIndex) = False
      iVar = 0
   Else
      RcSeting.DBRecordset.Fields(ColIndex) = True
      iVar = 1
   End If
   DgSeting.Columns(ColIndex).Value = RcSeting.DBRecordset.Fields(ColIndex)
   If ColIndex = 2 Then
      SendDataToServer (" Update [Tabel Pembantu]  Set [Kelompok Perkiraan] = " & iVar & " Where left(NoAccount,2)=N'" & Left(RcSeting.DBRecordset.Fields(0), 2) & "'")
   Else
      SendDataToServer (" Update [Tabel Pembantu]  Set [Kelompok Modal] = " & iVar & " Where left(NoAccount,2)=N'" & Left(RcSeting.DBRecordset.Fields(0), 2) & "'")
   End If
   
End If
End Sub

Private Sub DgSeting_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DgSeting.Col = 2 Or DgSeting.Col = 3 Then
   DgSeting.MarqueeStyle = dbgFloatingEditor
Else
   DgSeting.MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub Form_Load()
GridLayout
HiasForm Picture1, Me
CenterForm Picture2, Me
RcRelasi.DBOpen "SELECT  [Tabel Pembantu].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun] FROM         [Tabel Pembantu] INNER JOIN                       GlAccount ON [Tabel Pembantu].NoAccount = GlAccount.NoAccount", Cnn, lckLockBatch
Set DataCombo1.RowSource = RcRelasi.DBRecordset

RcFilter.DBOpen "SELECT     NoAccount, AccountName FROM         GlAccount WHERE     ([Group] = N'Group Account') ORDER BY NoAccount", Cnn, lckLockBatch
Set DataCombo2.RowSource = RcFilter.DBRecordset
DataCombo2.Text = "A"

SSTab1.Tab = 0
SSTab2.Tab = 0

lblAkun = DataCombo1.BoundText
RcRls.DBOpen "SELECT  NoAccount AS [No Akun] FROM  [Tabel Pembantu] WHERE     ([Seting Relasi] = 1) GROUP BY NoAccount", Cnn, lckLockBatch
Set DataCombo1.DataSource = RcRls.DBRecordset
With RcSeting
     .DBOpen "SELECT     [Tabel Pembantu].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun], [Tabel Pembantu].[Kelompok Perkiraan],                        [Tabel Pembantu].[Kelompok Modal] FROM         [Tabel Pembantu] INNER JOIN                       GlAccount ON [Tabel Pembantu].NoAccount = GlAccount.NoAccount WHERE     (GlAccount.[Group] = N'Group Account')", Cnn, lckLockBatch
     Set DgSeting.DataSource = .DBRecordset
End With
OpenPOS Left(DataCombo2.BoundText, 2)
lblAkun = DataCombo1.BoundText
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub Form_Resize()
'
'HiasForm Picture1, Me
'CenterForm Picture2, Me
'CmdKeluar.Top = Picture1.Height + 50 '+ (CmdKeluar.Height) + 100)
'CmdKeluar.Left = Picture1.Width - (CmdKeluar.Width + 100)
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSetupReport = Nothing
End Sub

Private Sub OpenSetup()
Select Case SSTab2.Tab
       Case 0:
            RcSetup.DBOpen " SELECT [Tabel Pembantu].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun] FROM [Tabel Pembantu] INNER JOIN GlAccount ON [Tabel Pembantu].NoAccount = GlAccount.NoAccount WHERE     ([Tabel Pembantu].[Seting ArusKas] = 1) AND (GlAccount.[Group] = N'Sub Account') ORDER BY [Tabel Pembantu].NoAccount", Cnn, lckLockBatch
            Combo1.ListIndex = 0
            Combo1_Click
       Case 1:
            RcSetup.DBOpen " SELECT     [Tabel Pembantu].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun] FROM         [Tabel Pembantu] INNER JOIN                       GlAccount ON [Tabel Pembantu].NoAccount = GlAccount.NoAccount WHERE     (GlAccount.[Group] = N'Sub Account') AND ([Tabel Pembantu].[Kelompok Modal] = 1) ORDER BY [Tabel Pembantu].NoAccount", Cnn, lckLockBatch
End Select
RcAccount.DBOpen "SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun] FROM         GlAccount WHERE     ([Group] = N'List Account')", Cnn, lckLockBatch
Set DgAccount.DataSource = RcAccount.DBRecordset
Set GrdArusKas(SSTab2.Tab).DataSource = RcSetup.DBRecordset
End Sub

Private Sub GrdArusKas_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
If RcSetup.DBRecordset.Recordcount <> 0 Then
   
   RcSetup.DBRecordset.Delete adAffectCurrent
End If
End Sub

Private Sub GrdArusKas_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub GrdArusKas_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If CmdKonfig(1).Enabled = True Then
   If GrdArusKas(SSTab2.Tab).Col = 0 Then
      GrdArusKas(SSTab2.Tab).MarqueeStyle = dbgFloatingEditor
   Else
      GrdArusKas(SSTab2.Tab).MarqueeStyle = dbgHighlightRow
   End If
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

OpenSetup
Combo1.ListIndex = 0
Combo1_Click
OpenControl
OpenPrefix
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
OpenSetup

End Sub

Private Sub OpenControl()
RcControl.DBOpen "SELECT     ID AS [Job No], Tipe AS Description FROM         AccType", Cnn, lckLockBatch
Set DgControl.DataSource = RcControl.DBRecordset
End Sub

Private Sub OpenPOS(ByVal Param As String)
RcPosition.DBOpen "SELECT     NoAccount AS Code, AccountName AS [Account Name],  [Default] AS [Position],[Group] FROM         GlAccount WHERE     (LEFT(NoAccount, 2) = N'" & Param & "')  and ([Group] = N'Detail Account') ORDER BY NoAccount", Cnn, lckLockBatch
Set DgPosition.DataSource = RcPosition.DBRecordset
End Sub

Private Sub GridLayout()
DgControl.Columns(0).Width = 1695.118
DgControl.Columns(1).Width = 8609.953
DgSeting.Columns(0).Width = 1964.976
DgSeting.Columns(1).Width = 6059.906
DgSeting.Columns(2).Width = 2324.977
End Sub

Private Sub OpenPrefix()
RcPrefix.DBOpen "SELECT     [No Index] AS [No], [Group Name], [Length Per Account] AS [Length Level], Prefix FROM         [Account Setup] ORDER BY [No Index]", Cnn, lckLockBatch
Set DgPrefix.DataSource = RcPrefix.DBRecordset
txtAccountLen(0) = LinkAccount
End Sub

Private Function LinkAccount() As String
Dim Rc As New Recordset
Dim I, j As Integer
Dim Avdata As Variant
Dim Str As String
Set Rc = RcPrefix.DBRecordset.Clone(adLockReadOnly)
Str = "1000000000000000000000000"
LinkAccount = ""
With Rc
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        j = 1
        For I = 0 To UBound(Avdata, 2)
           LinkAccount = LinkAccount & Mid(Str, j, Val(Avdata(2, I))) & Avdata(3, I)
           j = j + Val(Avdata(2, I))
           SendDataToServer (" UPDATE    [Account Setup]" & _
                             " Set [Group Name] = N'" & Avdata(1, I) & "', [Length Per Account] = " & Avdata(2, I) & ", Prefix = N'" & Avdata(3, I) & "'" & _
                             " WHERE     ([No Index] = " & Avdata(0, I) & ")")
        Next I
     End If
End With
End Function

Private Sub txtAccountLen_Change(Index As Integer)
If (CmdLevel(0).Enabled = False) Then
   If Len(txtAccountLen(0)) > 26 Then
      MessageBox "Length Chart Of Account tidak boleh lebih besar dari 25 digit", "Warning", msgOkOnly
      RcPrefix.DBRecordset.CancelBatch adAffectAllChapters
   End If
End If
End Sub


Private Function CekAccount()

End Function
