VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSetupAccount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Konfigurasi Rekening"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetupAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Tag             =   "Konfigurasi Rekening"
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   5880
      Width           =   11415
      Begin VB.CommandButton CmdKeluar 
         Caption         =   "&Keluar"
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
         Left            =   10575
         Picture         =   "FrmSetupAccount.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   0
      ScaleHeight     =   5910
      ScaleWidth      =   11400
      TabIndex        =   2
      Top             =   0
      Width           =   11430
      Begin TabDlg.SSTab SSTab1 
         Height          =   5655
         Left            =   60
         TabIndex        =   3
         Top             =   90
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   9975
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         TabsPerRow      =   8
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
         TabCaption(0)   =   "Setup"
         TabPicture(0)   =   "FrmSetupAccount.frx":834C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DgSeting"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CmdGrp(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CmdGrp(1)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Konfigurasi Alur Kas"
         TabPicture(1)   =   "FrmSetupAccount.frx":8368
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab2"
         Tab(1).Control(1)=   "DgAccount"
         Tab(1).Control(2)=   "CmdKonfig(0)"
         Tab(1).Control(3)=   "CmdKonfig(1)"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Master Rekening"
         TabPicture(2)   =   "FrmSetupAccount.frx":8384
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "PictSetup(0)"
         Tab(2).Control(1)=   "CmdControl(1)"
         Tab(2).Control(2)=   "CmdControl(0)"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Posisi"
         TabPicture(3)   =   "FrmSetupAccount.frx":83A0
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "PictSetup(1)"
         Tab(3).Control(1)=   "CmdPos(0)"
         Tab(3).Control(2)=   "CmdPos(1)"
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Level Format"
         TabPicture(4)   =   "FrmSetupAccount.frx":83BC
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "CmdLevel(4)"
         Tab(4).Control(1)=   "CmdLevel(3)"
         Tab(4).Control(2)=   "CmdLevel(2)"
         Tab(4).Control(3)=   "CmdLevel(0)"
         Tab(4).Control(4)=   "CmdLevel(1)"
         Tab(4).Control(5)=   "PictSetup(2)"
         Tab(4).Control(6)=   "txtAccountLen(0)"
         Tab(4).Control(7)=   "Line1(4)"
         Tab(4).Control(8)=   "Label1(2)"
         Tab(4).ControlCount=   9
         TabCaption(5)   =   "Basis Entri Jurnal"
         TabPicture(5)   =   "FrmSetupAccount.frx":83D8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "SSTab3"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Relasi Rekening"
         TabPicture(6)   =   "FrmSetupAccount.frx":83F4
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Picture1"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Non Relasi"
         TabPicture(7)   =   "FrmSetupAccount.frx":8410
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "PictSetup(3)"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).ControlCount=   1
         Begin VB.PictureBox PictSetup 
            Height          =   5175
            Index           =   3
            Left            =   -74925
            ScaleHeight     =   5115
            ScaleWidth      =   11025
            TabIndex        =   80
            Top             =   390
            Width           =   11085
            Begin MSDataGridLib.DataGrid GridNonRelasi 
               Height          =   5115
               Left            =   0
               TabIndex        =   81
               Top             =   0
               Width           =   11010
               _ExtentX        =   19420
               _ExtentY        =   9022
               _Version        =   393216
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               HeadLines       =   2
               RowHeight       =   15
               RowDividerStyle =   3
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
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "ID Rekening"
                  Caption         =   "ID Rekening"
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
                  DataField       =   "Tipe Rekening"
                  Caption         =   "Tipe Rekening"
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
                  DataField       =   "Kode Rekening"
                  Caption         =   "Kode Rekening"
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
                     ColumnWidth     =   1200.189
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   3495.118
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.CommandButton CmdLevel 
            Caption         =   "Batal"
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
            Height          =   390
            Index           =   4
            Left            =   -70875
            TabIndex        =   49
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdLevel 
            Caption         =   "Simpan"
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
            Height          =   390
            Index           =   3
            Left            =   -71880
            TabIndex        =   48
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdLevel 
            Caption         =   "Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   2
            Left            =   -72885
            TabIndex        =   47
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdLevel 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   -74895
            TabIndex        =   46
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdLevel 
            Caption         =   "Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   1
            Left            =   -73890
            TabIndex        =   45
            Top             =   5085
            Width           =   1000
         End
         Begin VB.PictureBox PictSetup 
            Height          =   4605
            Index           =   2
            Left            =   -74910
            ScaleHeight     =   4545
            ScaleWidth      =   10965
            TabIndex        =   43
            Top             =   390
            Width           =   11025
            Begin MSDataGridLib.DataGrid DgPrefix 
               Bindings        =   "FrmSetupAccount.frx":842C
               Height          =   4500
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   10860
               _ExtentX        =   19156
               _ExtentY        =   7938
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "No"
                  Caption         =   "No"
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
                  DataField       =   "Level"
                  Caption         =   "Level"
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
                  DataField       =   "Panjang"
                  Caption         =   "Panjang"
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
                  DataField       =   "Prefix"
                  Caption         =   "Prefix"
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
                  BeginProperty Column04 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox txtAccountLen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   -66090
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   5115
            Width           =   2100
         End
         Begin VB.CommandButton CmdPos 
            Caption         =   "Lock Grid"
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
            Height          =   390
            Index           =   1
            Left            =   -73890
            TabIndex        =   41
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdPos 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   -74895
            TabIndex        =   40
            Top             =   5085
            Width           =   1000
         End
         Begin VB.PictureBox PictSetup 
            Height          =   4605
            Index           =   1
            Left            =   -74910
            ScaleHeight     =   4545
            ScaleWidth      =   10980
            TabIndex        =   36
            Top             =   390
            Width           =   11040
            Begin MSDataGridLib.DataGrid DgPosition 
               Bindings        =   "FrmSetupAccount.frx":8441
               Height          =   4050
               Left            =   -15
               TabIndex        =   37
               Top             =   510
               Width           =   11010
               _ExtentX        =   19420
               _ExtentY        =   7144
               _Version        =   393216
               AllowUpdate     =   0   'False
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
                     ColumnWidth     =   3495.118
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   2
                  EndProperty
               EndProperty
            End
            Begin MSDataListLib.DataCombo ComboPosisi 
               Height          =   330
               Left            =   1590
               TabIndex        =   38
               Top             =   75
               Width           =   4140
               _ExtentX        =   7303
               _ExtentY        =   582
               _Version        =   393216
               Style           =   2
               ListField       =   "AccountName"
               BoundColumn     =   "NoAccount"
               Text            =   ""
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Group Account"
               Height          =   210
               Index           =   1
               Left            =   75
               TabIndex        =   39
               Top             =   135
               Width           =   1245
            End
         End
         Begin VB.PictureBox PictSetup 
            Height          =   4605
            Index           =   0
            Left            =   -74910
            ScaleHeight     =   4545
            ScaleWidth      =   10980
            TabIndex        =   34
            Top             =   390
            Width           =   11040
            Begin MSDataGridLib.DataGrid DgControl 
               Height          =   4575
               Left            =   -15
               TabIndex        =   35
               Top             =   -15
               Width           =   11010
               _ExtentX        =   19420
               _ExtentY        =   8070
               _Version        =   393216
               AllowUpdate     =   0   'False
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
               BeginProperty Column02 
                  DataField       =   "status"
                  Caption         =   "Status"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "AKTIF"
                     FalseValue      =   "NON-AKTIF"
                     NullValue       =   ""
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
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
                     Button          =   -1  'True
                     Locked          =   -1  'True
                  EndProperty
               EndProperty
            End
         End
         Begin VB.CommandButton CmdControl 
            Caption         =   "Simpan"
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
            Height          =   390
            Index           =   1
            Left            =   -73890
            TabIndex        =   33
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdControl 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   -74895
            TabIndex        =   32
            Top             =   5085
            Width           =   1000
         End
         Begin VB.CommandButton CmdKonfig 
            Caption         =   "Simpan"
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
            Height          =   390
            Index           =   1
            Left            =   -68475
            TabIndex        =   31
            Top             =   5100
            Width           =   1000
         End
         Begin VB.CommandButton CmdKonfig 
            Caption         =   "Mulai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   -69480
            TabIndex        =   30
            Top             =   5100
            Width           =   1000
         End
         Begin VB.CommandButton CmdGrp 
            Caption         =   "Simpan"
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
            Height          =   390
            Index           =   1
            Left            =   1335
            TabIndex        =   29
            Top             =   2640
            Width           =   1000
         End
         Begin VB.CommandButton CmdGrp 
            Caption         =   "Seting Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   105
            TabIndex        =   28
            Top             =   2640
            Width           =   1230
         End
         Begin VB.Frame Frame1 
            Height          =   1965
            Left            =   105
            TabIndex        =   22
            Top             =   3405
            Width           =   10920
            Begin VB.CommandButton CmdSet 
               Caption         =   "Simpan"
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
               Height          =   390
               Index           =   1
               Left            =   1305
               TabIndex        =   24
               Top             =   1440
               Width           =   1000
            End
            Begin VB.CommandButton CmdSet 
               Caption         =   "Seting Relasi"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   0
               Left            =   135
               TabIndex        =   23
               Top             =   1440
               Width           =   1170
            End
            Begin MSDataListLib.DataCombo ComboRelasi 
               DataField       =   "No Akun"
               Height          =   330
               Left            =   1425
               TabIndex        =   25
               Top             =   540
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               ListField       =   "Nama Akun"
               BoundColumn     =   "No Akun"
               Text            =   "- Rugi Laba Account - "
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Akun"
               Height          =   210
               Left            =   165
               TabIndex        =   79
               Top             =   960
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Akun Relasi"
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   27
               Top             =   600
               Width           =   915
            End
            Begin VB.Label lblAkun 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No Akun"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1440
               TabIndex        =   26
               Top             =   930
               Width           =   4050
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00EAAF6F&
            Height          =   5175
            Left            =   -74925
            ScaleHeight     =   5115
            ScaleWidth      =   11040
            TabIndex        =   4
            Top             =   390
            Width           =   11100
            Begin VB.Frame Frame3 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Relasi Data "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   4050
               TabIndex        =   11
               Top             =   4245
               Width           =   2895
               Begin VB.CommandButton CmdPanah 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   3
                  Left            =   2115
                  Picture         =   "FrmSetupAccount.frx":8456
                  Style           =   1  'Graphical
                  TabIndex        =   15
                  Tag             =   "True"
                  ToolTipText     =   "Pindah semua record"
                  Top             =   300
                  Width           =   650
               End
               Begin VB.CommandButton CmdPanah 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   2
                  Left            =   1470
                  Picture         =   "FrmSetupAccount.frx":854A
                  Style           =   1  'Graphical
                  TabIndex        =   14
                  Tag             =   "True"
                  ToolTipText     =   "Pindah 1 record"
                  Top             =   300
                  Width           =   650
               End
               Begin VB.CommandButton CmdPanah 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   1
                  Left            =   825
                  Picture         =   "FrmSetupAccount.frx":863C
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  Tag             =   "True"
                  ToolTipText     =   "Pilih semua record"
                  Top             =   300
                  Width           =   650
               End
               Begin VB.CommandButton CmdPanah 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   0
                  Left            =   180
                  Picture         =   "FrmSetupAccount.frx":8734
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  Tag             =   "True"
                  ToolTipText     =   "Pilih 1 record"
                  Top             =   300
                  Width           =   650
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00EAAF6F&
               Caption         =   " Cari Rekeninig "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   4050
               TabIndex        =   5
               Top             =   3225
               Width           =   2895
               Begin VB.OptionButton OptSearch 
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "&Nama"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   1
                  Left            =   165
                  TabIndex        =   10
                  Tag             =   "True"
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   780
               End
               Begin VB.OptionButton OptSearch 
                  BackColor       =   &H00EAAF6F&
                  Caption         =   "&Kode"
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
                  Left            =   945
                  TabIndex        =   9
                  Tag             =   "True"
                  Top             =   255
                  Width           =   975
               End
               Begin VB.CommandButton CmdFresh 
                  Height          =   330
                  Index           =   0
                  Left            =   150
                  Picture         =   "FrmSetupAccount.frx":8826
                  Style           =   1  'Graphical
                  TabIndex        =   8
                  Tag             =   "True"
                  ToolTipText     =   "Refresh"
                  Top             =   540
                  Width           =   345
               End
               Begin VB.TextBox TxtCarik 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   840
                  TabIndex        =   7
                  Tag             =   "True"
                  Top             =   540
                  Width           =   1935
               End
               Begin VB.CommandButton CmdFresh 
                  Height          =   330
                  Index           =   1
                  Left            =   495
                  Picture         =   "FrmSetupAccount.frx":F078
                  Style           =   1  'Graphical
                  TabIndex        =   6
                  Tag             =   "True"
                  ToolTipText     =   "Go"
                  Top             =   540
                  Width           =   345
               End
            End
            Begin MSDataGridLib.DataGrid GridRelasi 
               Height          =   4785
               Index           =   1
               Left            =   6990
               TabIndex        =   16
               Top             =   255
               Width           =   3960
               _ExtentX        =   6985
               _ExtentY        =   8440
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   1
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
                  BeginProperty Column00 
                     ColumnWidth     =   1005.165
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   2399.811
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ListView LView 
               Height          =   2925
               Left            =   4065
               TabIndex        =   17
               Top             =   255
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   5159
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   0
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Master Rekening"
                  Object.Width           =   3528
               EndProperty
            End
            Begin MSDataGridLib.DataGrid GridRelasi 
               Height          =   4785
               Index           =   0
               Left            =   45
               TabIndex        =   18
               Top             =   255
               Width           =   3960
               _ExtentX        =   6985
               _ExtentY        =   8440
               _Version        =   393216
               AllowUpdate     =   -1  'True
               HeadLines       =   1
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
                  BeginProperty Column00 
                     ColumnWidth     =   1005.165
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   2399.811
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label2 
               BackColor       =   &H8000000C&
               BackStyle       =   0  'Transparent
               Caption         =   "List Rekening Relasi"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   240
               Index           =   6
               Left            =   6990
               TabIndex        =   21
               Top             =   15
               Width           =   3960
            End
            Begin VB.Label Label2 
               BackColor       =   &H8000000C&
               BackStyle       =   0  'Transparent
               Caption         =   "Master Rekening"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   240
               Index           =   5
               Left            =   4095
               TabIndex        =   20
               Top             =   15
               Width           =   2835
            End
            Begin VB.Label Label2 
               BackColor       =   &H8000000C&
               BackStyle       =   0  'Transparent
               Caption         =   "List Rekening Non Relasi"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   240
               Index           =   4
               Left            =   105
               TabIndex        =   19
               Top             =   15
               Width           =   3915
            End
         End
         Begin MSDataGridLib.DataGrid DgAccount 
            Height          =   5055
            Left            =   -74910
            TabIndex        =   50
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
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   3
               EndProperty
            EndProperty
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   4560
            Left            =   -69480
            TabIndex        =   51
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Arus Kas"
            TabPicture(0)   =   "FrmSetupAccount.frx":158CA
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame5"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Perubahan Modal"
            TabPicture(1)   =   "FrmSetupAccount.frx":158E6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame2"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame2 
               Height          =   4170
               Left            =   -74925
               TabIndex        =   57
               Top             =   315
               Width           =   5340
               Begin MSDataGridLib.DataGrid GrdArusKas 
                  Height          =   3615
                  Index           =   1
                  Left            =   75
                  TabIndex        =   58
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
                     EndProperty
                     BeginProperty Column01 
                        DividerStyle    =   3
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
                  TabIndex        =   59
                  Top             =   165
                  Width           =   5085
               End
            End
            Begin VB.Frame Frame5 
               Height          =   4170
               Left            =   75
               TabIndex        =   52
               Top             =   315
               Width           =   5280
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
                  ItemData        =   "FrmSetupAccount.frx":15902
                  Left            =   1605
                  List            =   "FrmSetupAccount.frx":1590F
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   53
                  Top             =   525
                  Width           =   3570
               End
               Begin MSDataGridLib.DataGrid GrdArusKas 
                  Height          =   3150
                  Index           =   0
                  Left            =   90
                  TabIndex        =   54
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
                     EndProperty
                     BeginProperty Column01 
                        DividerStyle    =   3
                        ColumnWidth     =   2880
                     EndProperty
                     BeginProperty Column02 
                        Object.Visible         =   0   'False
                     EndProperty
                  EndProperty
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
                  TabIndex        =   56
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
                  TabIndex        =   55
                  Top             =   585
                  Width           =   1350
               End
            End
         End
         Begin MSDataGridLib.DataGrid DgSeting 
            Height          =   1890
            Left            =   105
            TabIndex        =   60
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
         Begin TabDlg.SSTab SSTab3 
            Height          =   5055
            Left            =   -74835
            TabIndex        =   61
            Top             =   435
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   8916
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Penerimaan"
            TabPicture(0)   =   "FrmSetupAccount.frx":15983
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture4"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Pengeluaran"
            TabPicture(1)   =   "FrmSetupAccount.frx":1599F
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Picture5"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.PictureBox Picture5 
               Height          =   4575
               Left            =   90
               ScaleHeight     =   4515
               ScaleWidth      =   10665
               TabIndex        =   69
               Top             =   390
               Width           =   10725
               Begin MSDataGridLib.DataGrid PaymentA 
                  Height          =   4530
                  Left            =   -15
                  TabIndex        =   75
                  Top             =   -15
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   7990
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
                     DataField       =   "No ID"
                     Caption         =   "No"
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
                     DataField       =   "Description"
                     Caption         =   "Tipe"
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
                     MarqueeStyle    =   4
                     ScrollBars      =   2
                     BeginProperty Column00 
                        ColumnWidth     =   1005.165
                     EndProperty
                     BeginProperty Column01 
                        ColumnWidth     =   3075.024
                     EndProperty
                  EndProperty
               End
               Begin MSDataGridLib.DataGrid PaymentB 
                  Height          =   4530
                  Left            =   6045
                  TabIndex        =   74
                  Top             =   -15
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   7990
                  _Version        =   393216
                  AllowUpdate     =   0   'False
                  BorderStyle     =   0
                  HeadLines       =   1
                  RowHeight       =   16
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
                     DataField       =   "No ID"
                     Caption         =   "No"
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
                     DataField       =   "Description"
                     Caption         =   "Tipe"
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
                     MarqueeStyle    =   4
                     ScrollBars      =   2
                     BeginProperty Column00 
                        ColumnWidth     =   1005.165
                     EndProperty
                     BeginProperty Column01 
                        ColumnWidth     =   3075.024
                     EndProperty
                  EndProperty
               End
               Begin VB.CommandButton CmdPanahPay 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Index           =   0
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":159BB
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   1065
                  Width           =   420
               End
               Begin VB.CommandButton CmdPanahPay 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Index           =   1
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":15AAD
                  Style           =   1  'Graphical
                  TabIndex        =   72
                  Top             =   1545
                  Width           =   420
               End
               Begin VB.CommandButton CmdPanahPay 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Index           =   2
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":15BA5
                  Style           =   1  'Graphical
                  TabIndex        =   71
                  Top             =   2025
                  Width           =   420
               End
               Begin VB.CommandButton CmdPanahPay 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Index           =   3
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":15C97
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  Top             =   2505
                  Width           =   420
               End
            End
            Begin VB.PictureBox Picture4 
               Height          =   4575
               Left            =   -74910
               ScaleHeight     =   4515
               ScaleWidth      =   10665
               TabIndex        =   62
               Top             =   390
               Width           =   10725
               Begin MSDataGridLib.DataGrid DataGrid2 
                  Height          =   4530
                  Left            =   6045
                  TabIndex        =   67
                  Top             =   0
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   7990
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
                     DataField       =   "No ID"
                     Caption         =   "No"
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
                     DataField       =   "Description"
                     Caption         =   "Tipe"
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
                     MarqueeStyle    =   4
                     ScrollBars      =   2
                     BeginProperty Column00 
                        ColumnWidth     =   1005.165
                     EndProperty
                     BeginProperty Column01 
                        ColumnWidth     =   3075.024
                     EndProperty
                  EndProperty
               End
               Begin MSDataGridLib.DataGrid DataGrid1 
                  Height          =   4530
                  Left            =   0
                  TabIndex        =   68
                  Top             =   0
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   7990
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
                     DataField       =   "No ID"
                     Caption         =   "No"
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
                     DataField       =   "Description"
                     Caption         =   "Tipe"
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
                     MarqueeStyle    =   4
                     ScrollBars      =   2
                     BeginProperty Column00 
                        ColumnWidth     =   1005.165
                     EndProperty
                     BeginProperty Column01 
                        ColumnWidth     =   3075.024
                     EndProperty
                  EndProperty
               End
               Begin VB.CommandButton CmdPanahRcv4 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":15D8B
                  Style           =   1  'Graphical
                  TabIndex        =   66
                  Top             =   2505
                  Width           =   420
               End
               Begin VB.CommandButton CmdPanahRcv3 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":15E7F
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  Top             =   2025
                  Width           =   420
               End
               Begin VB.CommandButton CmdPanahRcv2 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":15F71
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  Top             =   1545
                  Width           =   420
               End
               Begin VB.CommandButton CmdPanahRcv1 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   5130
                  Picture         =   "FrmSetupAccount.frx":16069
                  Style           =   1  'Graphical
                  TabIndex        =   63
                  Top             =   1065
                  Width           =   420
               End
            End
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   -66825
            X2              =   -65825
            Y1              =   5415
            Y2              =   5415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Format"
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
            Left            =   -66795
            TabIndex        =   78
            Top             =   5160
            Width           =   510
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000C&
            Caption         =   "NERACA / RUGI LABA"
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
            Left            =   105
            TabIndex        =   77
            Top             =   405
            Width           =   10920
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000C&
            Caption         =   "RELASI RUGI LABA"
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
            Left            =   105
            TabIndex        =   76
            Top             =   3120
            Width           =   10920
         End
      End
   End
End
Attribute VB_Name = "FrmSetupAccount"
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
Private RcFilters As New DBQuick
Private RcFiltersReceipt As New DBQuick
Private RcFilters2 As New DBQuick
Private RcFiltersPayment As New DBQuick
Private RcFilters3 As New DBQuick
Private RcFiltersGoods As New DBQuick
Private RcRelasiGrid1 As New DBQuick
Private RcRelasiGrid2 As New DBQuick
Private OldState As String
Dim strSQL As String

Private Sub OpenRelasi()
On Error GoTo RelErr
Dim I As Integer
Dim Avdata As Variant

If RcControl.DBRecordset.Recordcount <> 0 Then
    RcControl.DBRecordset.MoveFirst
    With RcControl.DBRecordset
        LView.ListItems.Clear
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            With LView.ListItems.Add(, , Avdata(0, I))
               .SubItems(1) = Avdata(1, I)
               If Avdata(2, I) = True Then
                .ForeColor = &H80000012
                
'                LView.ForeColor = &H80000012
               Else
                .ForeColor = &HFF&
                .Bold = True
'                LView.ForeColor = &HFF&
               End If
            End With
        Next I
    End With
Else
    Exit Sub
End If

Exit Sub
RelErr:
    MessageBox Err.Description, "Relasi Rekening", msgOkOnly, msgCrtical
End Sub

Private Sub LoadRelasi(Optional nID As Integer = 0)
If nID = 0 Then
   RcRelasiGrid2.DBOpen "SELECT NoAccount, AccountName, Type From GLAccount WHERE (Type IS NOT NULL)", CNN, lckLockBatch, lckLockSync
Else
   RcRelasiGrid2.DBOpen "SELECT NoAccount, AccountName, Type From GLAccount WHERE (Type IS NOT NULL) and ID= " & nID, CNN, lckLockBatch, lckLockSync
End If
Set GridRelasi(1).DataSource = RcRelasiGrid2.DBRecordset
RcRelasiGrid1.DBOpen "SELECT NoAccount, AccountName, Type From GLAccount WHERE (Type IS NULL)", CNN, lckLockBatch, lckLockSync
Set GridRelasi(0).DataSource = RcRelasiGrid1.DBRecordset

'With GridRelasi(1)
'    .Height = 4785
'    .Columns(0).width = 1000
'    .Columns(1).width = 2400
'End With

End Sub

Private Sub CmdFresh_Click(Index As Integer)
   Select Case Index
      Case 0
         TxtCarik.Text = ""
         OpenControl
         OpenRelasi
      Case 1
         If Trim(TxtCarik) <> "" Then
            If OptSearch(0).Value = True Then
               OpenControl "ID", TxtCarik.Text
            Else
               OpenControl "Tipe", TxtCarik.Text
            End If
            OpenRelasi
         End If
   End Select
End Sub

Private Sub CmdPanah_Click(Index As Integer)
Dim bMark As Variant
   Select Case Index
      Case 0
         SendDataToServer "UPDATE GLAccount set type='" & LView.SelectedItem.SubItems(1) & "',ID=" & LView.SelectedItem.Text & " where NoAccount='" & RcRelasiGrid1.DBRecordset.Fields("noAccount") & "'"
      Case 1
         With RcRelasiGrid1.DBRecordset
            .MoveFirst
            While Not .EOF
               SendDataToServer "UPDATE GLAccount set type='" & LView.SelectedItem.SubItems(1) & "',ID=" & LView.SelectedItem.Text & " where NoAccount='" & .Fields("noAccount") & "'"
               .MoveNext
            Wend
         End With
      Case 2
         SendDataToServer "UPDATE GLAccount set type=NULL,ID=0 where NoAccount='" & RcRelasiGrid2.DBRecordset.Fields("noAccount") & "'"
      Case 3
         With RcRelasiGrid2.DBRecordset
            .MoveFirst
            While Not .EOF
               SendDataToServer "UPDATE GLAccount set type=NULL,ID=0 where NoAccount='" & .Fields("noAccount") & "'"
               .MoveNext
            Wend
         End With
   End Select
'    bMark = RcRelasiGrid2.DBRecordset.Bookmark
    LoadRelasi LView.SelectedItem.Text
'    RcRelasiGrid2.DBRecordset.Bookmark = bMark
End Sub

Private Sub DgControl_ButtonClick(ByVal ColIndex As Integer)
   If CmdControl(1).Enabled = True Then
      If DgControl.Columns(2).Value = True Then
         DgControl.Columns(2).Value = False
      Else
         DgControl.Columns(2).Value = True
      End If
   End If
End Sub

Private Sub Form_Load()
GridLayout
HiasFormManTell Picture2, Me
RcRelasi.DBOpen "SELECT  [Tabel Pembantu].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun] " & _
                " FROM  [Tabel Pembantu] INNER JOIN GLAccount ON [Tabel Pembantu].NoAccount = GLAccount.NoAccount ORDER BY  GLAccount.AccountName", CNN, lckLockBatch
Set ComboRelasi.RowSource = RcRelasi.DBRecordset

RcFilter.DBOpen "SELECT     NoAccount, AccountName FROM         GLAccount WHERE     ([Group] = N'Group Account') ORDER BY NoAccount", CNN, lckLockBatch
Set ComboPosisi.RowSource = RcFilter.DBRecordset
ComboPosisi.Text = "A"

SSTab1.Tab = 0
SSTab2.Tab = 0

lblAkun = ComboRelasi.BoundText
RcRls.DBOpen "SELECT  NoAccount AS [No Akun] FROM  [Tabel Pembantu] WHERE ([Seting Relasi] = 1) GROUP BY NoAccount", CNN, lckLockBatch
Set ComboRelasi.DataSource = RcRls.DBRecordset
With RcSeting
     .DBOpen "SELECT     [Tabel Pembantu].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun], [Tabel Pembantu].[Kelompok Perkiraan],                        [Tabel Pembantu].[Kelompok Modal] FROM         [Tabel Pembantu] INNER JOIN                       GLAccount ON [Tabel Pembantu].NoAccount = GLAccount.NoAccount WHERE     (GLAccount.[Group] = N'Group Account')", CNN, lckLockBatch
     Set DgSeting.DataSource = .DBRecordset
End With
OpenPOS Left(ComboPosisi.BoundText, 2)
lblAkun = ComboRelasi.BoundText

RcFilters.DBOpen "SELECT ID AS [No ID], Tipe AS Description, receipt_group FROM AccType WHERE (receipt_group = N'') ", CNN, lckLockBatch, lckLockSync
Set DataGrid1.DataSource = RcFilters.DBRecordset

RcFiltersReceipt.DBOpen " Select  ID AS [No ID], Description, [Group Type] FROM [Table Filter Account] WHERE ([Table Filter Account].[Group Type] = N'RECEIPT') order by ID", CNN, lckLockBatch, lckLockSync
Set DataGrid2.DataSource = RcFiltersReceipt.DBRecordset

RcFilters2.DBOpen "SELECT ID AS [No ID], Tipe AS Description, payment_group FROM AccType WHERE (payment_group = N'') ", CNN, lckLockBatch, lckLockSync
Set PaymentA.DataSource = RcFilters2.DBRecordset

RcFiltersPayment.DBOpen " Select  ID AS [No ID], Description, [Group Type] FROM [Table Filter Account] WHERE ([Table Filter Account].[Group Type] = N'PAYMENT') order by ID", CNN, lckLockBatch, lckLockSync
Set PaymentB.DataSource = RcFiltersPayment.DBRecordset

RcFilters3.DBOpen "SELECT AccType.ID as [ID Rekening], AccType.Tipe AS [Tipe Rekening], GLAccount.NoAccount AS [Kode Rekening] " & _
                " FROM GLAccount RIGHT OUTER JOIN AccType ON GLAccount.ID = AccType.ID " & _
                " WHERE (GLAccount.ID IS NULL) AND (AccType.status = 1)", CNN, lckLockBatch, lckLockSync
'Debug.Print RcFilters3.DBRecordset.Source
Set GridNonRelasi.DataSource = RcFilters3.DBRecordset
'
'RcFiltersGoods.DBOpen " Select  ID AS [No ID], Description, [Group Type] FROM [Table Filter Account] WHERE ([Table Filter Account].[Group Type] = N'GOODS') order by ID", CNN, lckLockBatch, lckLockSync
'Set GridGoods(1).DataSource = RcFiltersGoods.DBRecordset



'********* Relasi Rekening ***********

LoadRelasi

'*************************************
Label2(0).FontBold = True
Label2(1).FontBold = True
'Label2(0).BackColor = &H8000000C
'Label2(1).BackColor = &H8000000C


End Sub
Private Sub UpdateMasterRekening()
   Dim isTrue As String
   Dim Avdata As Variant
   Dim Rc As New Recordset
   With RcControl.DBRecordset
      .MoveFirst
      While Not .EOF
         isTrue = IIf(.Fields("status") = True, "1", "0")
         SendDataToServer "update accType set status =" & isTrue & " where ID='" & .Fields("Job No") & "'"
         .MoveNext
      Wend
   End With
End Sub

Private Sub CmdControl_Click(Index As Integer)
Select Case Index
       Case 0:
            CmdControl(0).Enabled = False
            CmdControl(1).Enabled = True
       Case 1:
            CmdControl(0).Enabled = True
            CmdControl(1).Enabled = False
            UpdateMasterRekening
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
           If .Recordcount <> 0 Then .MoveFirst
           Do
             If .EOF Then Exit Do
             SendDataToServer ("Update [Tabel Pembantu] Set [Seting ArusKas] = 0,[Label ArusKas] = null where left(NoAccount,6)=N'" & Left(.Fields(0), 6) & "'")
             SendDataToServer ("Update [Tabel Pembantu] Set [Seting ArusKas] = 1,[Label ArusKas] = N'" & Combo1.Text & "' where left(NoAccount,6)=N'" & Left(.Fields(0), 6) & "'")
             .MoveNext
           Loop
           If .Recordcount <> 0 Then .MoveFirst
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
    Case 0: 'EDIT
        If RcPrefix.DBRecordset.Recordcount <> 0 Then
            CmdLevel(0).Enabled = False
            CmdLevel(1).Enabled = False
            CmdLevel(2).Enabled = False
            CmdLevel(3).Enabled = True
            CmdLevel(4).Enabled = True
        End If
    Case 1  'TAMBAH
            CmdLevel(0).Enabled = False
            CmdLevel(1).Enabled = False
            CmdLevel(2).Enabled = False
            CmdLevel(3).Enabled = True
            CmdLevel(4).Enabled = True
            RcPrefix.DBRecordset.AddNew
            RcPrefix.DBRecordset.Fields("no") = RcPrefix.DBRecordset.Recordcount
            RcPrefix.DBRecordset.Fields("panjang") = 0
            RcPrefix.DBRecordset.Fields("keterangan") = ""
            RcPrefix.DBRecordset.Fields("level") = ""
            RcPrefix.DBRecordset.Fields("prefix") = ""
            
            
    Case 2  'HAPUS
         If MessageBox("Yakin data akan dihapus ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
            SendDataToServer "delete from [account setup] where [no Index]=" & DgPrefix.Columns(0).Value
            CmdLevel(0).Enabled = True
            CmdLevel(1).Enabled = True
            CmdLevel(2).Enabled = True
            CmdLevel(3).Enabled = False
            CmdLevel(4).Enabled = False
            RcPrefix.DBRecordset.Delete adAffectCurrent
            RcPrefix.DBRecordset.UpdateBatch adAffectCurrent
         End If
         
    Case 3  'SIMPAN
        If MessageBox("Simpan data", "Konfigurasi Rekening", msgYesNo, msgQuestion) = 1 Then
            CmdLevel(0).Enabled = True
            CmdLevel(1).Enabled = True
            CmdLevel(2).Enabled = True
            CmdLevel(3).Enabled = False
            CmdLevel(4).Enabled = False
'            Debug.Print strSQL
            With RcPrefix.DBRecordset
               .MoveFirst
               While Not .EOF
                  SendDataToServer BuildSQL
                  .MoveNext
               Wend
            End With
        End If
        
    Case 4  'BATAL
      CmdLevel(0).Enabled = True
      CmdLevel(1).Enabled = True
      CmdLevel(2).Enabled = True
      CmdLevel(3).Enabled = False
      CmdLevel(4).Enabled = False
      RcPrefix.DBRecordset.CancelBatch adAffectCurrent

End Select

DgPrefix.Refresh
End Sub

Private Function BuildSQL() As String
   Dim strSQL As String
   Dim rsCek As New DBQuick
   Dim sPrefix As String
   With RcPrefix.DBRecordset
         sPrefix = GetPrefixValue
         rsCek.DBOpen "select * from [Account Setup] where ([No Index] = " & DgPrefix.Columns(0).Value & ")", CNN
         If rsCek.DBRecordset.Recordcount > 0 Then
            strSQL = " UPDATE [Account Setup] Set [Index Group] = " & FNumText(DgPrefix.Columns(1).Value) & ", [Group Name] = " & FNumText(DgPrefix.Columns(2).Value) & ",[Length Per Account] = " & FQty(DgPrefix.Columns(3).Value) & ",Prefix = " & sPrefix & _
                     " WHERE ([No Index] = " & DgPrefix.Columns(0).Value & ")"
         Else
            strSQL = "insert into [account setup] ([No Index],[index group],[group name],[Length Per Account],Prefix) values (" & DgPrefix.Columns(0).Value & "," & FNumText(DgPrefix.Columns(1).Value) & "," & FNumText(DgPrefix.Columns(2).Value) & "," & FQty(DgPrefix.Columns(3).Value) & "," & sPrefix & ")"
         End If
   End With
   rsCek.CloseDB
   BuildSQL = strSQL
End Function

Private Function GetPrefixValue() As String
On Error GoTo xErr
   If Len(DgPrefix.Columns(4).Value) > 0 Then
      GetPrefixValue = "'" & DgPrefix.Columns(4).Value & "'"
   Else
      GetPrefixValue = "''"
   End If
Exit Function
xErr:
   If Err.Number = 13 Then
      Err.Clear
      GetPrefixValue = "''"
   Else
      MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
   End If
End Function

'Private Sub PrepareQuery()
'With MyDDE
'    .PrepareAppend = " INSERT INTO [Transport] (ID, Expedisi, Address, Person, Phone, Fax, Type) " & _
'                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "', N'" & ValidString(txtBox(3)) & "', N'" & ValidString(txtBox(4)) & "', N'" & ValidString(txtBox(5)) & "','" & Mtype & "')"
'
'    .PrepareUpdate = " UPDATE [Transport] Set [Expedisi] = N'" & ValidString(txtBox(1)) & "',[Address] = N'" & ValidString(txtBox(2)) & "',[Person] = N'" & ValidString(txtBox(3)) & "',[Phone] = N'" & ValidString(txtBox(4)) & "',[fax] = N'" & ValidString(txtBox(5)) & "' WHERE     (ID = N'" & ValidString(txtBox(0)) & "') and (type ='" & Mtype & "')"
'
'    .PrepareDelete = " DELETE FROM [Transport] WHERE   (ID = N'" & ValidString(txtBox(0)) & "') and (Type='" & Mtype & "') "
'End With
'End Sub

'Private Sub CmdPanahGoods_Click(Index As Integer)
'On Error Resume Next
'Select Case Index
'    Case 0
'        If RcFilters3.DBRecordset.Recordcount <> 0 Then
'            RcFiltersGoods.DBRecordset.AddNew 0, RcFilters3.Fields(0)
'            RcFiltersGoods.DBRecordset.Fields(1) = RcFilters3.Fields(1)
'            RcFiltersGoods.DBRecordset.Fields("Group Type") = "GOODS"
'            RcFiltersGoods.DBRecordset.UpdateBatch adAffectCurrent
'
'            RcFilters3.DBRecordset.Fields("goods_group") = "GOODS"
'            RcFilters3.DBRecordset.UpdateBatch adAffectCurrent
'            RcFilters3.DBRecordset.Delete adAffectCurrent
'        End If
'    Case 1
'        If RcFilters3.DBRecordset.Recordcount <> 0 Then
'           RcFilters3.DBRecordset.MoveFirst
'           Do
'             If RcFilters3.DBRecordset.EOF Then Exit Do
'                RcFiltersGoods.DBRecordset.AddNew 0, RcFilters3.Fields(0)
'                RcFiltersGoods.DBRecordset.Fields(1) = RcFilters3.Fields(1)
'                RcFiltersGoods.DBRecordset.Fields("Group Type") = "GOODS"
'                RcFiltersGoods.DBRecordset.UpdateBatch adAffectCurrent
'
'                RcFilters3.DBRecordset.Fields("goods_group") = "GOODS"
'                RcFilters3.DBRecordset.UpdateBatch adAffectCurrent
'                RcFilters3.DBRecordset.Delete adAffectCurrent
'                RcFilters3.DBRecordset.MoveNext
'           Loop
'        End If
'    Case 2
'        If RcFiltersGoods.DBRecordset.Recordcount <> 0 Then
'            SendDataToServer ("UPDATE AccType SET  goods_group = N'' WHERE (ID = " & RcFiltersGoods.Fields("No ID") & "  ) and (goods_group= N'GOODS')")
'            RcFilters3.DBOpen "SELECT ID AS [No ID], Tipe AS Description, goods_group FROM AccType WHERE (goods_group = N'') ", CNN, lckLockBatch, lckLockSync
'            Set PaymentA.DataSource = RcFilters3.DBRecordset
'
'            RcFiltersGoods.DBRecordset.Delete adAffectCurrent
'            RcFiltersGoods.DBRecordset.UpdateBatch adAffectCurrent
'        End If
'    Case 3
'        If RcFiltersGoods.DBRecordset.Recordcount <> 0 Then
'           RcFiltersGoods.DBRecordset.MoveFirst
'          Do
'            If RcFiltersGoods.DBRecordset.EOF Then Exit Do
'            SendDataToServer ("UPDATE AccType SET  goods_group = N'' WHERE (ID = " & RcFiltersGoods.Fields("No ID") & "  ) and (goods_group= N'GOODS')")
'            RcFilters3.DBOpen "SELECT ID AS [No ID], Tipe AS Description, goods_group FROM AccType WHERE (goods_group = N'') ", CNN, lckLockBatch, lckLockSync
'            Set PaymentA.DataSource = RcFilters3.DBRecordset
'            SendDataToServer ("DELETE FROM [Table Filter Account] WHERE ([Group Type] = N'GOODS') AND (ID = " & RcFiltersGoods.Fields("No ID") & ")")
'            RcFiltersGoods.DBRecordset.Delete adAffectCurrent
'            RcFiltersGoods.DBRecordset.MoveNext
'           Loop
'        End If
'End Select
'Exit Sub
'Err.Clear
'End Sub

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
   ComboRelasi.Enabled = True
Else
   CmdSet(0).Enabled = True
   CmdSet(1).Enabled = False
   ComboRelasi.Enabled = False
   SendDataToServer ("  Update [Tabel Pembantu] Set [Seting Relasi] = 0  ")
   SendDataToServer ("  Update [Tabel Pembantu] Set [Seting Relasi] = 1  Where NoAccount=N'" & ComboRelasi.BoundText & "'")
End If
End Sub

Private Sub Combo1_Click()
If Combo1 <> "" Then
   RcSetup.DBOpen "SELECT     [Tabel Pembantu].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun] FROM         [Tabel Pembantu] INNER JOIN                       GLAccount ON [Tabel Pembantu].NoAccount = GLAccount.NoAccount WHERE     ([Tabel Pembantu].[Seting ArusKas] = 1) AND (GLAccount.[Group] = N'Sub Account') AND ([Tabel Pembantu].[Label ArusKas] = N'" & Combo1 & "') ORDER BY [Tabel Pembantu].NoAccount", CNN, lckLockBatch
Else
   RcSetup.DBOpen " SELECT [Tabel Pembantu].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun] FROM [Tabel Pembantu] INNER JOIN GLAccount ON [Tabel Pembantu].NoAccount = GLAccount.NoAccount WHERE     ([Tabel Pembantu].[Seting ArusKas] = 1) AND (GLAccount.[Group] = N'Sub Account') ORDER BY [Tabel Pembantu].NoAccount", CNN, lckLockBatch
End If
Set GrdArusKas(SSTab2.Tab).DataSource = RcSetup.DBRecordset
End Sub

Private Sub CmdPanahRcv1_Click()
If RcFilters.DBRecordset.Recordcount <> 0 Then
    Debug.Print RcFiltersReceipt.DBRecordset.Source
    RcFiltersReceipt.DBRecordset.AddNew 0, RcFilters.Fields(0)
    RcFiltersReceipt.DBRecordset.Fields(1) = RcFilters.Fields(1)
    RcFiltersReceipt.DBRecordset.Fields("Group Type") = "RECEIPT"
    RcFiltersReceipt.DBRecordset.UpdateBatch adAffectCurrent
    
    RcFilters.DBRecordset.Fields("receipt_group") = "RECEIPT"
    RcFilters.DBRecordset.UpdateBatch adAffectCurrent
    RcFilters.DBRecordset.Delete adAffectCurrent
End If
End Sub

Private Sub CmdPanahRcv2_Click()
On Error Resume Next
If RcFilters.DBRecordset.Recordcount <> 0 Then
   RcFilters.DBRecordset.MoveFirst
   Do
     If RcFilters.DBRecordset.EOF Then Exit Do
        RcFiltersReceipt.DBRecordset.AddNew 0, RcFilters.Fields(0)
        RcFiltersReceipt.DBRecordset.Fields(1) = RcFilters.Fields(1)
        RcFiltersReceipt.DBRecordset.Fields("Group Type") = "RECEIPT"
        RcFiltersReceipt.DBRecordset.UpdateBatch adAffectCurrent
        RcFilters.DBRecordset.Fields("receipt_group") = "RECEIPT"
        RcFilters.DBRecordset.UpdateBatch adAffectCurrent
        RcFilters.DBRecordset.Delete adAffectCurrent
        RcFilters.DBRecordset.MoveNext
   Loop
End If
Err.Clear
End Sub

Private Sub CmdPanahRcv3_Click()
On Error Resume Next
If RcFiltersReceipt.DBRecordset.Recordcount <> 0 Then
    SendDataToServer ("UPDATE AccType SET  receipt_group = N'' WHERE (ID = " & RcFiltersReceipt.Fields("No ID") & "  ) and (receipt_group= N'RECEIPT')")
    RcFilters.DBOpen "SELECT ID AS [No ID], Tipe AS Description, receipt_group FROM AccType WHERE (receipt_group = N'') ", CNN, lckLockBatch, lckLockSync
    Set DataGrid1.DataSource = RcFilters.DBRecordset
    
    RcFiltersReceipt.DBRecordset.Delete adAffectCurrent
    RcFiltersReceipt.DBRecordset.UpdateBatch adAffectCurrent
End If
Err.Clear
End Sub

Private Sub CmdPanahRcv4_Click()
On Error Resume Next
If RcFiltersReceipt.DBRecordset.Recordcount <> 0 Then
   RcFiltersReceipt.DBRecordset.MoveFirst
  Do
    If RcFiltersReceipt.DBRecordset.EOF Then Exit Do
    SendDataToServer ("UPDATE AccType SET  receipt_group = N'' WHERE (ID = " & RcFiltersReceipt.Fields("No ID") & "  ) and (receipt_group= N'RECEIPT')")
    RcFilters.DBOpen "SELECT ID AS [No ID], Tipe AS Description, receipt_group FROM AccType WHERE (receipt_group = N'') ", CNN, lckLockBatch, lckLockSync
    Set DataGrid1.DataSource = RcFilters.DBRecordset
    
    SendDataToServer ("DELETE FROM [Table Filter Account] WHERE ([Group Type] = N'RECEIPT') AND (ID = " & RcFiltersReceipt.Fields("No ID") & ")")
    RcFiltersReceipt.DBRecordset.Delete adAffectCurrent
    'RcFiltersReceipt.DBRecordset.UpdateBatch adAffectCurrent
    RcFiltersReceipt.DBRecordset.MoveNext
   Loop
End If
Err.Clear
End Sub


Private Sub CmdPanahPay_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        If RcFilters2.DBRecordset.Recordcount <> 0 Then
            RcFiltersPayment.DBRecordset.AddNew 0, RcFilters2.Fields(0)
            RcFiltersPayment.DBRecordset.Fields(1) = RcFilters2.Fields(1)
            RcFiltersPayment.DBRecordset.Fields("Group Type") = "PAYMENT"
            RcFiltersPayment.DBRecordset.UpdateBatch adAffectCurrent
            
            RcFilters2.DBRecordset.Fields("payment_group") = "PAYMENT"
            RcFilters2.DBRecordset.UpdateBatch adAffectCurrent
            RcFilters2.DBRecordset.Delete adAffectCurrent
        End If
    Case 1
        If RcFilters2.DBRecordset.Recordcount <> 0 Then
           RcFilters2.DBRecordset.MoveFirst
           Do
             If RcFilters2.DBRecordset.EOF Then Exit Do
                RcFiltersPayment.DBRecordset.AddNew 0, RcFilters2.Fields(0)
                RcFiltersPayment.DBRecordset.Fields(1) = RcFilters2.Fields(1)
                RcFiltersPayment.DBRecordset.Fields("Group Type") = "PAYMENT"
                RcFiltersPayment.DBRecordset.UpdateBatch adAffectCurrent
                
                RcFilters2.DBRecordset.Fields("payment_group") = "PAYMENT"
                RcFilters2.DBRecordset.UpdateBatch adAffectCurrent
                RcFilters2.DBRecordset.Delete adAffectCurrent
                RcFilters2.DBRecordset.MoveNext
           Loop
        End If
    Case 2
        If RcFiltersPayment.DBRecordset.Recordcount <> 0 Then
            SendDataToServer ("UPDATE AccType SET  payment_group = N'' WHERE (ID = " & RcFiltersPayment.Fields("No ID") & "  ) and (payment_group= N'PAYMENT')")
            RcFilters2.DBOpen "SELECT ID AS [No ID], Tipe AS Description, payment_group FROM AccType WHERE (payment_group = N'') ", CNN, lckLockBatch, lckLockSync
            Set PaymentA.DataSource = RcFilters2.DBRecordset
            
            RcFiltersPayment.DBRecordset.Delete adAffectCurrent
            RcFiltersPayment.DBRecordset.UpdateBatch adAffectCurrent
        End If
    Case 3
        If RcFiltersPayment.DBRecordset.Recordcount <> 0 Then
           RcFiltersPayment.DBRecordset.MoveFirst
          Do
            If RcFiltersPayment.DBRecordset.EOF Then Exit Do
            SendDataToServer ("UPDATE AccType SET  payment_group = N'' WHERE (ID = " & RcFiltersPayment.Fields("No ID") & "  ) and (payment_group= N'PAYMENT')")
            RcFilters2.DBOpen "SELECT ID AS [No ID], Tipe AS Description, payment_group FROM AccType WHERE (payment_group = N'') ", CNN, lckLockBatch, lckLockSync
            Set PaymentA.DataSource = RcFilters2.DBRecordset
            SendDataToServer ("DELETE FROM [Table Filter Account] WHERE ([Group Type] = N'PAYMENT') AND (ID = " & RcFiltersPayment.Fields("No ID") & ")")
            RcFiltersPayment.DBRecordset.Delete adAffectCurrent
            RcFiltersPayment.DBRecordset.MoveNext
           Loop
        End If
End Select
Exit Sub
Err.Clear
End Sub
Private Sub ComboRelasi_Click(Area As Integer)
lblAkun = ComboRelasi.BoundText
End Sub

Private Sub ComboPosisi_Change()
OpenPOS Left(ComboPosisi.BoundText, 1)
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
If DgAccount.col = 0 And CmdKonfig(1).Enabled = True Then
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
         SendDataToServer " UPDATE GLAccount" & _
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
If CmdControl(0).Enabled = False And DgControl.col >= 1 Then
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
      Dim GroupCode As String
      GroupCode = GetCode(RcPosition.DBRecordset.Fields(0))
      
      RcPosition.DBRecordset.Fields(ColIndex) = CBool(DgPosition.Columns(ColIndex).Value)
'      SendDataToServer (" UPDATE    GLAccount " & _
'                        " Set [Default] = " & BoolToInt(CBool(RcPosition.DBRecordset.Fields(ColIndex))) & _
'                        " WHERE     (LEFT(NoAccount, 4) = N'" & Left(RcPosition.DBRecordset.Fields(0), 4) & "') ")

      SendDataToServer " UPDATE    GLAccount " & _
                        " Set [Default] = " & BoolToInt(CBool(RcPosition.DBRecordset.Fields(ColIndex))) & _
                        " WHERE     NoAccount like '" & GroupCode & "%'"

   End If
End If
End Sub


Function GetCode(inputCode As String) As String
   Dim ln As Integer
   Dim x As Integer
   Dim res As String
   ln = Len(Trim(inputCode))
   res = ""
   For x = 1 To ln
      If Mid(inputCode, x, 1) = "0" Then
         Exit For
      Else
         res = res & Mid(inputCode, x, 1)
      End If
   Next
   GetCode = res
End Function


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
If CmdLevel(0).Enabled = False And DgPrefix.col > 0 Then
   If IsConfigReady = False And DgPrefix.col > 1 Then
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
      SendDataToServer (" Update [Tabel Pembantu]  Set [Kelompok Perkiraan] = " & iVar & " Where left(NoAccount,1)=N'" & Left(RcSeting.DBRecordset.Fields(0), 1) & "'")
   Else
      SendDataToServer (" Update [Tabel Pembantu]  Set [Kelompok Modal] = " & iVar & " Where left(NoAccount,1)=N'" & Left(RcSeting.DBRecordset.Fields(0), 1) & "'")
   End If
   
End If
End Sub

Private Sub DgSeting_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DgSeting.col = 2 Or DgSeting.col = 3 Then
   DgSeting.MarqueeStyle = dbgFloatingEditor
Else
   DgSeting.MarqueeStyle = dbgHighlightRow
End If
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
Set FrmSetupAccount = Nothing
End Sub

Private Sub OpenSetup()
Select Case SSTab2.Tab
       Case 0:
            RcSetup.DBOpen " SELECT [Tabel Pembantu].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun] FROM [Tabel Pembantu] INNER JOIN GLAccount ON [Tabel Pembantu].NoAccount = GLAccount.NoAccount WHERE     ([Tabel Pembantu].[Seting ArusKas] = 1) AND (GLAccount.[Group] = N'Sub Account') ORDER BY [Tabel Pembantu].NoAccount", CNN, lckLockBatch
            Combo1.ListIndex = 0
            Combo1_Click
       Case 1:
            RcSetup.DBOpen " SELECT     [Tabel Pembantu].NoAccount AS [No Akun], GLAccount.AccountName AS [Nama Akun] FROM         [Tabel Pembantu] INNER JOIN                       GLAccount ON [Tabel Pembantu].NoAccount = GLAccount.NoAccount WHERE     (GLAccount.[Group] = N'Sub Account') AND ([Tabel Pembantu].[Kelompok Modal] = 1) ORDER BY [Tabel Pembantu].NoAccount", CNN, lckLockBatch
End Select
RcAccount.DBOpen "SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun] FROM         GLAccount WHERE     ([Group] = N'List Account')", CNN, lckLockBatch
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
   If GrdArusKas(SSTab2.Tab).col = 0 Then
      GrdArusKas(SSTab2.Tab).MarqueeStyle = dbgFloatingEditor
   Else
      GrdArusKas(SSTab2.Tab).MarqueeStyle = dbgHighlightRow
   End If
End If
End Sub

Private Sub LView_ItemClick(ByVal Item As MSComctlLib.ListItem)
   LoadRelasi Val(Item.Text)
'   MsgBox LView.SelectedItem.Text
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
OpenSetup
Combo1.ListIndex = 0
Combo1_Click
OpenControl
OpenPrefix
OpenRelasi
SSTab3.Tab = 0
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
OpenSetup
End Sub

Private Sub OpenControl(Optional pBy As String = "", Optional sValue As String = "")
Dim x As Integer
On Error GoTo xErr
If (pBy = "") And (sValue = "") Then
   RcControl.DBOpen "SELECT ID AS [Job No], Tipe AS Description, status FROM AccType", CNN, lckLockBatch
Else
   If pBy = "Tipe" Then
      RcControl.DBOpen "SELECT ID AS [Job No], Tipe AS Description, status FROM AccType where tipe like '%" & sValue & "%'", CNN, lckLockBatch
   Else
      RcControl.DBOpen "SELECT ID AS [Job No], Tipe AS Description, status FROM AccType where ID=" & sValue, CNN, lckLockBatch
   End If
End If
Set DgControl.DataSource = RcControl.DBRecordset
Exit Sub
xErr:
   Err.Clear
End Sub

Private Sub OpenPOS(ByVal Param As String)
RcPosition.DBOpen "SELECT NoAccount AS Code, AccountName AS [Account Name],  [Default] AS [Position], [Group] FROM GLAccount " & _
        " WHERE (LEFT(NoAccount, 1) = N'" & Param & "')  and ([Group] = N'Sub Account') ORDER BY NoAccount", CNN, lckLockBatch
'Debug.Print RcPosition.DBRecordset.Source
Set DgPosition.DataSource = RcPosition.DBRecordset
End Sub

Private Sub GridLayout()
DgControl.Columns(0).width = 1695.118
DgControl.Columns(1).width = 7000.953
DgSeting.Columns(0).width = 1964.976
DgSeting.Columns(1).width = 6059.906
DgSeting.Columns(2).width = 2324.977
DgAccount.Columns(0).width = 1695.118
DgAccount.Columns(1).width = 3075.024
GrdArusKas(0).Columns(0).width = 1665.071
GrdArusKas(0).Columns(1).width = 2880

GrdArusKas(1).Columns(0).width = 1665.071
GrdArusKas(1).Columns(1).width = 2880

With DgPrefix
    .Columns(0).width = 800
    .Columns(1).width = 3000
    .Columns(2).width = 3000
    .Columns(3).width = 1000
    .Columns(4).width = 800
    .Columns(0).Alignment = dbgCenter
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Alignment = dbgCenter
End With

End Sub

Private Sub OpenPrefix()
RcPrefix.DBOpen "SELECT [No Index] AS No, [Index Group] AS Keterangan, " & _
        " [Group Name] AS [Level], [Length Per Account] AS Panjang, Prefix " & _
        " From dbo.[Account Setup] ORDER BY No", CNN, lckLockBatch
        
'Debug.Print RcPrefix.DBRecordset.Source
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
           LinkAccount = LinkAccount & Mid(Str, j, Val(Avdata(3, I))) & Avdata(4, I)
           j = j + Val(Avdata(2, I))
           SendDataToServer (" UPDATE [Account Setup]" & _
                             " Set [Group Name] = N'" & Avdata(2, I) & "', [Length Per Account] = " & Avdata(3, I) & ", Prefix = N'" & Avdata(4, I) & "'" & _
                             " WHERE ([No Index] = " & Avdata(0, I) & ")")
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





