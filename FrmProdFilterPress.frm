VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmProdFilterPress 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Press"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProdFilterPress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11715
   Begin VB.TextBox lblNoEkstraksi 
      Appearance      =   0  'Flat
      DataField       =   "noekstrasi"
      DataSource      =   "MyDDE"
      Height          =   315
      Left            =   2145
      TabIndex        =   0
      Tag             =   "FIL"
      Top             =   180
      Width           =   2055
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   6225
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1005
      BindFormTAG     =   "FIL"
      InitControlSet  =   1
      ActiveLanguage  =   1
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
      Height          =   6225
      Left            =   0
      ScaleHeight     =   6225
      ScaleWidth      =   11715
      TabIndex        =   5
      Top             =   0
      Width           =   11715
      Begin TabDlg.SSTab SSTab1 
         Height          =   3270
         Left            =   150
         TabIndex        =   23
         Top             =   2040
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   5768
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Detail"
         TabPicture(0)   =   "FrmProdFilterPress.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "gridDetail(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "PreCoating"
         TabPicture(1)   =   "FrmProdFilterPress.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gridDetail(1)"
         Tab(1).Control(1)=   "DTPrecoating"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Filter Press 1"
         TabPicture(2)   =   "FrmProdFilterPress.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "gridDetail(2)"
         Tab(2).Control(1)=   "DTPFilter1"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Filter Press 2"
         TabPicture(3)   =   "FrmProdFilterPress.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "gridDetail(3)"
         Tab(3).Control(1)=   "DTPFilter2"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Filter Press 3"
         TabPicture(4)   =   "FrmProdFilterPress.frx":68C2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "gridDetail(4)"
         Tab(4).Control(1)=   "DTPFilter3"
         Tab(4).ControlCount=   2
         Begin MSComCtl2.DTPicker DTPFilter3 
            Height          =   315
            Left            =   -72555
            TabIndex        =   33
            Top             =   1470
            Visible         =   0   'False
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   56098818
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker DTPFilter2 
            Height          =   315
            Left            =   -73050
            TabIndex        =   32
            Top             =   1575
            Visible         =   0   'False
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   56098818
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker DTPFilter1 
            Height          =   300
            Left            =   -72930
            TabIndex        =   31
            Top             =   1545
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   56098818
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker DTPrecoating 
            Height          =   330
            Left            =   -71760
            TabIndex        =   30
            Top             =   1245
            Visible         =   0   'False
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   56098818
            CurrentDate     =   36494
         End
         Begin MSDataGridLib.DataGrid gridDetail 
            Height          =   2805
            Index           =   0
            Left            =   75
            TabIndex        =   24
            Top             =   390
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   4948
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "waktu"
               Caption         =   "Waktu (Menit)"
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
               DataField       =   "suhu1"
               Caption         =   "Suhu Fp1"
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
               DataField       =   "suhu2"
               Caption         =   "Suhu Fp2"
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
            BeginProperty Column03 
               DataField       =   "suhu3"
               Caption         =   "Suhu Fp3"
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
            BeginProperty Column04 
               DataField       =   "ph1"
               Caption         =   "pH Fp1"
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
            BeginProperty Column05 
               DataField       =   "ph2"
               Caption         =   "pH Fp2"
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
            BeginProperty Column06 
               DataField       =   "ph3"
               Caption         =   "pH Fp3"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Alignment       =   2
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
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
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gridDetail 
            Height          =   2805
            Index           =   1
            Left            =   -74925
            TabIndex        =   25
            Top             =   390
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   4948
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "nama_precoating"
               Caption         =   "Precoating"
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
               DataField       =   "waktu_mulai_mixer"
               Caption         =   "Waktu Mulai Mixer"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "waktu_selesai_mixer"
               Caption         =   "Waktu Selesai Mixer"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "jml_air"
               Caption         =   "jml Air (Liter)"
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
            BeginProperty Column04 
               DataField       =   "type_filter_aid"
               Caption         =   "Type Filter Aid"
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
               DataField       =   "qty_filter_aid"
               Caption         =   "Qty Filter Aid"
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
            BeginProperty Column06 
               DataField       =   "suhu_larutan"
               Caption         =   "Suhu Larutan"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gridDetail 
            Height          =   2805
            Index           =   2
            Left            =   -74925
            TabIndex        =   26
            Top             =   390
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   4948
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
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "no"
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
               DataField       =   "waktu_mulai_precoating"
               Caption         =   "Waktu Mulai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "waktu_selesai_precoating"
               Caption         =   "Waktu Selesai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "waktu_mulai_pompa"
               Caption         =   "Waktu Mulai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "waktu_selesai_pompa"
               Caption         =   "Waktu Selesai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "waktu_mulai_bongkar"
               Caption         =   "Waktu Mulai Bongkar"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "waktu_selesai_pompa"
               Caption         =   "Waktu Selesai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "kondisi_solid_waste"
               Caption         =   "Kondisi Solid Waste"
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
            BeginProperty Column08 
               DataField       =   "kondisi_kain"
               Caption         =   "Kondisi Kain"
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
                  Alignment       =   2
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   1440
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
               EndProperty
               BeginProperty Column08 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gridDetail 
            Height          =   2805
            Index           =   3
            Left            =   -74925
            TabIndex        =   27
            Top             =   390
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   4948
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
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "no"
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
               DataField       =   "waktu_mulai_precoating"
               Caption         =   "Waktu Mulai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "waktu_selesai_precoating"
               Caption         =   "Waktu Selesai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "waktu_mulai_pompa"
               Caption         =   "Waktu Mulai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "waktu_selesai_pompa"
               Caption         =   "Waktu Selesai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "waktu_mulai_bongkar"
               Caption         =   "Waktu Mulai Bongkar"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "waktu_selesai_pompa"
               Caption         =   "Waktu Selesai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "kondisi_solid_waste"
               Caption         =   "Kondisi Solid Waste"
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
            BeginProperty Column08 
               DataField       =   "kondisi_kain"
               Caption         =   "Kondisi Kain"
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
                  Alignment       =   2
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   1440
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
               EndProperty
               BeginProperty Column08 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gridDetail 
            Height          =   2805
            Index           =   4
            Left            =   -74925
            TabIndex        =   28
            Top             =   390
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   4948
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "no"
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
               DataField       =   "waktu_mulai_precoating"
               Caption         =   "Waktu Mulai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "waktu_selesai_precoating"
               Caption         =   "Waktu Selesai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "waktu_mulai_pompa"
               Caption         =   "Waktu Mulai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "waktu_selesai_pompa"
               Caption         =   "Waktu Selesai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "waktu_mulai_bongkar"
               Caption         =   "Waktu Mulai Bongkar"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "waktu_selesai_pompa"
               Caption         =   "Waktu Selesai Pompa"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "kondisi_solid_waste"
               Caption         =   "Kondisi Solid Waste"
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
            BeginProperty Column08 
               DataField       =   "kondisi_kain"
               Caption         =   "Kondisi Kain"
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
                  Alignment       =   2
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   1440
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
               EndProperty
               BeginProperty Column08 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Filter Aid Untuk DAT2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   6165
         TabIndex        =   14
         Top             =   105
         Width           =   5415
         Begin VB.CommandButton cmdLink 
            Height          =   315
            Left            =   4710
            Picture         =   "FrmProdFilterPress.frx":68DE
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   315
            Width           =   390
         End
         Begin VB.TextBox txtKuantitas 
            Appearance      =   0  'Flat
            DataField       =   "qty"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   2595
            TabIndex        =   17
            Tag             =   "FIL"
            Top             =   675
            Width           =   2505
         End
         Begin VB.TextBox txtTypeFilter 
            Appearance      =   0  'Flat
            DataField       =   "Type_Filter_Aid"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   16
            Tag             =   "FIL"
            Top             =   315
            Width           =   2100
         End
         Begin VB.TextBox txtSuhuDat 
            Appearance      =   0  'Flat
            DataField       =   "Suhu_di_DAT"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   2610
            TabIndex        =   15
            Tag             =   "FIL"
            Top             =   1365
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "waktu_memasukkan"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   2610
            TabIndex        =   22
            Tag             =   "FIL"
            Top             =   1020
            Width           =   2490
            _ExtentX        =   4392
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
            CustomFormat    =   "dd MMM yyyy  HH:mm"
            Format          =   56098819
            CurrentDate     =   39634
         End
         Begin VB.Label lblKondisi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   3840
            TabIndex        =   29
            Top             =   1425
            Width           =   105
         End
         Begin VB.Label lblReaktor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kuantitas"
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
            Left            =   150
            TabIndex        =   21
            Top             =   735
            Width           =   675
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   12
            X1              =   2745
            X2              =   150
            Y1              =   975
            Y2              =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   11
            X1              =   2655
            X2              =   180
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Label lblKondisi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suhu di DAT"
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
            Left            =   150
            TabIndex        =   20
            Top             =   1425
            Width           =   870
         End
         Begin VB.Label lblTempatAlkali 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Memasukkan Filter Aid"
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
            Left            =   150
            TabIndex        =   19
            Top             =   1065
            Width           =   2100
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   9
            X1              =   2735
            X2              =   165
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            X1              =   2805
            X2              =   150
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Label lblFilterAid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type Filter Aid"
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
            Left            =   150
            TabIndex        =   18
            Top             =   390
            Width           =   1035
         End
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         DataField       =   "grup"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   2145
         TabIndex        =   3
         Tag             =   "FIL"
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         DataField       =   "desk_filter"
         DataSource      =   "MyDDE"
         Height          =   810
         Left            =   1770
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "FIL"
         Top             =   5370
         Width           =   6750
      End
      Begin MSComCtl2.DTPicker DcTanggal 
         DataField       =   "tanggal_press"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   2130
         TabIndex        =   2
         Tag             =   "FIL"
         Top             =   525
         Width           =   2070
         _ExtentX        =   3651
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
         Format          =   56098819
         CurrentDate     =   39634
      End
      Begin VB.Label Label4 
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
         Left            =   9765
         TabIndex        =   37
         Top             =   5835
         Width           =   1830
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   11295
         X2              =   8670
         Y1              =   6150
         Y2              =   6150
      End
      Begin VB.Label Label3 
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
         Left            =   8700
         TabIndex        =   36
         Top             =   5895
         Width           =   930
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OrderID"
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
         TabIndex        =   13
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label LbRefID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "orderID"
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
         Left            =   2145
         TabIndex        =   12
         Tag             =   "FIL"
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   2775
         X2              =   150
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   2415
         X2              =   120
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lblFilterAid 
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   2190
         X2              =   120
         Y1              =   825
         Y2              =   825
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
         TabIndex        =   8
         Top             =   570
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   2250
         X2              =   120
         Y1              =   1200
         Y2              =   1200
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
         TabIndex        =   7
         Top             =   945
         Width           =   435
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
         Left            =   255
         TabIndex        =   6
         Top             =   5955
         Width           =   840
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   2655
      X2              =   30
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "orderID"
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
      Left            =   2025
      TabIndex        =   35
      Tag             =   "FIL"
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OrderID"
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
      Left            =   0
      TabIndex        =   34
      Top             =   60
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
      TabIndex        =   10
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "FrmProdFilterPress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strSQL As String
Dim GridAltColor As String
Dim Changingsel As Byte
Dim Xval As String
Dim aCol As Integer
 
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1

Dim MEdit As Boolean
Dim sWCID As String

Private RsDetail As New DBQuick
Private rsPrecoating As New DBQuick
Private rsFilterPress1 As New DBQuick
Private rsFilterPress2 As New DBQuick
Private rsFilterPress3 As New DBQuick
Private rcPartner As New DBQuick




Private Sub cmdEkstraksi_Click()
    OPenPartner 0
End Sub

Private Sub cmdRefLink_Click()
    OPenPartner 1
End Sub

Private Sub cmdLink_Click()
   OPenPartner 0
End Sub

Private Sub DTPFilter1_Change()
   gridDetail(2).Columns(gridDetail(2).col).Value = DTPFilter1.Value
End Sub

Private Sub DTPFilter2_Change()
   gridDetail(3).Columns(gridDetail(3).col).Value = DTPFilter2.Value
End Sub

Private Sub DTPFilter3_Change()
   gridDetail(4).Columns(gridDetail(4).col).Value = DTPFilter3.Value
End Sub

Private Sub DTPrecoating_Change()
  
   'aCol = gridDetail(1).col
   gridDetail(1).Columns(aCol).Value = DTPrecoating.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    ScanKey KeyCode, Shift, MyDDE

    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
 '   lblNoEkstraksi = frmProduksi.txtBox(5)
 '   LbRefID = frmProduksi.lblSplNo
    HiasFormManTell Picture2, Me

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        .BindFormTAG = "FIL"
        .SetPermissions = UserDeleteDenied
        Set .ActiveConnection = CNN

        .PrepareQuery = "SELECT * From filter_Header where status=0"
        .SetPermissions = aksess.MayDo("Filter Press")
    End With
    
    gridDetail(0).HeadLines = 2
    gridDetail(1).HeadLines = 2
    gridDetail(2).HeadLines = 3
    gridDetail(3).HeadLines = 3
    gridDetail(4).HeadLines = 3
    
    gridDetail(1).RowHeight = 300
    gridDetail(2).RowHeight = 300
    gridDetail(3).RowHeight = 300
    gridDetail(4).RowHeight = 300
    
    Set mCall = New frmCaller
    
End Sub

Private Sub PrepareQuery()
    On Error GoTo xErr
    Dim ket As Byte

    With MyDDE
      .PrepareAppend = "insert into filter_header(id_filter,no_ekstraksi,tanggal_press,grup,desk_filter,orderID,type_filter_aid,qty,waktu_memasukkan,suhu_di_dat,status,issued_by) values ('" & _
                           lblNoEkstraksi.Text & "','" & lblNoEkstraksi & "','" & Format(DcTanggal.Value, "yyyy-MM-dd") & "','" & txtGroup & "','" & txtKeterangan.Text & "','" & _
                           LbRefID.Caption & "','" & txtTypeFilter & "'," & FQty(txtKuantitas) & ",'" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & "'," & FQty(txtSuhuDat) & ",0,'" & MainMenu.StatusBar1.Panels(1).Text & "')"
                           
      .PrepareUpdate = "update filter_header set tanggal_press = '" & Format(DcTanggal.Value, "yyyy-MM-dd") & "',grup='" & txtGroup & "',desk_filter='" & txtKeterangan.Text & _
                           "',orderID='" & LbRefID.Caption & "',type_filter_aid='" & txtTypeFilter & "',qty =" & FQty(txtKuantitas) & ",waktu_memasukkan='" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & _
                           "',suhu_di_dat =" & FQty(txtSuhuDat) & " where id_filter='" & lblNoEkstraksi & "'"
                           
      .PrepareDelete = "DELETE From filter_header Where id_filter = '" & lblNoEkstraksi & "'"
    End With

Exit Sub
xErr:
   Err.Clear
End Sub

Private Sub SimpanDetail()
   SendDataToServer "delete from filter_detail where id_filter='" & lblNoEkstraksi & "'"
   With RsDetail.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "insert into filter_detail (id_filter,waktu,suhu1,suhu2,suhu3,ph1,ph2,ph3) values ('" & _
                           lblNoEkstraksi & "'," & .Fields("waktu") & "," & .Fields("suhu1") & "," & _
                           .Fields("suhu2") & "," & .Fields("suhu3") & "," & .Fields("ph1") & "," & .Fields("ph2") & "," & .Fields("ph3") & ")"
         .MoveNext
      Wend
   End With
   
   SendDataToServer "delete from filter_precoating where id_filter='" & lblNoEkstraksi & "'"
   With rsPrecoating.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO [filter_precoating] ([id_filter],[nama_precoating],[waktu_mulai_mixer],[waktu_selesai_mixer],[jml_air],[type_filter_aid],[suhu_larutan],[qty_filter_aid]) Values ('" & _
                         lblNoEkstraksi & "','" & .Fields("nama_precoating") & "','" & Format(.Fields("waktu_mulai_mixer"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                         Format(.Fields("waktu_selesai_mixer"), "yyyy-MM-dd hh:mm:ss") & "'," & .Fields("jml_air") & ",'" & .Fields("type_filter_aid") & "'," & .Fields("suhu_larutan") & "," & .Fields("qty_filter_aid") & ")"
         .MoveNext
      Wend
   End With
   

   SendDataToServer "delete from filtrasi where id_filter ='" & lblNoEkstraksi & "' and tahapan='Filter Press 1'"
   With rsFilterPress1.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO filtrasi ([id_filter],[tahapan],[waktu_mulai_precoating],[waktu_selesai_precoating],[waktu_mulai_pompa],[waktu_selesai_pompa],[waktu_mulai_bongkar],[waktu_selesai_bongkar],[kondisi_solid_waste],[kondisi_kain],[no]) Values ('" & _
                              lblNoEkstraksi & "','Filter Press 1','" & Format(.Fields("waktu_mulai_precoating"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_precoating"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_mulai_pompa"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_pompa"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_mulai_bongkar"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_bongkar"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              .Fields("kondisi_solid_waste") & "','" & .Fields("kondisi_kain") & "'," & .Fields("no") & ")"
         .MoveNext
      Wend
   End With


   SendDataToServer "delete from filtrasi where id_filter ='" & lblNoEkstraksi & "' and tahapan='Filter Press 2'"
   With rsFilterPress2.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO filtrasi ([id_filter],[tahapan],[waktu_mulai_precoating],[waktu_selesai_precoating],[waktu_mulai_pompa],[waktu_selesai_pompa],[waktu_mulai_bongkar],[waktu_selesai_bongkar],[kondisi_solid_waste],[kondisi_kain],[no]) Values ('" & _
                              lblNoEkstraksi & "','Filter Press 2','" & Format(.Fields("waktu_mulai_precoating"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_precoating"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_mulai_pompa"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_pompa"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_mulai_bongkar"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_bongkar"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              .Fields("kondisi_solid_waste") & "','" & .Fields("kondisi_kain") & "'," & .Fields("no") & ")"
         .MoveNext
      Wend
   End With


   SendDataToServer "delete from filtrasi where id_filter ='" & lblNoEkstraksi & "' and tahapan='Filter Press 3'"
   With rsFilterPress3.DBRecordset
      .MoveFirst
      While Not .EOF
         SendDataToServer "INSERT INTO filtrasi ([id_filter],[tahapan],[waktu_mulai_precoating],[waktu_selesai_precoating],[waktu_mulai_pompa],[waktu_selesai_pompa],[waktu_mulai_bongkar],[waktu_selesai_bongkar],[kondisi_solid_waste],[kondisi_kain],[no]) Values ('" & _
                              lblNoEkstraksi & "','Filter Press 3','" & Format(.Fields("waktu_mulai_precoating"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_precoating"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_mulai_pompa"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_pompa"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_mulai_bongkar"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              Format(.Fields("waktu_selesai_bongkar"), "yyyy-MM-dd hh:mm:ss") & "','" & _
                              .Fields("kondisi_solid_waste") & "','" & .Fields("kondisi_kain") & "'," & .Fields("no") & ")"
         .MoveNext
      Wend
   End With

End Sub


Private Function GetWC(ByVal FormIDNya As String)
    On Error GoTo Masjid
    Dim RcGetWC As New DBQuick
    RcGetWC.DBOpen "SELECT wcenter_header.WCID From wcenter_header Where wcenter_header.formid = 38", CNN, lckLockReadOnly
    sWCID = RcGetWC.DBRecordset.Fields("WCID")
    Exit Function
Masjid:
    MessageBox "Konfigurasi Filter Press Kosong"
    Err.Clear
End Function

Private Function OPenPartner(ByVal index As Integer) As Boolean
On Error GoTo Hell:
Select Case index
       Case 0:
            rcPartner.DBOpen " SELECT noItem,InternalName as [Nama Bahan], UOM as Satuan from inventory where left(internalName,2)='FA'", CNN, lckLockReadOnly
            
End Select
If rcPartner.Recordcount <> 0 Then
   Select Case index
          Case 0:
            mCall.FromTagActive = "Bahan"
   End Select
   Set mCall.FormData = rcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Bahan Tidak Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Function
Hell:
    Err.Clear
End Function

Private Sub GridFilterPress_KeyDown(KeyCode As Integer, _
                                    Shift As Integer)

    If MEdit = False Then Exit Sub
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridFilterPress_KeyPress(KeyAscii As Integer)

End Sub


Private Sub GridFilterPress_RowColChange()
    Xval = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mCall = Nothing
End Sub

Private Sub gridDetail_RowColChange(index As Integer, LastRow As Variant, ByVal LastCol As Integer)
   
   Select Case index
      Case 1:
         If (gridDetail(1).col = 1) Or (gridDetail(1).col = 2) Then
            DTPrecoating.Value = gridDetail(1).Columns(gridDetail(1).col).Value
            DTPrecoating.Move gridDetail(1).Left + gridDetail(1).Columns(gridDetail(1).col).Left, _
                              gridDetail(1).Top + gridDetail(1).RowTop(gridDetail(1).row), _
                              gridDetail(1).Columns(gridDetail(1).col).width, _
                              gridDetail(1).RowHeight
            aCol = gridDetail(index).col
            DTPrecoating.Visible = True
         End If
      Case 2:
         If (gridDetail(2).col >= 1) And (gridDetail(2).col <= 6) Then
            DTPFilter1.Value = gridDetail(2).Columns(gridDetail(2).col)
            DTPFilter1.Move gridDetail(2).Left + gridDetail(2).Columns(gridDetail(2).col).Left, _
                              gridDetail(2).Top + gridDetail(2).RowTop(gridDetail(2).row), _
                              gridDetail(2).Columns(gridDetail(2).col).width, _
                              gridDetail(2).RowHeight
            DTPFilter1.Visible = True
         End If
      Case 3:
         If (gridDetail(3).col >= 1) And (gridDetail(3).col <= 6) Then
            DTPFilter2.Value = gridDetail(3).Columns(gridDetail(3).col)
            DTPFilter2.Move gridDetail(3).Left + gridDetail(3).Columns(gridDetail(3).col).Left, _
                              gridDetail(3).Top + gridDetail(3).RowTop(gridDetail(3).row), _
                              gridDetail(3).Columns(gridDetail(3).col).width, _
                              gridDetail(3).RowHeight
            DTPFilter2.Visible = True
         End If
      Case 4:
         If (gridDetail(4).col >= 1) And (gridDetail(4).col <= 6) Then
            DTPFilter3.Value = gridDetail(4).Columns(gridDetail(4).col)
            DTPFilter3.Move gridDetail(4).Left + gridDetail(4).Columns(gridDetail(4).col).Left, _
                              gridDetail(4).Top + gridDetail(4).RowTop(gridDetail(4).row), _
                              gridDetail(4).Columns(gridDetail(4).col).width, _
                              gridDetail(4).RowHeight
            DTPFilter3.Visible = True
         End If
   End Select
End Sub

Private Sub lblNoEkstraksi_LostFocus()
   Dim rsCek As New DBQuick
   rsCek.DBOpen "select * from statusProduksi where noEkstraksi='" & lblNoEkstraksi.Text & "'", CNN, lckLockBatch
   If rsCek.DBRecordset.Recordcount > 0 Then
      rsCek.DBOpen "select * from filter_Header where no_Ekstraksi='" & lblNoEkstraksi.Text & "'", CNN, lckLockBatch
      If rsCek.DBRecordset.Recordcount > 0 Then
         MessageBox "Nomor Ekstraksi Ini Sudah Diinput...!", "Peringatan", msgOkOnly, msgCrtical
         lblNoEkstraksi.Text = ""
      End If
   Else
      MessageBox "Nomor Ekstraksi Ini tidak ditemukan...!", "Peringatan", msgOkOnly, msgCrtical
      lblNoEkstraksi.Text = ""
   End If
   rsCek.CloseDB
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, _
                               ByVal pRecordset As ADODB.Recordset)

    Select Case TagForm
        Case "Bahan"
           txtTypeFilter = mCall.GetFieldByName("nama bahan")
    End Select

End Sub

Private Sub KonfigurasiFilterPress()
    Dim ncount As Integer

    With MyDDE.ActiveRecordset
        Set RsDetail = New DBQuick
        strSQL = "SELECT DISTINCT ProdFormulaEkstraksi.EksNo,Prodconfigproses.ID_ANALYSIS,Prodconfigproses.ProsesID,ProdProses.Prosedur, Prodconfigproses.Analysis,Prodconfigproses.MinValue,Prodconfigproses.Methods," & " ProdAnalysis.unit,ProdFormulaEkstraksi.TypeTrans From ProdFormulaEkstraksi_Detail INNER JOIN ProdFormulaEkstraksi ON (ProdFormulaEkstraksi_Detail.EksNo =  ProdFormulaEkstraksi.EksNo) INNER JOIN Prodconfigproses ON (ProdFormulaEkstraksi_Detail.ProsesID =Prodconfigproses.ProsesID) " & "INNER JOIN ProdProses ON (Prodconfigproses.ProsesID = ProdProses.ProsesID) INNER JOIN ProdAnalysis ON (Prodconfigproses.Analysis = ProdAnalysis.Analysis) where ProdFormulaEkstraksi.typeTrans = 'PRO-FIL'"
        'Debug.Print strSQL
        RsDetail.DBOpen strSQL, CNN

        If RsDetail.Recordcount < 1 Then
            MessageBox "Konfigurasi Filter Press Masih Kosong.", "Peringatan", msgOkOnly
            'cmdSPPH.Enabled = False
        End If

        Set MyDDE.ChildRecordset = RsDetail.DBRecordset.Clone(adLockBatchOptimistic)
      
        

   
    
    End With
  
End Sub


Private Sub OpenDetail(ByVal ParameterString As String)
    
    If ParameterString = "" Then ParameterString = "11111111111" ': Exit Sub
         RsDetail.DBOpen "select * from filter_detail where id_filter='" & ParameterString & "'", CNN, lckLockBatch
         rsPrecoating.DBOpen "select * from filter_precoating where id_filter='" & ParameterString & "'", CNN, lckLockBatch
         rsFilterPress1.DBOpen "select * from filtrasi where id_filter='" & ParameterString & "' and tahapan='Filter Press 1'", CNN, lckLockBatch
         rsFilterPress2.DBOpen "select * from filtrasi where id_filter='" & ParameterString & "' and tahapan='Filter Press 2'", CNN, lckLockBatch
         rsFilterPress3.DBOpen "select * from filtrasi where id_filter='" & ParameterString & "' and tahapan='Filter Press 3'", CNN, lckLockBatch
    
    
    Select Case SSTab1.Tab
      Case 0: Set MyDDE.ChildRecordset = RsDetail.DBRecordset
      Case 1: Set MyDDE.ChildRecordset = rsPrecoating.DBRecordset
      Case 2: Set MyDDE.ChildRecordset = rsFilterPress1.DBRecordset
      Case 3: Set MyDDE.ChildRecordset = rsFilterPress2.DBRecordset
      Case 4: Set MyDDE.ChildRecordset = rsFilterPress3.DBRecordset
    End Select
    
    Set gridDetail(SSTab1.Tab).DataSource = MyDDE.ChildRecordset
    
End Sub



Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    Dim IDGen As New IDGenerator

    Select Case AdReasonActiveDb

        Case tmbSave
            SimpanDetail
            SaveToMO

        Case tmbAddNew
        
            MEdit = True
            Me.Tag = "baru"
            LbRefID.Caption = frmProduksi.txtBox(5)
           ' lblNoEkstraksi = frmProduksi.txtBox(1)
            MyDDE.GetFieldByName("id_filter") = lblNoEkstraksi
            MyDDE.GetFieldByName("desk_filter") = "-"
            DTPicker1.Value = Now
            'KonfigurasiFilterPress
            SetDataDetail

        Case tmbDelete
            PrepareQuery
    End Select

End Sub

Private Sub SetDataDetail()
   '*** detail
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 30
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 60
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 90
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 120
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 150
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 180
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 210
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 240
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 270
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   RsDetail.DBRecordset.AddNew
   RsDetail.DBRecordset.Fields("waktu") = 300
   RsDetail.DBRecordset.Fields("suhu1") = 0
   RsDetail.DBRecordset.Fields("suhu2") = 0
   RsDetail.DBRecordset.Fields("suhu3") = 0
   RsDetail.DBRecordset.Fields("ph1") = 0
   RsDetail.DBRecordset.Fields("ph2") = 0
   RsDetail.DBRecordset.Fields("ph3") = 0
   
   
   '*** Precoating
   rsPrecoating.DBRecordset.AddNew
   rsPrecoating.DBRecordset.Fields("nama_precoating") = "Precoating 1"
   rsPrecoating.DBRecordset.Fields("waktu_mulai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("waktu_selesai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("jml_air") = 0
   rsPrecoating.DBRecordset.Fields("suhu_larutan") = 0
   rsPrecoating.DBRecordset.Fields("qty_filter_aid") = 0

   
   rsPrecoating.DBRecordset.AddNew
   rsPrecoating.DBRecordset.Fields("nama_precoating") = "Precoating 2"
   rsPrecoating.DBRecordset.Fields("waktu_mulai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("waktu_selesai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("jml_air") = 0
   rsPrecoating.DBRecordset.Fields("suhu_larutan") = 0
   rsPrecoating.DBRecordset.Fields("qty_filter_aid") = 0
   
   
   rsPrecoating.DBRecordset.AddNew
   rsPrecoating.DBRecordset.Fields("nama_precoating") = "Precoating 3"
   rsPrecoating.DBRecordset.Fields("waktu_mulai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("waktu_selesai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("jml_air") = 0
   rsPrecoating.DBRecordset.Fields("suhu_larutan") = 0
   rsPrecoating.DBRecordset.Fields("qty_filter_aid") = 0
   
   
   rsPrecoating.DBRecordset.AddNew
   rsPrecoating.DBRecordset.Fields("nama_precoating") = "Precoating 4"
   rsPrecoating.DBRecordset.Fields("waktu_mulai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("waktu_selesai_mixer") = Now
   rsPrecoating.DBRecordset.Fields("jml_air") = 0
   rsPrecoating.DBRecordset.Fields("suhu_larutan") = 0
   rsPrecoating.DBRecordset.Fields("qty_filter_aid") = 0
   
   
   
   '*** filter press 1
   rsFilterPress1.DBRecordset.AddNew
   rsFilterPress1.DBRecordset.Fields("no") = 1
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress1.DBRecordset.AddNew
   rsFilterPress1.DBRecordset.Fields("no") = 2
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress1.DBRecordset.AddNew
   rsFilterPress1.DBRecordset.Fields("no") = 3
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress1.DBRecordset.AddNew
   rsFilterPress1.DBRecordset.Fields("no") = 4
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress1.DBRecordset.AddNew
   rsFilterPress1.DBRecordset.Fields("no") = 5
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress1.DBRecordset.AddNew
   rsFilterPress1.DBRecordset.Fields("no") = 6
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress1.DBRecordset.Fields("waktu_selesai_bongkar") = Now


   '*** filter press 2
   rsFilterPress2.DBRecordset.AddNew
   rsFilterPress2.DBRecordset.Fields("no") = 1
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress2.DBRecordset.AddNew
   rsFilterPress2.DBRecordset.Fields("no") = 2
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress2.DBRecordset.AddNew
   rsFilterPress2.DBRecordset.Fields("no") = 3
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress2.DBRecordset.AddNew
   rsFilterPress2.DBRecordset.Fields("no") = 4
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress2.DBRecordset.AddNew
   rsFilterPress2.DBRecordset.Fields("no") = 5
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress2.DBRecordset.AddNew
   rsFilterPress2.DBRecordset.Fields("no") = 6
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress2.DBRecordset.Fields("waktu_selesai_bongkar") = Now


   '*** filter press 3
   rsFilterPress3.DBRecordset.AddNew
   rsFilterPress3.DBRecordset.Fields("no") = 1
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress3.DBRecordset.AddNew
   rsFilterPress3.DBRecordset.Fields("no") = 2
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress3.DBRecordset.AddNew
   rsFilterPress3.DBRecordset.Fields("no") = 3
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress3.DBRecordset.AddNew
   rsFilterPress3.DBRecordset.Fields("no") = 4
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress3.DBRecordset.AddNew
   rsFilterPress3.DBRecordset.Fields("no") = 5
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_bongkar") = Now

   rsFilterPress3.DBRecordset.AddNew
   rsFilterPress3.DBRecordset.Fields("no") = 6
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_precoating") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_pompa") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_mulai_bongkar") = Now
   rsFilterPress3.DBRecordset.Fields("waktu_selesai_bongkar") = Now


End Sub

Private Sub SaveToMO()
    Dim dStart As Date
    Dim dFinish As Date
    Dim ActualTime As Double
    Dim rsCek As New DBQuick
    Dim sWCID As String
   
    ActualTime = Val(SelisihHariJam(dStart, dFinish, 2))
   
    rsCek.DBOpen "select WCID from WCenter_Header where FormID = 41", CNN

    If rsCek.DBRecordset.Recordcount > 0 Then
        sWCID = rsCek.DBRecordset.Fields(0)
        SendDataToServer "update [order output detail] set actual_time=" & ActualTime & " where OrderID='" & LbRefID.Caption & "' and WCID='" & sWCID & "'"
    End If

    rsCek.CloseDB
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

    Select Case AdReasonActiveDb

        Case tmbSave

            If MyDDE.CheckEmptyControl = False Then
                MyDDE.IsChildMemberReady = True
                SimpanDetail
                SaveToMO
            Else
                MyDDE.IsChildMemberReady = False
            End If

    End Select

End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                               ByVal pError As ADODB.Error, _
                               adStatus As ADODB.EventStatusEnum, _
                               ByVal pRecordset As ADODB.Recordset)


    OpenDetail IIf(IsNull(MyDDE.GetFieldByName("id_filter")), "", MyDDE.GetFieldByName("id_filter"))
    Label4.Caption = IIf(IsNull(MyDDE.GetFieldByName("Approved_by")), "", MyDDE.GetFieldByName("Approved_by"))
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

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Tab
      Case 0: Set MyDDE.ChildRecordset = RsDetail.DBRecordset
      Case 1: Set MyDDE.ChildRecordset = rsPrecoating.DBRecordset
      Case 2: Set MyDDE.ChildRecordset = rsFilterPress1.DBRecordset
      Case 3: Set MyDDE.ChildRecordset = rsFilterPress2.DBRecordset
      Case 4: Set MyDDE.ChildRecordset = rsFilterPress3.DBRecordset
   End Select
   Set gridDetail(SSTab1.Tab).DataSource = MyDDE.ChildRecordset
End Sub

