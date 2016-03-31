VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPExtration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4725
      Left            =   150
      ScaleHeight     =   4695
      ScaleWidth      =   10290
      TabIndex        =   0
      Top             =   495
      Width           =   10320
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   30
         ScaleHeight     =   4545
         ScaleWidth      =   10185
         TabIndex        =   1
         Top             =   30
         Width           =   10215
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Tangki"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   2
            Left            =   2250
            TabIndex        =   6
            Tag             =   "alkali"
            Top             =   1215
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "No_ekstrasi"
            DataSource      =   "DDE"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   0
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   5
            Tag             =   "alkali"
            Top             =   315
            Width           =   2160
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Grup"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   1
            Left            =   2250
            TabIndex        =   4
            Tag             =   "alkali"
            Top             =   915
            Width           =   1695
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2640
            Left            =   120
            TabIndex        =   2
            Tag             =   "minta_sampel"
            Top             =   1800
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4657
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "ItemName"
               Caption         =   "NAMA PRODUK"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "M/d/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Jumlah"
               Caption         =   "JUMLAH"
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
               DataField       =   "tanggal_butuh"
               Caption         =   "DI BUTUHKAN TANGGAL"
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
               DataField       =   "Name"
               Caption         =   "DI TUJUKAN KE"
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
               DataField       =   "keterangan"
               Caption         =   "KETERANGAN"
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
                  Button          =   -1  'True
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column03 
                  Button          =   -1  'True
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column04 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "AT_tanggal"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   0
            Left            =   2250
            TabIndex        =   7
            Tag             =   "alkali"
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   20643843
            CurrentDate     =   39365
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   2355
            X2              =   315
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   2355
            X2              =   315
            Y1              =   915
            Y2              =   915
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   2370
            X2              =   330
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Tangki"
            Height          =   255
            Index           =   12
            Left            =   315
            TabIndex        =   11
            Top             =   1290
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
            Height          =   255
            Index           =   2
            Left            =   315
            TabIndex        =   10
            Top             =   990
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   1
            Left            =   315
            TabIndex        =   9
            Top             =   675
            Width           =   2055
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "No Ekstrasi"
            Height          =   255
            Index           =   0
            Left            =   315
            TabIndex        =   8
            Top             =   390
            Width           =   2055
         End
      End
   End
   Begin SemeruDC.SemeruOleDC mydc 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Tag             =   "minta_sampel"
      Top             =   6405
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1005
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPExtration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
