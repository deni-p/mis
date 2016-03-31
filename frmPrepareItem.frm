VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPrepareItem 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrepareItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   11295
   Begin SemeruDC.SemeruForm SemeruForm1 
      Height          =   6375
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11245
      BackColor       =   16777215
      Begin SemeruDC.SemeruPanels SemeruPanels1 
         Height          =   5715
         Left            =   120
         TabIndex        =   1
         Top             =   345
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   10081
         BackColorTop    =   -2147483646
         BackColor       =   15917781
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Item/Service Dalam Persiapan"
         Begin MSComCtl2.DTPicker DateItem 
            Height          =   315
            Left            =   1710
            TabIndex        =   7
            Top             =   1320
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            Format          =   58064896
            CurrentDate     =   38272
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "CurrID"
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   1
            Left            =   1710
            MaxLength       =   5
            TabIndex        =   6
            Tag             =   "Partner"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "CurrID"
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   0
            Left            =   1710
            MaxLength       =   5
            TabIndex        =   4
            Tag             =   "Partner"
            Top             =   660
            Width           =   2475
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmPrepareItem.frx":08CA
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Tag             =   "Partner"
            Top             =   2850
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   3889
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   6
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
                  LCID            =   1033
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
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               RecordSelectors =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Trans:"
            Height          =   210
            Index           =   2
            Left            =   555
            TabIndex        =   8
            Top             =   1380
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Item:"
            Height          =   210
            Index           =   1
            Left            =   585
            TabIndex        =   5
            Top             =   1050
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial No:"
            Height          =   210
            Index           =   0
            Left            =   600
            TabIndex        =   3
            Top             =   720
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frmPrepareItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
'
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub Form_Resize()
SemeruForm1.Left = 0
SemeruForm1.Top = 0
SemeruForm1.Width = Me.ScaleWidth
SemeruForm1.Height = Me.ScaleHeight
CenterForm SemeruPanels1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub
