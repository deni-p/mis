VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPO 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmPO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   6480
      Left            =   75
      ScaleHeight     =   6420
      ScaleWidth      =   11280
      TabIndex        =   2
      Top             =   0
      Width           =   11340
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   5490
         Left            =   165
         ScaleHeight     =   5430
         ScaleWidth      =   9990
         TabIndex        =   3
         Top             =   555
         Width           =   10050
         Begin VB.CheckBox chkPo 
            BackColor       =   &H80000010&
            Caption         =   "P.O Reminder"
            Enabled         =   0   'False
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
            Height          =   330
            Left            =   1755
            TabIndex        =   14
            Top             =   1665
            Width           =   1995
         End
         Begin VB.TextBox txtBox 
            DataField       =   "PurchaseID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   6450
            MaxLength       =   15
            TabIndex        =   13
            Tag             =   "PO"
            Top             =   135
            Width           =   2970
         End
         Begin VB.TextBox txtBox 
            DataField       =   "TermPayment"
            Height          =   315
            Index           =   1
            Left            =   1740
            MaxLength       =   5
            TabIndex        =   12
            Tag             =   "PO"
            Top             =   1275
            Width           =   690
         End
         Begin VB.TextBox txtBox 
            DataField       =   "PartnerID"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   6450
            TabIndex        =   11
            Tag             =   "PO"
            Top             =   480
            Width           =   2970
         End
         Begin VB.TextBox txtBox 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   6450
            MaxLength       =   25
            TabIndex        =   10
            Top             =   1455
            Width           =   2970
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Kurs"
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
            Left            =   6435
            MaxLength       =   5
            TabIndex        =   9
            Tag             =   "PO"
            Top             =   2130
            Width           =   1260
         End
         Begin VB.CommandButton cmdLink 
            Caption         =   "..."
            Height          =   330
            Index           =   0
            Left            =   9450
            TabIndex        =   5
            Top             =   480
            Width           =   435
         End
         Begin VB.CommandButton cmdLink 
            Caption         =   "..."
            Height          =   330
            Index           =   1
            Left            =   9450
            TabIndex        =   4
            Top             =   1455
            Width           =   435
         End
         Begin MSDataListLib.DataCombo CboUang 
            DataField       =   "CurrID"
            Height          =   330
            Left            =   6450
            TabIndex        =   6
            Tag             =   "PO"
            Top             =   1800
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "Currency Name"
            BoundColumn     =   "CurrID"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Bindings        =   "FrmPO.frx":08CA
            Height          =   2235
            Left            =   120
            TabIndex        =   7
            Tag             =   "Partner"
            Top             =   2610
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   3942
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "NoItem"
               Caption         =   "Item/Service"
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
               DataField       =   "ItemName"
               Caption         =   "Nama Item/Service"
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
               DataField       =   "ItemSupplierID"
               Caption         =   "Sup. Item Code"
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
               DataField       =   "QTYPO"
               Caption         =   "QTY"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "POPrice"
               Caption         =   "Harga"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "VAT"
               Caption         =   "PPN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "FldTotal"
               Caption         =   "Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "tmp"
               Caption         =   "Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "ScheduleDate"
               Caption         =   "Tgl. Kirim"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "d/MMM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2280.189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1260.284
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   510.236
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   434.835
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1260.284
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1260.284
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1140.095
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo CboLoco 
            DataField       =   "TypeLoco"
            Height          =   315
            Left            =   1740
            TabIndex        =   8
            Tag             =   "PO"
            Top             =   480
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "TypeLoco"
            BoundColumn     =   "TypeLoco"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "DatePurchase"
            Height          =   330
            Left            =   1740
            TabIndex        =   15
            Tag             =   "PO"
            Top             =   135
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dddd dd/MMMM/yyyy"
            Format          =   22740995
            CurrentDate     =   38272
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toleransi Deliver:             /Hari"
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
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   30
            Top             =   945
            Width           =   2790
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P.O. ID:"
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
            Height          =   210
            Index           =   0
            Left            =   5685
            TabIndex        =   29
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. PO:"
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
            Height          =   210
            Index           =   1
            Left            =   975
            TabIndex        =   28
            Top             =   180
            Width           =   705
         End
         Begin VB.Label LblDeliVer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   1740
            TabIndex        =   27
            Top             =   975
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Term Bayar:             /Hari"
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
            Height          =   210
            Index           =   3
            Left            =   615
            TabIndex        =   26
            Top             =   1335
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner ID:"
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
            Height          =   210
            Index           =   4
            Left            =   5370
            TabIndex        =   25
            Top             =   525
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Partner:"
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
            Height          =   210
            Index           =   5
            Left            =   5130
            TabIndex        =   24
            Top             =   1515
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency:"
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
            Height          =   210
            Index           =   6
            Left            =   5520
            TabIndex        =   23
            Top             =   1815
            Width           =   870
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner ID"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   6450
            TabIndex        =   22
            Top             =   870
            Width           =   840
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner Name"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   6450
            TabIndex        =   21
            Top             =   1215
            Width           =   1125
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CurrID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   7800
            TabIndex        =   20
            Tag             =   "PO"
            Top             =   2205
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Height          =   210
            Index           =   7
            Left            =   7275
            TabIndex        =   19
            Top             =   5010
            Width           =   525
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   7845
            TabIndex        =   18
            Top             =   5010
            Width           =   2070
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kurs:"
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
            Height          =   210
            Index           =   8
            Left            =   5835
            TabIndex        =   17
            Top             =   2145
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loco/Franco:"
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
            Height          =   210
            Index           =   9
            Left            =   480
            TabIndex        =   16
            Top             =   525
            Width           =   1200
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   2940
      Width           =   4680
      Begin VB.Label Batal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmPO.frx":08DF
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   10920
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   31
      Top             =   2475
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   820
      BindFormTAG     =   "Partner"
   End
End
Attribute VB_Name = "FrmPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Mcaption As String
Private WithEvents RcDetail As Recordset
Attribute RcDetail.VB_VarHelpID = -1
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner As Recordset
Private RcRemind As New Recordset
Private RcUang As Recordset
Private MyData As New clsTransaksi
Private mEdit, mEditPO As Boolean
Private mAccount As String
