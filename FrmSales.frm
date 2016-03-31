VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmSales 
   Caption         =   "Order Penjualan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSales.frx":0000
   KeyPreview      =   -1  'True
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
      TabIndex        =   0
      Top             =   0
      Width           =   11340
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5220
         Left            =   165
         ScaleHeight     =   5160
         ScaleWidth      =   9990
         TabIndex        =   1
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
            TabIndex        =   11
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
            TabIndex        =   10
            Tag             =   "PO"
            Top             =   135
            Width           =   2970
         End
         Begin VB.TextBox txtBox 
            DataField       =   "TermPayment"
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
            Index           =   1
            Left            =   1740
            MaxLength       =   5
            TabIndex        =   9
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
            TabIndex        =   8
            Tag             =   "PO"
            Top             =   480
            Width           =   2970
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Bank Name"
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
            Index           =   3
            Left            =   6450
            MaxLength       =   25
            TabIndex        =   7
            Tag             =   "PO"
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
            Index           =   4
            Left            =   6450
            MaxLength       =   5
            TabIndex        =   6
            Tag             =   "PO"
            Top             =   2145
            Width           =   1260
         End
         Begin VB.CommandButton cmdLink 
            Caption         =   "..."
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
            Height          =   330
            Index           =   0
            Left            =   9450
            TabIndex        =   3
            Top             =   480
            Width           =   435
         End
         Begin VB.CommandButton cmdLink 
            Caption         =   "..."
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
            Height          =   330
            Index           =   1
            Left            =   9450
            TabIndex        =   2
            Top             =   1455
            Width           =   435
         End
         Begin MSDataListLib.DataCombo CboUang 
            DataField       =   "CurrID"
            Height          =   330
            Left            =   6450
            TabIndex        =   4
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
            Bindings        =   "FrmSales.frx":08CA
            Height          =   2235
            Left            =   120
            TabIndex        =   5
            Tag             =   "Partner"
            Top             =   2535
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
               Caption         =   "No Barang"
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
               Caption         =   "Nama Barang"
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
               Caption         =   "UOM"
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
               Caption         =   "Total A"
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
               Caption         =   "Total B"
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
                  ColumnWidth     =   1785.26
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2700.284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1950.236
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1860.095
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   555.024
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   2220.094
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   2220.094
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1440
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "DatePurchase"
            Height          =   330
            Left            =   1740
            TabIndex        =   12
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
            Format          =   22806531
            CurrentDate     =   38272
         End
         Begin MSDataListLib.DataCombo CboBayar 
            DataField       =   "TypeLoco"
            Height          =   330
            Left            =   1740
            TabIndex        =   13
            Tag             =   "PO"
            Top             =   495
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "TypeFreight"
            BoundColumn     =   "TypeLoco"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toleransi Deliver             /Hari"
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
            TabIndex        =   28
            Top             =   945
            Width           =   2730
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P.O. ID"
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
            TabIndex        =   27
            Top             =   180
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. PO"
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
            TabIndex        =   26
            Top             =   180
            Width           =   645
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
            TabIndex        =   25
            Top             =   960
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Term Bayar             /Hari"
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
            TabIndex        =   24
            Top             =   1335
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner ID"
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
            TabIndex        =   23
            Top             =   525
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Partner"
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
            TabIndex        =   22
            Top             =   1515
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
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
            TabIndex        =   21
            Top             =   1815
            Width           =   810
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner ID"
            DataField       =   "CompanyName"
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
            Index           =   0
            Left            =   6450
            TabIndex        =   20
            Tag             =   "PO"
            Top             =   870
            Width           =   840
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partner Name"
            DataField       =   "Address"
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
            Index           =   1
            Left            =   6450
            TabIndex        =   19
            Tag             =   "PO"
            Top             =   1215
            Width           =   1125
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CurrID"
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
            Height          =   195
            Index           =   2
            Left            =   7800
            TabIndex        =   18
            Tag             =   "PO"
            Top             =   2205
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            TabIndex        =   17
            Top             =   4845
            Width           =   465
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   16
            Top             =   4845
            Width           =   2070
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kurs"
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
            TabIndex        =   15
            Top             =   2160
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loco/Franco"
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
            TabIndex        =   14
            Top             =   525
            Width           =   1140
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   29
      Top             =   2505
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1217
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMytr As New DBQuick
Private RcUang As New DBQuick
Private RcDetail As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData As New clsTransaksi
Private mEdit, mEditPO As Boolean
Private mAccount As String

Private Sub cmdLink_Click(Index As Integer)
 OpenPartner Index
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
Call DGPurchase_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
DTPicker1.Value = dDateBegin
OpenTypeBayarPO
MataUang
With MyDDE
     .EditModeReplace = False
     Set .BindForm = FrmPurchasing
     .BindFormTAG = "PO"
     Set .ActiveConnection = Cnn
     .PrepareQuery = "SELECT     [PO Order].PurchaseID, [PO Order].PartnerID, [PO Order].Kurs, [PO Order].DatePurchase, [PO Order].TermPayment, [PO Order].Taxes, [PO Order].Status,  [PO Order].Periode, [PO Order].TypeTrans, [PO Order].TypeLoco, [PO Order].CurrID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City,  [Bank Partner].Account, [Bank Partner].[Bank Name] FROM [PO Order] INNER JOIN PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID INNER JOIN [Bank Partner] ON PartnerDB.PartnerID = [Bank Partner].PartnerID WHERE ([PO Order].TypeTrans = 'PO') ORDER BY [PO Order].PurchaseID"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
RcUang.CloseDB
clsMytr.CloseDB
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMaximized Then
   Me.Height = MainMenu.ScaleHeight
   Me.Width = MainMenu.ScaleWidth
End If
HiasForm Picture1, Me
CenterForm Picture2
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub mCall_ActionJackson(ByVal ColIndex As Integer)
Dim I As Integer
I = MessageBox("Kirim Data Ke Bagian Purchasing......?", "Peringatan", msgYesNo)
End Sub

Private Sub mCall_BeforeRefresh()
If mCall.FromTagActive = "MASTER BANK" Then mCall.SetFormat(3) = "YES/NO"

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
If pRecordset.Recordcount <> 0 Then
Select Case TagForm:
       Case "MASTER SUPPLIER":
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName(1)
            MyDDE.GetFieldByName("Address") = mCall.GetFieldByName(2)
       Case "MASTER BANK":
            mAccount = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("Bank Name") = mCall.GetFieldByName(1)
            MyDDE.GetFieldByName("CurrID") = mCall.GetFieldByName(2)
       Case "MASTER BARANG", "REMINDER":
            MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("No barang")
            MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("nama barang")
            MyDDE.ChildRecordset.Fields("ItemSupplierID") = mCall.GetFieldByName("UOM")
            MyDDE.ChildRecordset.Fields("POPrice") = 0
            MyDDE.ChildRecordset.Fields("vat") = mCall.GetFieldByName("vat")
            DGPurchase.Columns(7).Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
            If chkPo.Value = 0 Then
               MyDDE.ChildRecordset.Fields("QTYPO") = 0
            Else
               MyDDE.ChildRecordset.Fields("QTYPO") = mCall.GetFieldByName(3)
               If CDbl(DGPurchase.Columns(3).Value) <> 0 Then
                  MyDDE.ChildRecordset.Fields("tmp") = CDbl((DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4)))
               Else
                  DGPurchase.Columns("tmp").Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100)
               End If
            End If
End Select
End If
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim I As Integer
Dim mStok As Long
Dim mTmp As Variant
Select Case ColIndex
       Case 3, 4, 5:
            If CBool(IIf(Not IsNull(MyDDE.ChildRecordset.Fields("StatusTrans")), MyDDE.ChildRecordset.Fields("StatusTrans"), False)) = False Then
               If CDbl(DGPurchase.Columns(ColIndex).Value) <> 0 Then
                  mTmp = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
                  DGPurchase.Columns(7).Value = mTmp
               Else
                  mTmp = (DGPurchase.Columns(3) * DGPurchase.Columns(4))
                  DGPurchase.Columns(7).Value = mTmp
               End If
            Else
               MessageBox "Data Tidak Bisa Diedit Karena Digunakan Oleh Receive Notes Transaksi", "Peringatan", msgOkOnly
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
            End If
End Select
HitungTotal
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
'If mEdit = False Then Exit Sub
If Shift = 2 And KeyCode = vbKeyF3 Then
    mEdit = True
    DGPurchase.Columns(3) = 0
    DGPurchase.Columns(4) = 0
    DGPurchase.Columns(5) = 0
    DGPurchase.Columns(7) = 0
    MyDDE.ChildRecordset.Fields("QtyTemp") = 0
    DGPurchase.Columns(8) = CDate(Format(DTPicker1.Value, "dd/mm/yyyy")) + Val(txtBox(1))
    If chkPo.Value = 1 Then OpenPartner 2 Else OpenPartner 3
    DGPurchase.SetFocus
ElseIf Shift = 2 And KeyCode = vbKeyF2 Then
   If MyDDE.CheckEmptyControl = False Then
        mEdit = True
        If chkPo.Value = 1 Then OpenPartner 2 Else OpenPartner 3
        DGPurchase.SetFocus
   Else
      MessageBox "Data Transaksi Belum Ada.Harap Diisi Dulu.", "Peringatan"
   End If
End If
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
   Exit Sub
End If
With DGPurchase
     Select Case .Col
            Case 0, 1, 2, 6, 7:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = False
            Case Else:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
     End Select
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit, tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = CBool(IsHeaderOk(txtBox(0)))
               If MyDDE.CancelTrans = True Then
                  If Me.Caption = "P.O Transaksi" Then
                     MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi PO Sudah Valid/Closed Oleh Transaksi RN."
                  Else
                     MessageBox "Transaksi SC Tidak Bisa Diedit.Karena Transaksi SC Sudah Valid/Closed Oleh Transaksi DN."
                  End If
               End If
            End If
            'DGPurchase.Columns(0).Button = False
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                  'PrepareQuery
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
txtBox(0).Enabled = False
txtBox(2).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            mEdit = True
            mEditPO = True
            If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = mEdit
       Case tmbAddNew:
            mEdit = True
            MyDDE.GetFieldByName("DatePurchase") = CDate(dDateBegin)
            MyDDE.GetFieldByName("TermPayment") = 0
            MyDDE.GetFieldByName("Kurs") = 1
            MyDDE.GetFieldByName("PurchaseID") = MyData.PrepareIndex(tmbTransaksiPO, 5, "1", TglIndex)
            DGPurchase.Columns(6).Visible = False
            DGPurchase.Columns(7).Visible = True
            DTPicker1.SetFocus
            chkPo.Enabled = mEdit

       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail mEditPO
               mEdit = False
               chkPo.Enabled = mEdit
               mEditPO = False
            End If

       Case tmbCancel:
            mEdit = False
            DGPurchase.Columns(6).Visible = True
            DGPurchase.Columns(7).Visible = False
            If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = False

       Case tmbDetail:
            Call DGPurchase_KeyDown(vbKeyF3, 2)
       Case tmbPrint:
            CallRPTReport "Purchase Order.rpt", "Select * From [purchase Order] where PurchaseID ='" & txtBox(0) & "'"
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select
cmdLink(0).Enabled = mEdit
cmdLink(1).Enabled = mEdit
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Opendetail MyDDE.GetFieldByName("PurchaseID")
HitungTotal
ListTotalDeliver MyDDE.GetFieldByName("PurchaseID")
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:

Set RcPartner = New DBQuick
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'SUPPLIER') ORDER BY PartnerID", Cnn, lckLockReadOnly
            mCall.FromTagActive = "MASTER SUPPLIER"
            mCall.txtCari = txtBox(2)
       Case 1:
            RcPartner.DBOpen "SELECT     Account AS [No Rekening], [Bank Name] AS [Nama Bank], Currency AS [Mata Uang], [Default] FROM         [Bank Partner] WHERE     (PartnerID = N'" & txtBox(2) & "') ORDER BY [Default], [Bank Name]", Cnn, lckLockReadOnly
            mCall.FromTagActive = "MASTER BANK"
            mCall.txtCari = txtBox(3)
       Case 2:
            RcPartner.DBOpen "SELECT [Remainder PO].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100)   + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", Cnn, lckLockReadOnly
            mCall.FromTagActive = "REMINDER"
            mCall.txtCari = txtBox(3)
       Case 3:
            RcPartner.DBOpen "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn FROM   Inventory ORDER BY NoItem", Cnn, lckLockReadOnly
            mCall.FromTagActive = "MASTER BARANG"
            If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
End Select
If RcPartner.Recordcount <> 0 Then
    Set mCall.FormData = RcPartner.DBRecordset
    If mCall.FromTagActive = "MASTER BANK" Then mCall.SetFormat(3) = "YES/NO"
    mCall.Show vbModal
    Set mCall = Nothing
    If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
       DGPurchase.SetFocus
    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub Opendetail(ByVal ParameterString As String)
Set RcDetail = New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
RcDetail.DBOpen "SELECT [Detail PO].NoItem, Inventory.ItemName, [Detail PO].ItemSupplierID, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate, ( ([Detail PO].QTYPO * [Detail PO].POPrice) * ([Detail PO].VAT / 100)) + ([Detail PO].QTYPO * [Detail PO].POPrice) AS FldTotal, [Detail PO].POPrice AS TMP, [Detail PO].PurchaseID,[Detail PO].QTYTemp,[Detail PO].StatusTrans FROM [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem WHERE     ([Detail PO].PurchaseID = N'" & ParameterString & "') ORDER BY [Detail PO].NoItem", Cnn, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
DGPurchase.Columns(6).Visible = True
DGPurchase.Columns(7).Visible = False
End Sub

Private Sub SimpanDetail(ByVal Tipical As Boolean)
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [Detail PO] WHERE     (PurchaseID = N'" & txtBox(0) & "')") = True Then
           Do
              If .EOF = True Then Exit Do
              SendDataToServer " INSERT INTO [Detail PO] ( PurchaseID, NoItem, QTYPO, ItemSupplierID, POPrice, ScheduleDate, VAT,QtyTemp)" & _
                               " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "', " & .Fields("QTYPO") & ", N'" & .Fields("ItemSupplierID") & "', " & CDbl(.Fields("POPrice")) & ", convert(Datetime,'" & Format(.Fields("ScheduleDate"), "dd/mm/yy") & "',3), " & CDbl(.Fields("VAT")) & ", " & .Fields("QTYPO") & "  )"
              .MoveNext
           Loop
           End If
           .MoveLast
           DGPurchase.Refresh
     End If
End With
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "PO/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub HitungTotal()
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mTmp, mTotal As Variant
Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotal = 0
mTmp = 0
With RcTotal
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.GetRows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            mTmp = ((Avdata(3, I) * Avdata(4, I)) * (Avdata(5, I) / 100)) + (Avdata(3, I) * Avdata(4, I))
            mTotal = mTotal + mTmp
        Next I
     Else
        mTotal = 0
     End If
End With
LblAmount = FormatNumber(mTotal, 0)
Set Avdata = Nothing
RcTotal.CloseDB
End Sub

Private Sub PrepareQuery()
On Error Resume Next
Dim mPoSc As String
With MyDDE
    .PrepareAppend = " INSERT INTO  [PO Order] ( PurchaseID, PartnerID,  DatePurchase, TermPayment,  Periode,Kurs, TypeTrans,Account,TypeLoco,CurrID) " & _
                     " VALUES (N'" & txtBox(0) & "', N'" & txtBox(2) & "',convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3) , " & txtBox(1) & ", " & Val(Month(DTPicker1.Value)) & "," & CDbl(txtBox(4)) & ", N'PO',N'" & mAccount & "' ,'" & CboBayar.BoundText & "','" & CboUang.BoundText & "')"
                     
    .PrepareUpdate = " UPDATE [PO Order]" & _
                     " Set PartnerID = N'" & txtBox(2) & "', Kurs = " & CDbl(txtBox(4)) & ", DatePurchase = convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3), TermPayment = " & CDbl(txtBox(1)) & ", Periode = " & Val(Month(DTPicker1.Value)) & ", TypeTrans = N'PO',Account=N'" & mAccount & "'" & _
                     " ,Currid='" & CboUang.BoundText & "', TypeLoco = '" & CboBayar.BoundText & "' WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (Status = 0)"

    .PrepareDelete = " DELETE FROM  [PO Order] WHERE (PurchaseID = N'" & txtBox(0) & "')"
End With
Err.Clear
End Sub

Private Function IsHeaderOk(ByVal NoPo As String) As Boolean
Dim RcIs As New DBQuick
RcIs.DBOpen "SELECT  [PO Order].Status, [Detail PO].StatusTrans FROM [Detail PO] INNER JOIN [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE  ([PO Order].PurchaseID = N'" & txtBox(0) & "') GROUP BY [PO Order].Status, [Detail PO].StatusTrans HAVING      ([Detail PO].StatusTrans = 1)", Cnn, lckLockReadOnly
IsHeaderOk = False
With RcIs
     If .Recordcount <> 0 Then If CBool(.Fields(0)) = True Or CBool(.Fields(1)) = True Then IsHeaderOk = True
End With
RcIs.CloseDB
End Function

Private Sub OpenTypeBayarPO()
clsMytr.DBOpen "SELECT TypeFreight, TypeFreight AS TypeLoco FROM         [Type Bayar] ORDER BY TypeFreight", Cnn, lckLockReadOnly
Set CboBayar.RowSource = clsMytr.DBRecordset
End Sub

Private Sub MataUang()
RcUang.DBOpen "Select * from [Currency Table]", Cnn, lckLockReadOnly
Set CboUang.RowSource = RcUang.DBRecordset
End Sub

Private Sub UpdateTotal()
Dim rcUpdate As New DBQuick
Dim iLast, mRow As Integer
Dim Avdata As Variant
Set rcUpdate.DBRecordset = MyDDE.ChildRecordset.Clone(adLockBatchOptimistic)
With rcUpdate
     If .Recordcount <> 0 Then
        mRow = MyDDE.ChildRecordset.AbsolutePosition
        Avdata = .DBRecordset.GetRows(.Recordcount, adBookmarkFirst)
        For iLast = 0 To UBound(Avdata, 2)
            .AbsolutePosition = iLast + 1
            .Fields("Tmp") = Avdata(7, iLast)
        Next iLast
     End If
End With
Set MyDDE.ChildRecordset = rcUpdate.Clone(adLockBatchOptimistic)
If MyDDE.ChildRecordset.Recordcount <> 0 Then
   MyDDE.ChildRecordset.AbsolutePosition = mRow
End If
rcUpdate.CloseDB
End Sub

Private Function CekDetailItem(ByVal PoNumber As String, ByVal NoItemData As String) As Boolean
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT NoItem, PurchaseID FROM [Detail PO] WHERE     (NoItem = N'" & NoItemData & "') AND (PurchaseID = N'" & PoNumber & "')", Cnn, lckLockReadOnly
If RcCek.Recordcount <> 0 Then CekDetailItem = True
RcCek.CloseDB
End Function

Private Sub ListTotalDeliver(ByVal ParamString As String)
Dim RcDN As New DBQuick
If ParamString = "" Then ParamString = "XXXXX"
RcDN.DBOpen "SELECT DateTrans FROM TransData GROUP BY DateTrans, PurchaseID HAVING      (PurchaseID = N'" & ParamString & "')", Cnn, lckLockReadOnly
With RcDN
     If .Recordcount <> 0 Then
        LblDeliVer = Abs(CDate(Format(DTPicker1.Value, "dd/mm/yyyy")) - CDate(Format(.Fields(0), "dd/mm/yyyy")))
     Else
        LblDeliVer = 0
     End If
End With
RcDN.CloseDB
End Sub

Private Function CekGridKosong() As Boolean
Dim RcKsg As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcKsg
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.GetRows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            If Val(Avdata(3, I)) = 0 Or Val(Avdata(4, I)) = 0 Then
               MessageBox "Data Item Untuk Quantity Atau Harga Ada Yang Berisi NOl.Harap Dicek Dulu", "Peringatan"
               CekGridKosong = True
               Exit For
            End If
        Next I
     Else
        CekGridKosong = True
     End If
End With
RcKsg.CloseDB
End Function




