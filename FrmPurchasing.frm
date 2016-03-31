VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPurchasing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order Pembelian"
   ClientHeight    =   6945
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "FrmPurchasing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   41
      Top             =   6375
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   1005
      EditModeReplace =   -1
      BindFormTAG     =   "PO"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   11430
      TabIndex        =   17
      Top             =   0
      Width           =   11430
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   4830
         Locked          =   -1  'True
         MaxLength       =   99
         TabIndex        =   46
         Top             =   5955
         Width           =   2490
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "keterangan2"
         DataSource      =   "MyDDE"
         Height          =   840
         Left            =   3840
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   43
         Tag             =   "PO"
         Text            =   "FrmPurchasing.frx":6852
         Top             =   5100
         Width           =   3480
      End
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "ketPayment"
         Height          =   330
         Index           =   3
         Left            =   7830
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "PO"
         Text            =   "-"
         Top             =   810
         Width           =   2205
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   10095
         MaskColor       =   &H00404080&
         Picture         =   "FrmPurchasing.frx":6856
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   818
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.ListBox ListCurrency 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   3105
         TabIndex        =   35
         Top             =   2700
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox chkPo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3225
         TabIndex        =   20
         Top             =   187
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PurchaseID"
         Height          =   330
         Index           =   0
         Left            =   1605
         MaxLength       =   25
         TabIndex        =   0
         Tag             =   "PO"
         Top             =   150
         Width           =   3315
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4530
         MaskColor       =   &H000000C0&
         Picture         =   "FrmPurchasing.frx":6BE0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1214
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Discount"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   7110
         MaxLength       =   99
         TabIndex        =   10
         Tag             =   "PO"
         Top             =   1530
         Width           =   675
      End
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "TermPayment"
         Height          =   330
         Index           =   1
         Left            =   7110
         MaxLength       =   5
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "PO"
         Top             =   810
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Close P.O"
         DataField       =   "Statussj"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3600
         TabIndex        =   19
         Tag             =   "PO"
         Top             =   915
         Width           =   1410
      End
      Begin VB.TextBox lblBank 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         Height          =   330
         Index           =   0
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "PO"
         Top             =   1206
         Width           =   2925
      End
      Begin VB.TextBox lblBank 
         Appearance      =   0  'Flat
         DataField       =   "termMethod"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   7110
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1170
         Width           =   675
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   7785
         MaskColor       =   &H00404080&
         Picture         =   "FrmPurchasing.frx":6F6A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1185
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "keterangan"
         DataSource      =   "MyDDE"
         Height          =   1200
         Left            =   105
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   12
         Tag             =   "PO"
         Top             =   5100
         Width           =   3660
      End
      Begin MSDataListLib.DataList listTipeItem 
         Height          =   1035
         Left            =   1395
         TabIndex        =   18
         Top             =   3375
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1826
         _Version        =   393216
         Appearance      =   0
         ListField       =   "tipeid"
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2730
         Left            =   75
         TabIndex        =   11
         Top             =   2070
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   4815
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         RowDividerStyle =   6
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "tipe_item"
            Caption         =   "Tipe Item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "noItem"
            Caption         =   "Kode Barang"
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
         BeginProperty Column03 
            DataField       =   "uom"
            Caption         =   "Unit"
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
            DataField       =   "QTYPO"
            Caption         =   "QTY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "POPrice"
            Caption         =   "Harga"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "CurID"
            Caption         =   "Currency"
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
         BeginProperty Column07 
            DataField       =   "rate"
            Caption         =   "Rate"
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
         BeginProperty Column08 
            DataField       =   "VAT"
            Caption         =   "PPN"
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
         BeginProperty Column09 
            DataField       =   "fldTotal"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "tmp"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "ScheduleDate"
            Caption         =   "Tgl. Kirim"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd MMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "Require Date"
         Height          =   330
         Left            =   7110
         TabIndex        =   5
         Tag             =   "PO"
         Top             =   90
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   582
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
         Format          =   71630851
         CurrentDate     =   38272
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DatePurchase"
         Height          =   315
         Left            =   1605
         TabIndex        =   1
         Tag             =   "PO"
         Top             =   502
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   71630851
         CurrentDate     =   38272
      End
      Begin MSDataListLib.DataCombo cboType 
         DataField       =   "Typetrans"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1605
         TabIndex        =   2
         Tag             =   "PO"
         Top             =   854
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "POID"
         BoundColumn     =   "CurrID"
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
      Begin VB.ListBox ListRate 
         Height          =   2010
         Left            =   5865
         TabIndex        =   36
         Top             =   2445
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   3840
         X2              =   5730
         Y1              =   6270
         Y2              =   6270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   16
         Left            =   3855
         TabIndex        =   47
         Top             =   6030
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   45
         Top             =   4890
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spesifikasi && persyaratan lain terkait keamanan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   3855
         TabIndex        =   44
         Top             =   4905
         Width           =   3390
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "RELEASED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   9315
         TabIndex        =   37
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   8565
         X2              =   9700
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   8580
         TabIndex        =   40
         Top             =   1613
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Termin Pembayaran"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   5325
         TabIndex        =   38
         Top             =   885
         Width           =   1425
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   150
         X2              =   1650
         Y1              =   1150
         Y2              =   1150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No PO"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   34
         Top             =   210
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   33
         Top             =   1245
         Width           =   570
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         DataField       =   "Address"
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
         Height          =   195
         Index           =   0
         Left            =   1620
         TabIndex        =   32
         Tag             =   "PO"
         Top             =   1575
         Width           =   885
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
         Height          =   210
         Index           =   2
         Left            =   9135
         TabIndex        =   31
         Tag             =   "PO"
         Top             =   735
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
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
         Index           =   7
         Left            =   7470
         TabIndex        =   30
         Top             =   5100
         Width           =   795
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   8985
         TabIndex        =   13
         Top             =   5070
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diskon                                              %"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   5325
         TabIndex        =   29
         Top             =   1605
         Width           =   2700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   28
         Top             =   555
         Width           =   570
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   150
         X2              =   1650
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5310
         X2              =   7110
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   150
         X2              =   1650
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   5310
         X2              =   7200
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   7515
         X2              =   10860
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   150
         X2              =   1650
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   8985
         TabIndex        =   14
         Top             =   5370
         Width           =   2235
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   8985
         TabIndex        =   15
         Top             =   5670
         Width           =   2235
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   8985
         TabIndex        =   16
         Top             =   5970
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diskon Pembelian"
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
         Index           =   3
         Left            =   7470
         TabIndex        =   27
         Top             =   5400
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
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
         Index           =   10
         Left            =   7470
         TabIndex        =   26
         Top             =   5700
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   11
         Left            =   7470
         TabIndex        =   25
         Top             =   6000
         Width           =   420
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   7515
         X2              =   10890
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   7515
         X2              =   10845
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   7515
         X2              =   10875
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe PO"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   24
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Kebutuhan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   13
         Left            =   5325
         TabIndex        =   23
         Top             =   150
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   5325
         X2              =   7380
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label LblDeliVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7110
         TabIndex        =   22
         Tag             =   "PO"
         Top             =   450
         Width           =   675
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   5310
         X2              =   7200
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   5310
         X2              =   7125
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cara Pembayaran"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   14
         Left            =   5325
         TabIndex        =   21
         Top             =   1245
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toleransi Pengiriman                        ( Hari )"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   5325
         TabIndex        =   6
         Top             =   510
         Width           =   3045
      End
   End
End
Attribute VB_Name = "FrmPurchasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMytr                                           As New DBQuick
Private RcUang                                            As New DBQuick
Private RcDetail                                          As New DBQuick
Attribute RcDetail.VB_VarHelpID = -1
Private RcPartner                                         As New DBQuick
Private RcPOType                                          As New DBQuick
Private RcTipeItem                                        As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private MEdit, mEditPO, mFirstCaller, mVarDetailPOClose   As Boolean
Private mAccount                                          As String
Private isHistoryMode As Boolean
Dim SQLInit As String
Private pWhere As String
Private ErrBtn As Integer
Private doPrint As Boolean
Private CurrSPPID As String


Public Property Let IDParams(vData As String)
   isHistoryMode = True
   pWhere = vData
End Property

Private Sub CboBayar_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub


Private Sub CboUang_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_Click(Index As Integer)
If Index = 0 Then
   Select Case UCase(Trim(cboType.Text))
      Case "NORMAL"
         OpenPartner 0
      Case "MRP"
         OpenPartner 5
      Case "ORDER"
         OpenPartner 6
   End Select
Else
   OpenPartner Index
End If
End Sub

Private Sub cmdLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then OpenPartner Index
End Sub

Private Sub DGPurchase_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If MEdit = True Then
   '4 = QtyPO
   '5 = PO Price
'   If ColIndex = 4 Then
      If IsStatusPO(MyDDE.ChildRecordset.Fields("NoItem")) = True Then
         MessageBox "Kode Barang " & vbCrLf & MyDDE.ChildRecordset.Fields("NoItem") & vbCrLf & "tidak bisa diedit,karena barang sudah dikirim Oleh Supplier " & vbCrLf & lblBank(0) & vbCrLf & " dan telah diterima bagian gudang.", "Peringatan", msgOkOnly
         DGPurchase.AllowUpdate = False
         DGPurchase.Columns(4).Value = MyDDE.ChildRecordset.Fields("QTYPO")
         mVarDetailPOClose = True
      Else
         DGPurchase.AllowUpdate = True
      End If
'   End If
End If
End Sub

Private Sub DGPurchase_ButtonClick(ByVal ColIndex As Integer)
   'colIndex 1 = tipe item
   '        6 = Currency
   If MEdit Then
      Select Case ColIndex
         Case 0:
            listTipeItem.Visible = True
            If MyDDE.CancelTrans = False Then
               listTipeItem.Move DGPurchase.Columns(0).Left + 100, (DGPurchase.RowTop(DGPurchase.row) + DGPurchase.Top + 250)
            End If
         Case 1:
            Select Case Trim(UCase(cboType.Text))
               Case "NORMAL":
                  Select Case MyDDE.ChildRecordset.Fields("tipe_item")
                     Case "I":      OpenPartner 3
                     Case "FA":     OpenPartner 7
                     Case "CHARGE": OpenPartner 8
                  End Select
               Case "ORDER":        OpenPartner 1
               Case "MRP":          OpenPartner 5
            End Select
         Case 6:
            ListCurrency.Visible = True
            If Not MyDDE.CancelTrans Then
               ListCurrency.Move DGPurchase.Columns(6).Left + 100, (DGPurchase.RowTop(DGPurchase.row) + DGPurchase.Top + 250)
            End If
      End Select
   End If
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DTPicker1_Change()
If MEdit = True Then If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.Fields("ScheduleDate").Value = DTPicker1.Value + CDbl(txtBox(1))
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub OpenTipeItem()
   RcTipeItem.DBOpen "select tipeid,keterangan from tipe_itemtrans", CNN, lckLockReadOnly
   Set listTipeItem.RowSource = RcTipeItem.DBRecordset
   Set listTipeItem.DataSource = MyDDE.ChildRecordset
   listTipeItem.BoundColumn = "tipeid"
   listTipeItem.DataField = "tipe_item"
   listTipeItem.ListField = "keterangan"
End Sub

Private Sub OpenPOType()
   RcPOType.DBOpen "Select * from tipe_po", CNN, lckLockReadOnly
   Set cboType.RowSource = RcPOType.DBRecordset
End Sub

Private Sub Form_Load()
SQLInit = " SELECT [PO Order].PurchaseID, [PO Order].PartnerID, [PO Order].Kurs, [PO Order].DatePurchase, " & _
        " [PO Order].TermPayment, [PO Order].Taxes, [PO Order].Status, [PO Order].Periode, " & _
        " [PO Order].TypeTrans, [PO Order].TypeLoco, PartnerDB.CompanyName, PartnerDB.Address, " & _
        " PartnerDB.City, [PO Order].Account, [PO Order].Discount,[PO Order].StatusSJ," & _
        " [PO Order].[Require Date], [PO Order].TermMethod, [PO Order].keterangan, [PO Order].keterangan2, [PO Order].ketPayment, [PO Order].approved_by  " & _
        " FROM  [PO Order] INNER JOIN   PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE  [po order].typetrans <> 'SO' and [po order].typetrans <> 'PCASH' and [po order].typetrans <> 'BLANKED' and "
'GridLayout
HiasFormManTell Picture2, Me
'HiasForm Picture1, Me
mVarDetailPOClose = False
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
DTPicker2.Value = dDateBegin
MataUang
OpenPOType
With MyDDE
     .EditModeReplace = False
     Set .BindForm = Me
     .BindFormTAG = "PO"
     Set .ActiveConnection = CNN
     If Trim(pWhere = "") Then
'      .PrepareQuery = SQLInit & "([PO Order].StatusSJ = 0) AND (LEFT([PO Order].PurchaseID, 2) = 'PO') ORDER BY [PO Order].PurchaseID"
      .PrepareQuery = SQLInit & "([PO Order].StatusSJ = 0) ORDER BY [PO Order].datePurchase desc"
     Else
      .PrepareQuery = SQLInit & "[PO Order].PurchaseID ='" & pWhere & "'"
     End If
     MyDDE.SetReadOnlyMode = isHistoryMode
End With
MyDDE.SetPermissions = aksess.MayDo("Order Pembelian")
OpenTipeItem
SetLabelStatus True
End Sub

Private Sub SetLabelStatus(Optional ByAccess As Boolean)
lblStatus.FontBold = True
Select Case MyDDE.GetFieldByName("Status")
    Case 0:
        lblStatus.Caption = "OPEN"
        If ByAccess Then MyDDE.SetPermissions = aksess.MayDo("Order Pembelian")
    Case 1:
        lblStatus.Caption = "CLOSED"
        If ByAccess Then MyDDE.SetPermissions = UserEditDenied
    Case 2:
        lblStatus.Caption = "RELEASED"
        If ByAccess Then MyDDE.SetPermissions = UserEditDeleteDenied
End Select
txtBox(4).Text = IIf(IsNull(MyDDE.GetFieldByName("approved_by")), "", MyDDE.GetFieldByName("approved_by"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
RcUang.CloseDB
clsMytr.CloseDB
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next

Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPurchasing = Nothing
   pWhere = ""
isHistoryMode = False
End Sub



Private Sub ListCurrency_Click()
   '6 = currency/curID
   DGPurchase.Columns(6) = ListCurrency.Text
   ListRate.ListIndex = ListCurrency.ListIndex
   'DGPurchase.Columns(7) = ListRate.Text
   'DGPurchase.columns(9).Value = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate"))) * (Val(DGPurchase.Columns(8)) / 100) + (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate")))
   ListCurrency.Visible = False
End Sub

Private Sub listTipeItem_Click()
   listTipeItem.Visible = False
   MyDDE.ChildRecordset.Fields("tipe_item") = listTipeItem.BoundText
End Sub


Private Sub mCall_BeforeUnload()
On Error GoTo xErr
Dim rsItemDetail As New DBQuick
Dim rsDefaultCurr As New DBQuick

Select Case UCase(mCall.FromTagActive)
       Case "MASTER BARANG":
            If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Then
                  If MyDDE.ChildRecordset.Fields(0) = "" Then
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  End If
               End If
            End If
            mFirstCaller = False
            If DGPurchase.Enabled = True Then
               DGPurchase.AllowUpdate = True
               DGPurchase.col = 3
               DGPurchase.SetFocus
            End If
       Case "MASTER BANK":
            'CboUang.SetFocus
       
       Case "MASTER SUPPLIER":
            txtBox(1).SetFocus
       
       Case "SPP":
         rsItemDetail.DBOpen "select * from QuerySPpOrder where partnerID='" & MyDDE.GetFieldByName("PartnerID") & "'", CNN, lckLockReadOnly
         With rsItemDetail.DBRecordset
            If .Recordcount > 0 Then
               While Not .EOF
               
                  MyDDE.ChildRecordset.AddNew
                  '*** insert data into grid detail ***
                  MyDDE.ChildRecordset.Fields("sppid") = .Fields("SPPID")
                  MyDDE.ChildRecordset.Fields("NoItem") = .Fields("noItem")
                  MyDDE.ChildRecordset.Fields("ItemName") = .Fields("itemName")
                  MyDDE.ChildRecordset.Fields("uom") = .Fields("uom")
                  MyDDE.ChildRecordset.Fields("POPrice") = .Fields("Pricein")
                  MyDDE.ChildRecordset.Fields("vat") = 10
                  MyDDE.ChildRecordset.Fields("ScheduleDate") = Date
                  MyDDE.ChildRecordset.Fields("QTYTemp") = 0
                  MyDDE.ChildRecordset.Fields("QTYPO") = .Fields("qty_spp")
                  MyDDE.ChildRecordset.Fields("StatusTrans") = False
                  MyDDE.ChildRecordset.Fields("tipe_item") = "I"
                  MyDDE.ChildRecordset.Fields("TMP") = .Fields("Pricein")
                  If Not IsEmpty(.Fields("CurrID")) Then
                     MyDDE.ChildRecordset.Fields("CurID") = .Fields("CurrID")
                     MyDDE.ChildRecordset.Fields("Rate") = .Fields("Rate")
                  Else
                     '*** cari default currency kalo ngaak ada set IDR ***'
                     
                     rsDefaultCurr.DBOpen "select CurrID,Rate from [currency setup] where functional =1", CNN, lckLockReadOnly
                     If rsDefaultCurr.DBRecordset.Recordcount > 0 Then
                        MyDDE.ChildRecordset.Fields("CurID") = rsDefaultCurr.DBRecordset.Fields(0)
                        MyDDE.ChildRecordset.Fields("Rate") = rsDefaultCurr.DBRecordset.Fields(1)
                     Else
                        MyDDE.ChildRecordset.Fields("CurID") = "IDR"
                        MyDDE.ChildRecordset.Fields("Rate") = 1
                     End If
                    
                  End If
                  
                  DGPurchase.Columns(10).Value = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("Rate"))) * (Val(DGPurchase.Columns(8)) / 100) + (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)))
                   
                  CurrSPPID = .Fields("SPPID")
                  '***
                  
                  .MoveNext
               Wend
            End If
            
         End With
         
End Select
rsItemDetail.CloseDB
rsDefaultCurr.CloseDB
Exit Sub
xErr:
   MessageBox Err.Description, "frmPurchasing : mCall_BeforeUnload", msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub mCall_CallLinkForm()
If mCall.FromTagActive <> "MASTER BARANG" Then
   frmMasterSup.SetFocus
   frmMasterSup.ZOrder (0)
Else
   FrmItemData.SetFocus
   FrmItemData.ZOrder (0)
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Dim tt As String
Dim rsDefaultCurr As New DBQuick
If pRecordset.Recordcount <> 0 Then
Select Case TagForm:
       Case "Supplier List":
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName(1)
            MyDDE.GetFieldByName("Address") = mCall.GetFieldByName(2)
            MyDDE.GetFieldByName("termPayment") = IIf(IsNull(mCall.GetFieldByName("TermPayment")), "0", mCall.GetFieldByName("TermPayment"))
            MyDDE.GetFieldByName("termMethod") = mCall.GetFieldByName("code")
            DTPicker2.Value = DTPicker1 + IIf(IsNull(mCall.GetFieldByName("TermPayment")), "0", mCall.GetFieldByName("TermPayment"))
            MyDDE.GetFieldByName("Require Date") = DTPicker2.Value
            
       
       Case "Daftar SPP":
            'col 4 Qty PO
            '    5 PO Price
            '    6 PPN/VAT
            '    8 Total/Tmp
            MyDDE.ChildRecordset.Fields("sppid") = mCall.GetFieldByName("SPPID")
            MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("noItem")
            MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("itemName")
            MyDDE.ChildRecordset.Fields("uom") = mCall.GetFieldByName("uom")
            MyDDE.ChildRecordset.Fields("POPrice") = mCall.GetFieldByName("Pricein")
            MyDDE.ChildRecordset.Fields("vat") = 10
            MyDDE.ChildRecordset.Fields("ScheduleDate") = Date
            MyDDE.ChildRecordset.Fields("QTYTemp") = 0
            MyDDE.ChildRecordset.Fields("QTYPO") = mCall.GetFieldByName("qty_spp")
            MyDDE.ChildRecordset.Fields("StatusTrans") = False
            MyDDE.ChildRecordset.Fields("tipe_item") = "I"
            MyDDE.ChildRecordset.Fields("TMP") = mCall.GetFieldByName("Pricein")
            If Not IsEmpty(mCall.GetFieldByName("CurrID")) Then
               MyDDE.ChildRecordset.Fields("CurID") = mCall.GetFieldByName("CurrID")
               MyDDE.ChildRecordset.Fields("Rate") = mCall.GetFieldByName("Rate")
            Else
               '*** cari default currency kalo ngaak ada set IDR ***'
               
               rsDefaultCurr.DBOpen "select CurrID,Rate from [currency setup] where functional =1", CNN, lckLockReadOnly
               If rsDefaultCurr.DBRecordset.Recordcount > 0 Then
                  MyDDE.ChildRecordset.Fields("CurID") = rsDefaultCurr.DBRecordset.Fields(0)
                  MyDDE.ChildRecordset.Fields("Rate") = rsDefaultCurr.DBRecordset.Fields(1)
               Else
                  MyDDE.ChildRecordset.Fields("CurID") = "IDR"
                  MyDDE.ChildRecordset.Fields("Rate") = 1
               End If
              
            End If
            
            DGPurchase.Columns(10).Value = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("Rate"))) * (Val(DGPurchase.Columns(8)) / 100) + (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)))
            CurrSPPID = mCall.GetFieldByName("SPPID")

               
            
       Case "Inventory List", "Remindier", "FA", "CHARGE":
            'col 4 Qty PO
            '    5 PO Price
            '    6 PPN/VAT
            '    8 Total/Tmp
            MyDDE.ChildRecordset.Fields("sppid") = " "
            MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("No barang")
            MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("nama barang")
            MyDDE.ChildRecordset.Fields("uom") = mCall.GetFieldByName("UOM")
            MyDDE.ChildRecordset.Fields("POPrice") = mCall.GetFieldByName("Harga")
            MyDDE.ChildRecordset.Fields("vat") = 10 'mCall.GetFieldByName("vat")
            MyDDE.ChildRecordset.Fields("ScheduleDate") = Date 'DTPicker1.Value + CDbl(txtBox(1))
            MyDDE.ChildRecordset.Fields("QTYTemp") = 0
            MyDDE.ChildRecordset.Fields("StatusTrans") = False
            MyDDE.ChildRecordset.Fields("tipe_item") = "I"
            MyDDE.ChildRecordset.Fields("TMP") = mCall.GetFieldByName("PriceIn")
            If Not IsEmpty(mCall.GetFieldByName("CurrID")) Then
               MyDDE.ChildRecordset.Fields("CurID") = mCall.GetFieldByName("CurrID")
               MyDDE.ChildRecordset.Fields("Rate") = mCall.GetFieldByName("Rate")
            Else
               '*** cari default currency kalo ngaak ada set IDR ***'
              
               rsDefaultCurr.DBOpen "select CurrID,Rate from [currency setup] where functional =1", CNN, lckLockReadOnly
               If rsDefaultCurr.DBRecordset.Recordcount > 0 Then
                  MyDDE.ChildRecordset.Fields("CurID") = rsDefaultCurr.DBRecordset.Fields(0)
                  MyDDE.ChildRecordset.Fields("Rate") = rsDefaultCurr.DBRecordset.Fields(1)
               Else
                  MyDDE.ChildRecordset.Fields("CurID") = "IDR"
                  MyDDE.ChildRecordset.Fields("Rate") = 1
               End If
               
            End If
            DGPurchase.Columns(10).Value = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate"))) * (Val(DGPurchase.Columns(8)) / 100) + (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)))
            If chkPo.Value = 0 Then
               MyDDE.ChildRecordset.Fields("QTYPO") = 1
            Else
               MyDDE.ChildRecordset.Fields("QTYPO") = mCall.GetFieldByName(3)
               If CDbl(DGPurchase.Columns(4).Value) <> 0 Then
                  MyDDE.ChildRecordset.Fields("tmp") = CDbl((DGPurchase.Columns(4) * DGPurchase.Columns(5) * DGPurchase.Columns("rate")) * (DGPurchase.Columns(8) / 100) + (DGPurchase.Columns(4) * DGPurchase.Columns(5)))
               Else
                  DGPurchase.Columns(10).Value = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate"))) * (Val(DGPurchase.Columns(8)) / 100)
               End If
            End If
            
         Case "Term Method":
            MyDDE.GetFieldByName("TermMethod") = mCall.GetFieldByName("Kode")
            lblBank(2).Text = MyDDE.GetFieldByName("termMethod")
            
         Case "SPP":
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName(1)
            MyDDE.GetFieldByName("Address") = mCall.GetFieldByName(2)
            MyDDE.GetFieldByName("termPayment") = IIf(IsNull(mCall.GetFieldByName("TermPayment")), "0", mCall.GetFieldByName("TermPayment"))
            MyDDE.GetFieldByName("termMethod") = mCall.GetFieldByName("code")
            DTPicker2.Value = DTPicker1 + IIf(IsNull(mCall.GetFieldByName("TermPayment")), "0", mCall.GetFieldByName("TermPayment"))
            MyDDE.GetFieldByName("Require Date") = DTPicker2.Value
            
         
         
End Select
End If
Set rsDefaultCurr = Nothing
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim I As Integer
Dim mStok As Long
Dim mTmp As Variant
'col 4 Qty PO
'    5 POPrice
'    8 PPN/VAT
'    10 PPH
'    11 Total/Tmp

Select Case ColIndex
       Case 4, 5, 6, 7, 8, 9:
'            If CBool(IIf(Not IsNull(MyDDE.ChildRecordset.Fields("StatusTrans")), MyDDE.ChildRecordset.Fields("StatusTrans"), False)) = False Then
             If DGPurchase.Columns(ColIndex) = "" Or IsNull(DGPurchase.Columns(ColIndex)) Then DGPurchase.Columns(ColIndex).Value = 0
               If CDbl(DGPurchase.Columns(ColIndex).Value) <> 0 Then
                 ' mTmp = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate"))) * (Val(DGPurchase.Columns(8)) / 100) + (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate")))
                  mTmp = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5))) * (Val(DGPurchase.Columns(8)) / 100) + (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)) * Val(DGPurchase.Columns("rate")))
                  DGPurchase.Columns(10).Visible = True
                  DGPurchase.Columns(9).Visible = False
                  DGPurchase.Columns(10).Value = mTmp
               Else
                  mTmp = (Val(DGPurchase.Columns(4)) * Val(DGPurchase.Columns(5)))
                  DGPurchase.Columns(10).Visible = True
                  DGPurchase.Columns(9).Visible = False
                  DGPurchase.Columns(10).Value = mTmp
               End If
'            Else
'               MessageBox "Data Tidak Bisa Diedit Karena Digunakan Oleh Receive Notes Transaksi", "Peringatan", msgOkOnly
'               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'            End If
End Select
HitungTotal
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If MEdit = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Exit Sub
End If
With DGPurchase
     Select Case .col
            Case 0, 1, 2, 3, 6, 9, 10:
'                DGPurchase.MarqueeStyle = dbgHighlightRow
                .AllowUpdate = False
'            Case 3:
'                If IsDetailOK(MyDDE.ChildRecordset.Fields("NoItem")) = True Then
'                   DGPurchase.MarqueeStyle = dbgHighlightRow
'                   .AllowUpdate = False
''                   MessageBox "Kode Barang " & MyDDE.ChildRecordset.Fields("NoItem") & vbCrLf & " tidak bisa diedit,karena barang sudah dikirim Oleh Supplier " & vbCrLf & lblBank(0) & vbCrLf & " dan telah diterima oleh bagian gudang.", "Peringatan", msgOkOnly
'                Else
'                   DGPurchase.MarqueeStyle = dbgFloatingEditor
'                   .AllowUpdate = True
'                End If
            Case Else:
'                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
     End Select
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_OnReverseAction()
   If MessageBox("Apakah Data ini akan di Reverse ?", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
      If MyDDE.GetFieldByName("status") = 2 Then
         SendDataToServer "update [po Order] set status=0 where purchaseID='" & MyDDE.GetFieldByName("PurchaseID") & "'"
         MyDDE.GetFieldByName("status") = 0
         MyDDE.RefreshControl
      Else
         MessageBox "Data Tidak bisa di Reverse, Data sudah di Tutup "
      End If
   End If
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
ErrBtn = AdReasonActiveDb
Select Case AdReasonActiveDb
       Case tmbAddNew:
       Case tmbEdit:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = CBool(IsHeaderOk(txtBox(0)))
               If MyDDE.CancelTrans = True Then MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi sudah divalidasi."
            End If
       Case tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = CBool(IsHeaderOk(txtBox(0)))
               If MyDDE.CancelTrans = True Then MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi sudah divalidasi."
            
               If cboType.Text = "ORDER" Then
                  If MyDDE.ChildRecordset.Recordcount > 0 Then
                     SendDataToServer "update spp_line set status=0 where sppid='" & _
                                MyDDE.ChildRecordset.Fields("SPPID") & "' and noItem ='" & _
                                MyDDE.ChildRecordset.Fields("noItem") & "'"
                  End If
               End If
            End If
       
       Case tmbDetail:
            MyDDE.CancelTrans = IsStatusPO
            If MyDDE.CancelTrans = False Then
                If MyData.CheckGridKosong(MyDDE.ChildRecordset, "fldtotal") = True Then
                   MyDDE.CancelTrans = True
                   MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly
                End If
            Else
               MessageBox "Tidak bisa menambah detail PO ,karena barang sudah dikirim Oleh Supplier " & lblBank(0) & " dan telah diterima bagian gudang.", "Peringatan", msgOkOnly
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then                     'CekGridKosong = False And
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                  PrepareQuery
               Else
                  MyDDE.IsChildMemberReady = False
                  MyDDE.CancelTrans = True
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbCancel:
         If (MyDDE.ChildRecordset.Recordcount > 0) And (cboType.Text = "ORDER") Then
            SendDataToServer "update spp_line set status=0 where sppid='" & _
                             MyDDE.ChildRecordset.Fields("SPPID") & "' and noItem ='" & _
                             MyDDE.ChildRecordset.Fields("noItem") & "'"
         End If
       Case tmbPrint:
         doPrint = True
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'On Error Resume Next
Dim IDGen As New IDGenerator
Dim newPO As String

'lblBank(0).Enabled = False
AdReasonActiveDb = ErrBtn
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'col 9 Total/fldTotal
            '    10 Total/Tmp

            MEdit = True
            DTPicker1.Value = Now
            DTPicker2.Value = CDate(Format(Date, "dd/mm/yyyy"))
            MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
            MyDDE.GetFieldByName("TermPayment") = 0
            MyDDE.GetFieldByName("Discount") = 0
            MyDDE.GetFieldByName("Kurs") = 1
            MyDDE.GetFieldByName("TypeTrans") = "NORMAL"
            MyDDE.GetFieldByName("keterangan") = "-"
            MyDDE.GetFieldByName("keterangan2") = "-"
            MyDDE.GetFieldByName("ketPayment") = "-"
            MyDDE.GetFieldByName("status") = 0
            
            newPO = IDGen.GetID("PO")   'MyData.PrepareIndex(tmbTransaksiPO, 5, "1", TglIndex)
            MyDDE.GetFieldByName("PurchaseID") = newPO
            MyDDE.GetFieldByName("status") = 0
            DGPurchase.Columns(9).Visible = False
            DGPurchase.Columns(10).Visible = True
            DTPicker1.SetFocus
            chkPo.Enabled = MEdit
            Check1.Enabled = False
            Check1.Value = 0
            SetLabelStatus False
       Case tmbEdit:
            MEdit = True
            mEditPO = True
            Call DGPurchase_RowColChange(DGPurchase.row, DGPurchase.col)
            If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = MEdit
            Check1.Enabled = False
       
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail mEditPO
               UpdateStatusSPP
               UpdateInventory
               MEdit = False
               chkPo.Enabled = MEdit
               mEditPO = False
               'MyData.EditHeaderRN txtBox(0), mVarLoginActive, CboUang.BoundText, MyDDE.GetFieldByName("PartnerID"), txtBox(1), CDbl(txtBox(2)), txtBox(4), False, MyDDE.ChildRecordset
               OpenDetail txtBox(0)
               mVarDetailPOClose = False
               SetLabelStatus
            Else
               MessageBox "Detail transaksi Purchase belum ada datanya.", "Peringatan", msgOkOnly
            End If
            
       Case tmbCancel:
            'col 9 Total/fldTotal
            '    10 Total/Tmp
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               MEdit = False
               DGPurchase.Columns(9).Visible = True
               DGPurchase.Columns(10).Visible = False
               If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = False
               mVarDetailPOClose = False
             Else
               DGPurchase.Columns(9).Visible = False
               DGPurchase.Columns(10).Visible = True
               'mEdit = True
             End If
             
       Case tmbDetail:
            'col 9 Total/fldTotal
            '    10 Total/Tmp
            MyDDE.ChildRecordset.Fields("tipe_item") = "I"
'               Select Case Trim(UCase(cboType.Text))
'                  Case "NORMAL":
'                     Select Case MyDDE.ChildRecordset.Fields("tipe_item")
'                        Case "I":      OpenPartner 3
'                        Case "FA":     OpenPartner 7
'                        Case "CHARGE": OpenPartner 8
'                     End Select
'                  Case "ORDER":        OpenPartner 1
'                  Case "MRP":          OpenPartner 5
'               End Select
            OpenPartner 3
            DGPurchase.Columns(9).Visible = False
            DGPurchase.Columns(10).Visible = True
            MEdit = True
            mVarDetailPOClose = False
       
       Case tmbPrint:
            If doPrint Then
               If Not IsNull(MyDDE.GetFieldByName("approved_by")) Then
                    If MyDDE.GetFieldByName("status") = 0 Then
                       SendDataToServer "update [PO Order] set status=2 where purchaseID='" & MyDDE.GetFieldByName("PurchaseId") & "'"
                       MyDDE.GetFieldByName("status") = 2
                    End If
                    Dim aReport As New utility
                    aReport.CallReportView "select * from reportPurchasing where PurchaseID='" & MyDDE.GetFieldByName("PurchaseId") & "'", "ReportPurchasing.rpt", ReportPath, "Purchase Order", Konversi(CCur(LblAmount(3).Caption))
                    Set aReport = Nothing
                    
                    doPrint = False
                Else
                    MessageBox "Dokumen ini belum di Approve !!", "Informasi", msgOkOnly, msgInfo
                End If
            End If
            SetLabelStatus True
            'PrintToContinous
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select

cmdLink(0).Enabled = MEdit
cmdLink(1).Enabled = MEdit
cmdLink(4).Enabled = MEdit


cboType.Enabled = MEdit
Err.Clear
End Sub

Private Sub UpdateInventory()
   If MyDDE.ChildRecordset.Recordcount > 0 Then
      MyDDE.ChildRecordset.MoveFirst
      While Not MyDDE.ChildRecordset.EOF
         SendDataToServer "update inventory set partnerID='" & MyDDE.GetFieldByName("partnerID") & "',CurrID='" & MyDDE.ChildRecordset.Fields("CurID") & "' where noItem ='" & MyDDE.ChildRecordset.Fields("noItem") & "'"
         MyDDE.ChildRecordset.MoveNext
      Wend
   End If
End Sub

Private Sub UpdateStatusSPP()
   Dim rsPODetail As New DBQuick
   Dim rsSPPDetail As New DBQuick
   rsPODetail.DBOpen "select SPPID from [Detail PO] where PurchaseID='" & MyDDE.GetFieldByName("purchaseID") & "' group by SPPID", CNN
   If rsPODetail.DBRecordset.Recordcount > 0 Then
      While Not rsPODetail.DBRecordset.EOF
         rsSPPDetail.DBOpen "select * from spp_line where SPPHID='" & rsPODetail.DBRecordset.Fields(0) & "' and status = 0", CNN
         SendDataToServer "update spp_header set status = " & IIf(rsSPPDetail.Recordcount > 0, "0", "1") & " where SPPID ='" & rsPODetail.DBRecordset.Fields(0) & "'"
         SendDataToServer "update spp_line set status = 2 where SPPHID='" & rsPODetail.DBRecordset.Fields(0) & "'"
         rsPODetail.DBRecordset.MoveNext
      Wend
   End If
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
lblBank(2).Text = IIf(IsNull(MyDDE.GetFieldByName("TermMethod")), "", MyDDE.GetFieldByName("TermMethod"))

SetLabelStatus True
OpenDetail MyDDE.GetFieldByName("PurchaseID")
HitungTotal
ListTotalDeliver MyDDE.GetFieldByName("PurchaseID")
MEdit = False
'MyDDE.SetPermissions = aksess.MayDo("Order Pembelian")
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0: 'Order normal
            RcPartner.DBOpen MyData.UploadQuery("Supplier"), CNN, lckLockReadOnly
       Case 1: 'Detail Order SPP
            RcPartner.DBOpen "select * from QuerySPpOrder where partnerID='" & MyDDE.GetFieldByName("PartnerID") & "'", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT [Remainder PO].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100)   + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", CNN, lckLockReadOnly
       Case 3:  'detail order normal
            RcPartner.DBOpen "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOMPurchase as uom, PPn,PriceIn AS Harga,inventory.CurrID,[Currency Setup].rate FROM Inventory left outer join [currency Setup] on Inventory.currid = [currency setup].currID WHERE (Manufacture = 0) ORDER BY ItemName", CNN, lckLockReadOnly
            mFirstCaller = True
       Case 4:  'termpayment
            RcPartner.DBOpen "Select Code as Kode, Description as Keterangan,  [Bal_ Account Type], [Bal_ Account No_] from TermMethod ", CNN, lckLockReadOnly
       Case 5: 'order MRP
            RcPartner.DBOpen "Select * from QueryOrderMRP ", CNN, lckLockReadOnly
       Case 6: 'Order SPP
            RcPartner.DBOpen "select * from QuerySPHSupplier", CNN, lckLockReadOnly
       Case 7: 'FA
            RcPartner.DBOpen "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOMPurchase as uom, PPn,PriceIn AS Harga,inventory.CurrID,[Currency Setup].rate FROM Inventory left outer join [currency Setup] on Inventory.currid = [currency setup].currID WHERE (noGroup='FA') ORDER BY ItemName", CNN, lckLockReadOnly
       Case 8: 'CHARGE
            RcPartner.DBOpen "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOMPurchase as uom, PPn,PriceIn AS Harga,inventory.CurrID,[Currency Setup].rate FROM Inventory left outer join [currency Setup] on Inventory.currid = [currency setup].currID WHERE (noGroup='CH') and (categID='JS') ORDER BY ItemName", CNN, lckLockReadOnly
End Select

If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Supplier List"
            mCall.txtCari = lblBank(0)
            mCall.CaptionLink = "Supplier"
          Case 1:
            mCall.FromTagActive = "Daftar SPP"
            mCall.CaptionLink = "Daftar SPP"
          Case 2:
            mCall.FromTagActive = "Remindier"
            mCall.txtCari = lblBank(1)
          Case 3:
            mCall.FromTagActive = "Inventory List"
            mCall.CaptionLink = "Barang"
            If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
          Case 4:
            mCall.FromTagActive = "Term Method"
            mCall.CaptionLink = "Term Method"
          Case 5:
            mCall.FromTagActive = "Item Charge"
            mCall.CaptionLink = "Item Charge"
          Case 6:
            mCall.FromTagActive = "SPP"
            mCall.CaptionLink = "SPP"
          Case 7:
            mCall.FromTagActive = "FA"
            mCall.CaptionLink = "FA"
          Case 8:
            mCall.FromTagActive = "CHARGE"
            mCall.CaptionLink = "CHARGE"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
'col 9 Total/fldTotal
'    10 Total/Tmp
Set RcDetail = New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
'RcDetail.DBOpen " SELECT [Detail PO].NoItem, Inventory.ItemName, Inventory.uom, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate,  [Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2)   + ([Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2))   * ROUND([Detail PO].VAT / 100, 2) AS FldTotal, [Detail PO].POPrice AS TMP, [Detail PO].PurchaseID, [Detail PO].QTYTemp, [Detail PO].StatusTrans, [Detail PO].SPPID, [detail PO].tipe_item,[detail PO].curID,[detail PO].rate " & _
'                " FROM [Detail PO] INNER JOIN  Inventory ON [Detail PO].NoItem = Inventory.NoItem INNER JOIN [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE     ([Detail PO].PurchaseID = N'" & ParameterString & "') ORDER BY [Detail PO].NoItem", CNN, lckLockBatch

RcDetail.DBOpen "SELECT [Detail PO].NoItem, Inventory.ItemName, Inventory.uomPurchase as uom, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate,  " & _
       "(([Detail PO].QTYPO * [Detail PO].POPrice * [Detail PO].Rate) - (([Detail PO].QTYPO * [Detail PO].POPrice * [Detail PO].Rate) * ROUND([PO Order].Discount / 100, 2)) + (([Detail PO].QTYPO * [Detail PO].POPrice * [Detail PO].Rate)*(ROUND([Detail PO].VAT / 100, 2)))) as fldTotal," & _
       "[Detail PO].QTYPO as TMP," & _
       "[Detail PO].PurchaseID , [Detail PO].QTYTemp, [Detail PO].StatusTrans, [Detail PO].SPPID, [Detail PO].tipe_item, [Detail PO].curID, [Detail PO].Rate " & _
       " FROM [Detail PO] INNER JOIN  Inventory ON [Detail PO].NoItem = Inventory.NoItem INNER JOIN [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE     ([Detail PO].PurchaseID = N'" & ParameterString & "') ORDER BY [Detail PO].NoItem", CNN, lckLockBatch
               
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
DGPurchase.Columns(9).Visible = True
DGPurchase.Columns(10).Visible = False
End Sub

Private Sub SimpanDetail(ByVal Tipical As Boolean)
Dim rsCek As New DBQuick
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [Detail PO] WHERE     (PurchaseID = N'" & txtBox(0) & "')") = True Then
           Do
              If .EOF = True Then Exit Do
              
              '*** update data detail ***'
              SendDataToServer " INSERT INTO [Detail PO] ( PurchaseID, NoItem, QTYPO, ItemSupplierID, POPrice, ScheduleDate, VAT,QtyTemp,TCredit,Hpp,sppid,tipe_item,curID,rate)" & _
                               " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "', " & FQty(.Fields("QTYPO")) & ", N'" & MyDDE.GetFieldByName("PartnerID") & "', " & FQty(.Fields("POPrice")) & ", convert(Datetime,'" & Format(.Fields("ScheduleDate"), "dd/mm/yy") & "',3), " & CDbl(.Fields("VAT")) & ", " & FQty(.Fields("QTYPO")) & "," & CCur(LblAmount(3)) & "," & FQty(.Fields("POPrice")) & ",'" & .Fields("sppid") & "','" & .Fields("tipe_item") & "','" & .Fields("curID") & "'," & FQty(.Fields("rate")) & ")"
              
              If UCase(Trim(cboType.Text)) = "ORDER" Then
                  '***  Update status SPP  ***'
                  SendDataToServer "update spph_line set result = 1 where noItem='" & .Fields("NoItem") & _
                                   "' and SPPID = '" & CurrSPPID & "'"
              End If

              .MoveNext
           Loop
           End If
           .MoveLast
           DGPurchase.Refresh
           
           If UCase(Trim(cboType.Text)) = "ORDER" Then
               rsCek.DBOpen "select * from spph_item_count where SPPHID ='" & CurrSPPID & "'", CNN
               If rsCek.DBRecordset.Recordcount > 0 Then
                  If MyDDE.ChildRecordset.Recordcount >= Val(rsCek.Fields("jml")) Then
                     SendDataToServer "update spph_header set status=1 where spphID='" & CurrSPPID & "'"
                  End If
               End If
           End If
           
           
     End If
End With
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   listTipeItem.Visible = False
End Sub

Private Sub txtBox_Change(Index As Integer)
If Index = 2 And MEdit = True Then
   If txtBox(Index) = "" Then txtBox(Index) = 0
   If CInt(txtBox(Index)) > 100 Then txtBox(Index) = 0
   MyDDE.GetFieldByName("Discount") = txtBox(Index)
   HitungTotal
ElseIf Index = 1 And MEdit = True Then
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      If txtBox(Index) = "" Then txtBox(Index) = "0"
      MyDDE.ChildRecordset.Fields("ScheduleDate").Value = DTPicker1.Value + Val(txtBox(Index))
   End If
End If
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "PO-" & Format(Day(Date), "0#") & Format(Month(Date), "0#") & Right(Format(Year(Date), "0#"), 2) & "-"
End Function

Private Sub HitungTotal()
On Error Resume Next
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mDisc, mPPn, mTotal, mStDisc As Variant
Dim mTmpDisc As Byte
Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotal = 0
mDisc = 0
mPPn = 0
mStDisc = 0
LblAmount(0) = 0
LblAmount(1) = 0
LblAmount(2) = 0
LblAmount(3) = 0
mTmpDisc = IIf(Not IsNull(MyDDE.GetFieldByName("Discount")), MyDDE.GetFieldByName("Discount"), 0)
With RcTotal
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        ' 3 = QTY  4 = Harga 5 = Vat
        For I = 0 To UBound(Avdata, 2)
            If mTmpDisc > 0 Then
               mDisc = mDisc + (Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100)
               mStDisc = mStDisc + ((Avdata(3, I) * Avdata(4, I)) - ((Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100)))
            Else
               'mStDisc = mStDisc + (Avdata(3, I) * Avdata(4, I) * Avdata(15, I))
               mStDisc = mStDisc + (Avdata(3, I) * Avdata(4, I))
               mDisc = mDisc + 0
            End If
            If Avdata(5, I) > 0 Then
              'mPPn = mPPn + ((((Avdata(3, I) * Avdata(4, I) * Avdata(15, I)) - ((Avdata(3, I) * Avdata(4, I) * Avdata(15, I)) * (mTmpDisc / 100))) * (Avdata(5, I) / 100)))
               mPPn = mPPn + ((((Avdata(3, I) * Avdata(4, I)) - ((Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100))) * (Avdata(5, I) / 100)))
            Else
               mPPn = mPPn + 0
            End If
            'mTotal = mTotal + Avdata(3, I) * Avdata(4, I) * Avdata(15, I)
            mTotal = mTotal + Avdata(3, I) * Avdata(4, I)
        Next I
     Else
        mTotal = 0
     End If
End With
LblAmount(0) = FormatNumber(mTotal, 0)
LblAmount(1) = FormatNumber(mDisc, 0)
LblAmount(2) = FormatNumber(mPPn, 0)
LblAmount(3) = FormatNumber(mStDisc + mPPn, 0)
Set Avdata = Nothing
Set mTotal = Nothing
Set mDisc = Nothing
Set mPPn = Nothing
Set mStDisc = Nothing
Err.Clear
End Sub

Private Sub PrepareQuery()

On Error Resume Next
Dim strSQL As String
Dim mPoSc As String

Select Case UCase(cboType.Text)
   Case "NORMAL": mPoSc = "2"
   Case "ORDER": mPoSc = "1"
   Case "MRP": mPoSc = "3"
End Select

With MyDDE
   
      strSQL = " INSERT INTO  [PO Order] ( [REquire Date], PurchaseID, EmpID, PartnerID, DatePurchase, TermPayment, Periode, TypeTrans,Account ,Discount ,termMethod, keterangan, type_trans_order,keterangan2,ketpayment) " & _
                        " VALUES ('" & Format(DTPicker2.Value, "yyyy-MM-dd") & "',N'" & txtBox(0).Text & "',N'" & MainMenu.StatusBar1.Panels(1).Text & "', N'" & MyDDE.GetFieldByName("PartnerID") & "'," & _
                        " '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "', " & txtBox(1) & ", " & Val(Month(DTPicker1.Value)) & ", N'PO',N'" & mAccount & "' ," & FQty(txtBox(2)) & ",'" & MyDDE.GetFieldByName("termMethod") & "','" & Text1.Text & "'," & mPoSc & ",'" & Text2 & "','" & txtBox(3) & "')"
   
    .PrepareAppend = strSQL
                     
      strSQL = " UPDATE [PO Order]" & _
                       " Set [Require Date] = Convert(datetime,'" & Format(DTPicker2.Value, "dd/mm/yy") & "',3), empID=N'" & MainMenu.StatusBar1.Panels(1).Text & "', PartnerID = N'" & MyDDE.GetFieldByName("PartnerID") & _
                       "', DatePurchase = convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3), TermPayment = " & CDbl(txtBox(1)) & ", Periode = " & Val(Month(DTPicker1.Value)) & ", TypeTrans = N'" & cboType.Text & "',Account=N'" & mAccount & "',keterangan2='" & Text2 & "',ketpayment='" & txtBox(3) & "' " & _
                       ",Discount=" & FQty(txtBox(2)) & ",termMethod= '" & MyDDE.GetFieldByName("termMethod") & "', keterangan ='" & Text1.Text & "' WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (Status = 0)"
'    Debug.Print strSQL
    .PrepareUpdate = strSQL
                     
    .PrepareDelete = " DELETE FROM  [PO Order] WHERE (PurchaseID = N'" & txtBox(0) & "')"
End With
Err.Clear
End Sub

Private Function IsHeaderOk(ByVal NoPo As String) As Boolean
Dim RcIs As New DBQuick
RcIs.DBOpen "SELECT  StatusSJ FROM [PO Order] WHERE     (PurchaseID = N'" & NoPo & "')", CNN, lckLockReadOnly
IsHeaderOk = False
With RcIs
     If .Recordcount <> 0 Then IsHeaderOk = CBool(.Fields(0))
End With
RcIs.CloseDB
End Function

Private Function IsStatusPO(Optional ByVal NoItem As String) As Boolean
Dim RcIs As New DBQuick
If NoItem = "" Then
   RcIs.DBOpen "SELECT SUM(QTY_Receive) AS QTY FROM [Detail TransData] WHERE     (DNID = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
Else
   RcIs.DBOpen "SELECT     QTY_Receive AS QTY FROM         [Detail TransData] WHERE     (DNID = N'" & txtBox(0) & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
End If
With RcIs
     If .Recordcount <> 0 Then If .Fields(0) <> 0 Then IsStatusPO = True
End With
RcIs.CloseDB
End Function


Private Sub MataUang()
RcUang.DBOpen MyData.UploadQuery("mata uang"), CNN, lckLockReadOnly
'Set CboUang.RowSource = RcUang.DBRecordset
ListCurrency.Clear
ListRate.Clear
If RcUang.DBRecordset.Recordcount > 0 Then
   While Not RcUang.DBRecordset.EOF
      ListCurrency.AddItem RcUang.DBRecordset.Fields("CurrID")
      ListRate.AddItem RcUang.DBRecordset.Fields("rate")
      RcUang.DBRecordset.MoveNext
   Wend
End If
End Sub

Private Sub UpdateTotal()
Dim rcUpdate As New DBQuick
Dim iLast, mRow As Integer
Dim Avdata As Variant
Set rcUpdate.DBRecordset = MyDDE.ChildRecordset.Clone(adLockBatchOptimistic)
With rcUpdate
     If .Recordcount <> 0 Then
        mRow = MyDDE.ChildRecordset.AbsolutePosition
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        For iLast = 0 To UBound(Avdata, 2)
            .AbsolutePosition = iLast + 1
            .Fields("Tmp") = Avdata(7, iLast)
        Next iLast
     End If
End With
Set MyDDE.ChildRecordset = rcUpdate.DBRecordset.Clone(adLockBatchOptimistic)
If MyDDE.ChildRecordset.Recordcount <> 0 Then
   MyDDE.ChildRecordset.AbsolutePosition = mRow
End If
rcUpdate.CloseDB
End Sub

Private Function CekDetailItem(ByVal PoNumber As String, ByVal NoItemData As String) As Boolean
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT NoItem, PurchaseID FROM [Detail PO] WHERE     (NoItem = N'" & NoItemData & "') AND (PurchaseID = N'" & PoNumber & "')", CNN, lckLockReadOnly
If RcCek.Recordcount <> 0 Then CekDetailItem = True
RcCek.CloseDB
End Function

Private Sub ListTotalDeliver(ByVal ParamString As String)
Dim RcDN As New DBQuick
If ParamString = "" Then ParamString = "XXXXX"
RcDN.DBOpen "SELECT DateTrans FROM TransData GROUP BY DateTrans, PurchaseID HAVING      (PurchaseID = N'" & ParamString & "')", CNN, lckLockReadOnly
With RcDN
     If .Recordcount <> 0 Then
        LblDeliVer = Abs(CDate(Format(MyDDE.GetFieldByName("DatePurchase"), "dd/mm/yyyy")) - CDate(Format(.Fields(0), "dd/mm/yyyy")))
     Else
        LblDeliVer = 0
     End If
End With
End Sub

Private Function CekGridKosong() As Boolean
Dim RcKsg As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Dim Temp As String
Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcKsg
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            Temp = IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), "")
            If Temp <> "" Then
                If Val(IIf(Not IsNull(Avdata(3, I)), Avdata(3, I), 0)) <= 0# Or Val(IIf(Not IsNull(Avdata(4, I)), Avdata(4, I), 0)) <= 0# Then
                   MessageBox "Quantity Atau Harga harus diisi.", "Peringatan"
                   CekGridKosong = True
                   Exit For
                End If
            Else
               MessageBox "Data Item Tidak Lengkap.Harap Dicek Dulu", "Peringatan"
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

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
RcCek.Open "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) = N'RN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
     .Close
End With
Set RcCek = Nothing
End Function


Private Sub GridLayout()
DGPurchase.Columns(0).width = 1814.74
DGPurchase.Columns(1).width = 1500
DGPurchase.Columns(2).width = 2324.977
DGPurchase.Columns(3).width = 764.7874
DGPurchase.Columns(4).width = 764.7874
DGPurchase.Columns(5).width = 1335.118
DGPurchase.Columns(6).width = 764.7874
DGPurchase.Columns(7).width = 1440
DGPurchase.Columns(8).width = 1440
DGPurchase.Columns(9).width = 1440
End Sub


Private Sub PrintToContinous()
On Error GoTo MASALAH
Dim MaxBaris As Integer
Dim x As Integer
Dim sQty As String
Dim sNamaItem As String
Dim sPrice As String
Dim sJml As String
Dim sTerbilang As String
Dim sTglPengiriman As String
Dim sPembayaran As String

MaxBaris = 10

MyDDE.ChildRecordset.MoveFirst
While Not MyDDE.ChildRecordset.EOF
   Open "LPT1:" For Output As 1
   Print #1, Chr$(27); "@";
   Print #1, Chr$(27); "(C"; Chr$(2); Chr$(0); Chr$(188); Chr$(7);
   Print #1, Chr$(27); "!"; Chr$(1);
   Print #1, Chr$(27); "x"; Chr$(0);
   Print #1, Chr$(27); "W1"; Chr$(27); "w1";
   Print #1, Chr$(27); "W0"; Chr$(27); "w0";
   Print #1, "                                                                                   "; MyDDE.GetFieldByName("purchaseID")
   Print #1, "                                                                                   "; Format(MyDDE.GetFieldByName("DatePurchase"), "dd MMMM yyyy")
   Print #1, Chr$(27); Chr$(103);
   Print #1, "           "; MyDDE.GetFieldByName("CompanyName")
   Print #1, "           "; MyDDE.GetFieldByName("Address")
   Print #1, "           "; MyDDE.GetFieldByName("City")
   Print #1, Chr$(27)
   Print #1, Chr$(27)
   Print #1, Chr$(27)
   Print #1, Chr$(27)
   Print #1, Chr$(27)
   
   
   For x = 1 To MaxBaris
      If MyDDE.ChildRecordset.EOF Then
         Print #1, " "
      Else
         RSet sQty = Trim(Str(MyDDE.ChildRecordset.Fields("qtyPO")))
         LSet sNamaItem = Mid(MyDDE.ChildRecordset.Fields("ItemName"), 1, 40)
         RSet sPrice = Format(MyDDE.ChildRecordset.Fields("Price"), "#,##0.00")
         RSet sJml = Format(MyDDE.ChildRecordset.Fields("Price") * MyDDE.ChildRecordset.Fields("qtyPO"), "#,##0.00")
    '             1         2         3         4         5         6         7         8         9         0         1
         Print #1, x; "  "; sQty; "  "; sNamaItem; "  "; sPrice; "  "; sJml
         MyDDE.ChildRecordset.MoveNext
      End If
   Next
   Print #1, "                                                               "; LblAmount(0).Caption
   Print #1, sTerbilang
   Print #1, " "
   Print #1, sTglPengiriman; "                                               "; sPembayaran
   Print #1, " "
   Print #1, " "
   Print #1, " "
   Print #1, " "
   Print #1, " "
   Print #1, " "
   Close #1
Wend
    
Exit Sub
    
MASALAH:
    Close #1
    MessageBox Err.Description
End Sub

