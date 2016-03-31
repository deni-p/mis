VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmSalesQuote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quotation Penjualan"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSalesQuote.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11460
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5970
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   6060
      Left            =   0
      ScaleHeight     =   6060
      ScaleWidth      =   11520
      TabIndex        =   23
      Top             =   -15
      Width           =   11520
      Begin MSDataListLib.DataCombo CboUang 
         DataField       =   "currid"
         Height          =   330
         Left            =   6555
         TabIndex        =   8
         Tag             =   "PO"
         Top             =   450
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "TermPayment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1485
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "PO"
         Top             =   1890
         Width           =   690
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PurchaseID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "PO"
         Top             =   97
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Customer PO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1470
         MaxLength       =   25
         TabIndex        =   5
         Tag             =   "PO"
         Top             =   1530
         Width           =   3450
      End
      Begin VB.TextBox lblBank 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1485
         TabIndex        =   4
         Tag             =   "PO"
         Top             =   810
         Width           =   3105
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DatePurchase"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1485
         TabIndex        =   2
         Tag             =   "PO"
         Top             =   450
         Width           =   3450
         _ExtentX        =   6085
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
         Format          =   71630848
         CurrentDate     =   38272
      End
      Begin VB.TextBox lblBank 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1530
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.TextBox lblBank 
         Appearance      =   0  'Flat
         DataField       =   "termmethod"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   6560
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1177
         Width           =   3105
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   6560
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "PO"
         Top             =   810
         Width           =   1260
      End
      Begin VB.TextBox txtBox 
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
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   6560
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "PO"
         Top             =   97
         Width           =   1260
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   9660
         MaskColor       =   &H00404080&
         Picture         =   "frmSalesQuote.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1185
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   9660
         Picture         =   "frmSalesQuote.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1545
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4590
         Picture         =   "frmSalesQuote.frx":6F66
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   825
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   9300
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "PO"
         Top             =   810
         Width           =   690
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2040
         Left            =   105
         TabIndex        =   15
         Tag             =   "Partner"
         Top             =   2295
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   3598
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
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
            Caption         =   "Satuan"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
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
            BeginProperty Column07 
               Alignment       =   1
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   8865
         TabIndex        =   19
         Top             =   4395
         Width           =   2460
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
         Height          =   330
         Index           =   1
         Left            =   8865
         TabIndex        =   20
         Top             =   4740
         Width           =   2460
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
         Height          =   330
         Index           =   2
         Left            =   8865
         TabIndex        =   21
         Top             =   5085
         Width           =   2460
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
         Height          =   330
         Index           =   3
         Left            =   8865
         TabIndex        =   22
         Top             =   5430
         Width           =   2460
      End
      Begin VB.Label LbLInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1485
         TabIndex        =   18
         Top             =   5085
         Width           =   2160
      End
      Begin VB.Label LbLInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1485
         TabIndex        =   17
         Top             =   4740
         Width           =   2160
      End
      Begin VB.Label LbLInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1485
         TabIndex        =   16
         Top             =   4395
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   45
         Top             =   518
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   5385
         TabIndex        =   44
         Top             =   885
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penjualan Persediaan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   6855
         TabIndex        =   43
         Top             =   4463
         Width           =   1545
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CurrID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   7905
         TabIndex        =   42
         Tag             =   "PO"
         Top             =   885
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner Name"
         DataField       =   "Address"
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
         Index           =   1
         Left            =   1725
         TabIndex        =   41
         Tag             =   "PO"
         Top             =   1590
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         DataField       =   "Address"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   1485
         TabIndex        =   40
         Tag             =   "PO"
         Top             =   1230
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mata Uang"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   5385
         TabIndex        =   39
         Top             =   525
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank "
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   5385
         TabIndex        =   38
         Top             =   1605
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   37
         Top             =   885
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batas Bayar                            /Hari    Toleransi                 /Hari"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   36
         Top             =   1958
         Width           =   4410
      End
      Begin VB.Label LblDeliVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   3675
         TabIndex        =   35
         Top             =   1958
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Order"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   34
         Top             =   165
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer PO"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   33
         Top             =   1605
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1560
         Y1              =   412
         Y2              =   412
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   1560
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   135
         X2              =   1560
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   135
         X2              =   1560
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line1 
         Index           =   6
         Visible         =   0   'False
         X1              =   5355
         X2              =   6780
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5370
         X2              =   6795
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   5370
         X2              =   6795
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   6825
         X2              =   9000
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5370
         X2              =   6795
         Y1              =   1492
         Y2              =   1492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount                    %"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   8610
         TabIndex        =   32
         Top             =   870
         Width           =   1680
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   135
         X2              =   1560
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   6825
         X2              =   9000
         Y1              =   5745
         Y2              =   5745
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   6825
         X2              =   9000
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   6825
         X2              =   9000
         Y1              =   5055
         Y2              =   5055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   11
         Left            =   6855
         TabIndex        =   31
         Top             =   5498
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diskon Penjualan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   12
         Left            =   6855
         TabIndex        =   30
         Top             =   5153
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   13
         Left            =   6855
         TabIndex        =   29
         Top             =   4808
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Penjualan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   14
         Left            =   5385
         TabIndex        =   28
         Top             =   165
         Width           =   1050
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   5370
         X2              =   6870
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Availability"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   15
         Left            =   165
         TabIndex        =   27
         Top             =   4470
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   16
         Left            =   165
         TabIndex        =   26
         Top             =   4815
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Price"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   17
         Left            =   165
         TabIndex        =   25
         Top             =   5153
         Width           =   765
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   135
         X2              =   1560
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   135
         X2              =   1560
         Y1              =   5055
         Y2              =   5055
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   135
         X2              =   1560
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Method"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   18
         Left            =   5385
         TabIndex        =   24
         Top             =   1245
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmSalesQuote"
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
Private MEdit, mEditSQ, mFirstCaller As Boolean
Private pWhere As String
Private pMode As String
Private SalesType As String
Dim SQLInit As String

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim I As Integer
Dim mTmp, mStok As Variant
Select Case ColIndex
       Case 3, 4, 5:
            If CBool(IIf(Not IsNull(MyDDE.ChildRecordset.Fields("StatusTrans")), MyDDE.ChildRecordset.Fields("StatusTrans"), False)) = False Then
               mStok = CekStock(MyDDE.ChildRecordset.Fields("NoItem")) - MyDDE.ChildRecordset.Fields("QtyPo")
               If mStok < 0 Then
                  MessageBox "Stock Tidak Cukup Untuk Melakukan Transaksi." & vbCrLf & "Stok Kurang -> " & mStok & " Untuk Memenuhi Transaksi Sales Quote", "Peringatan", msgOkOnly, msgExclamation
                  MyDDE.ChildRecordset.Fields("QtyPo") = 0
               Else
                    If DGPurchase.Columns(ColIndex) = "" Or IsNull(DGPurchase.Columns(ColIndex)) Then DGPurchase.Columns(ColIndex) = 0
                    If CDbl(DGPurchase.Columns(ColIndex).Value) <> 0 Then
                       mTmp = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
                       DGPurchase.Columns(7).Value = mTmp
                    Else
                       mTmp = (DGPurchase.Columns(3) * DGPurchase.Columns(4))
                       DGPurchase.Columns(7).Value = mTmp
                    End If
                    MyDDE.ChildRecordset.Fields("Qtytemp") = MyDDE.ChildRecordset.Fields("QtyPo")
               End If
            Else
               MessageBox "Data Tidak Bisa Diedit Karena Digunakan Oleh Receive Notes Transaksi", "Peringatan", msgOkOnly, msgExclamation
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
            End If
End Select
HitungTotal
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Exit Sub
End If
With DGPurchase
     Select Case .col
            Case 3, 4, 5, 8:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
            Case Else:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = False
     End Select
End With
End Sub

Private Sub Form_Load()
On Error Resume Next
'SQLInit = " SELECT [PO Order].PurchaseID, [PO Order].PartnerID, [PO Order].Kurs, [PO Order].DatePurchase, [PO Order].TermPayment, [PO Order].Taxes, [PO Order].Status,  [PO Order].Periode, [PO Order].TypeTrans, [PO Order].TypeLoco, [PO Order].CurrID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, " & _
'                     " [PO Order].Account, [PO Order].[GDG ID],[PO Order].[Customer PO],[PO Order].Discount FROM         [PO Order] INNER JOIN PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE     ([PO Order].TypeTrans = 'SC' and "

SQLInit = " SELECT [PO Order].PurchaseID, [PO Order].PartnerID, [PO Order].Kurs, [PO Order].DatePurchase, [PO Order].TermPayment, [PO Order].Taxes, [PO Order].Status,  [PO Order].Periode, [PO Order].TypeTrans, [PO Order].TypeLoco, [PO Order].CurrID,[PO Order].termmethod, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, " & _
                     " [PO Order].[GDG ID],[PO Order].[Customer PO],[PO Order].Discount FROM         [PO Order] INNER JOIN PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE     ([PO Order].status=0 and [PO Order].TypeTrans = 'QUOTE' and "


GridLayout
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = Now
MataUang
With MyDDE
     .EditModeReplace = False
     Set .BindForm = frmSalesQuote
    .BindFormTAG = "PO"
     Set .ActiveConnection = CNN
     If pWhere = "" Then
        SQLInit = SQLInit & " [PO Order].statusSJ=0 )"
     Else
        SQLInit = SQLInit & " [PO Order].purchaseID='" & pWhere & "' )"
     End If
     .PrepareQuery = SQLInit
End With
txtBox(5) = "QUOTE"
Set DTPicker1.DataSource = MyDDE.ActiveRecordset
DTPicker1.Enabled = False
CboUang.Enabled = False
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1814.74
DGPurchase.Columns(1).width = 2324.977
DGPurchase.Columns(2).width = 764.7874
DGPurchase.Columns(3).width = 764.7874
DGPurchase.Columns(4).width = 1335.118
DGPurchase.Columns(5).width = 764.7874
DGPurchase.Columns(6).width = 1440
DGPurchase.Columns(7).width = 1440
DGPurchase.Columns(8).width = 1440
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
Select Case mCall.FromTagActive
          Case "MASTER CUSTOMER":
                If txtBox(2).Enabled = True Then txtBox(2).SetFocus
          Case "MASTER BANK":
                If CboUang.Enabled = True Then CboUang.SetFocus
          Case Else
                If mCall.FromTagActive = "MASTER BARANG" Then
                    If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
                       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgInfo
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
                End If
                If DGPurchase.Enabled = True Then DGPurchase.SetFocus
                mFirstCaller = False
End Select
Err.Clear
End Sub

Private Sub mCall_CallLinkForm()
Select Case mCall.FromTagActive
       Case "MASTER CUSTOMER":
            frmPartner.SetFocus
            frmPartner.ZOrder (0)
       Case "MASTER BANK":
            frmBankPartner.SetFocus
            frmBankPartner.ZOrder (0)
       Case "GUDANG":
            frmWareHouse.SetFocus
            frmWareHouse.ZOrder (0)
      ' Case "TERM METHOD"
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
Select Case UCase(TagForm):
       Case "MASTER CUSTOMER":
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName(1)
            MyDDE.GetFieldByName("Address") = mCall.GetFieldByName(2)
       Case "MASTER BANK":
            MyDDE.GetFieldByName("Account") = mCall.GetFieldByName(0)
            lblBank(1) = mCall.GetFieldByName(1)
            MyDDE.GetFieldByName("CurrID") = mCall.GetFieldByName(2)
       Case "MASTER BARANG":
            MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("No barang")
            MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("nama barang")
            MyDDE.ChildRecordset.Fields("ItemSupplierID") = mCall.GetFieldByName("Satuan")
            MyDDE.ChildRecordset.Fields("ScheduleDate") = DTPicker1.Value + CDbl(txtBox(1))
            MyDDE.ChildRecordset.Fields("POPrice") = mCall.GetFieldByName("Harga")
            MyDDE.ChildRecordset.Fields("vat") = 10 'mCall.GetFieldByName("PPn")
            DGPurchase.Columns(7).Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
            MyDDE.ChildRecordset.Fields("QTYPO") = 0
            If CDbl(DGPurchase.Columns(3).Value) <> 0 Then
               MyDDE.ChildRecordset.Fields("tmp") = CDbl((DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4)))
            Else
               DGPurchase.Columns(7).Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100)
            End If
        Case "TERM METHOD"
            lblBank(2).Text = mCall.GetFieldByName(0)
        
        
End Select
End If
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
txtBox(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
            mEditSQ = True
            DTPicker1.Enabled = True
            CboUang.Enabled = True
       Case tmbAddNew:
            Dim IDGen As New IDGenerator
            MEdit = True
            MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
            MyDDE.GetFieldByName("TermPayment") = 0
            MyDDE.GetFieldByName("Kurs") = 1
            MyDDE.GetFieldByName("PurchaseID") = IDGen.GetID("SQ")   'MyData.PrepareIndex(tmbTransaksiSC, 5, "1", TglIndex)
            MyDDE.GetFieldByName("Customer PO") = "-"
            MyDDE.GetFieldByName("Discount") = 0
            DGPurchase.Columns(6).Visible = False
            DGPurchase.Columns(7).Visible = True
            DTPicker1.Enabled = True
            CboUang.Enabled = True
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail mEditSQ
               MEdit = False
               mEditSQ = False
               DTPicker1.Enabled = False
               CboUang.Enabled = False
            End If
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               MEdit = False
               DGPurchase.Columns(6).Visible = True
               DGPurchase.Columns(7).Visible = False
             Else
               DGPurchase.Columns(6).Visible = False
               DGPurchase.Columns(7).Visible = True
'               mEdit = True
             End If
              Set DTPicker1.DataSource = MyDDE.ActiveRecordset
              DTPicker1.Enabled = False
              CboUang.Enabled = False
       Case tmbDetail:
            If mFirstCaller = False Then
               OpenPartner 3
               DGPurchase.Columns(6).Visible = False
               DGPurchase.Columns(7).Visible = True
               MEdit = True
            End If
       Case tmbPrint:
            CallRPTReport "Sales Contract.rpt", "Select * From [Sales Contract] where PurchaseID ='" & txtBox(0) & "'"
       Case tmbQuit:

End Select
cmdLink(0).Enabled = MEdit
cmdLink(1).Enabled = MEdit
cmdLink(4).Enabled = MEdit
Err.Clear
End Sub

Private Sub MataUang()
RcUang.DBOpen "Select * from [Currency Setup]", CNN, lckLockReadOnly
Set CboUang.RowSource = RcUang.DBRecordset
End Sub

Private Sub OpenTypeBayarPO()
'clsMytr.DBOpen "SELECT  TypeLoco, TypeFreight  FROM [Type Bayar] WHERE     ([Local] = 1) ORDER BY TypeFreight", CNN, lckLockReadOnly
'clsMytr.DBOpen "SELECT  TypeLoco, TypeFreight  FROM [Type Bayar] WHERE     ([Local] = 0) ORDER BY TypeFreight", CNN, lckLockReadOnly
'Set CboBayar.RowSource = clsMytr.DBRecordset
End Sub

Public Property Let FormMode(vMode As String)
   pMode = vMode
   If vMode = "QUOTE" Then
      SalesType = "SQ"
      Me.Caption = "Sales Quote"
   ElseIf vMode = "CONTRACT" Then
      Me.Caption = "Sales Contract"
      SalesType = "SC"
   Else
      Me.Caption = "Sales Order"
      SalesType = "SO"
   End If
End Property


Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Set RcPartner = New DBQuick
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'CUSTOMER') ORDER BY PartnerID", CNN, lckLockReadOnly
       Case 1:
            'RcPartner.DBOpen "SELECT     Account AS [No Rekening], [Bank Name] AS [Nama Bank], Currency AS [Mata Uang], [Default] FROM         [Bank Partner] WHERE     (PartnerID = N'" & MyDDE.GetFieldByName("PartnerID") & "') ORDER BY [Default], [Bank Name]", CNN, lckLockReadOnly
            RcPartner.DBOpen "SELECT     Account AS [No Rekening], [Bank Name] AS [Nama Bank], Currency AS [Mata Uang], [Default] FROM         [Bank Partner]", CNN, lckLockReadOnly
       Case 2:
            If Not IsNull(MyDDE.GetFieldByName("PartnerID")) Then
               RcPartner.DBOpen "SELECT [Gudang Customer].[GDG ID], [Gudang Customer].[Nama Gudang], [Gudang Customer].Alamat, Regional.[RG Name] FROM [Gudang Customer] INNER JOIN  Regional ON [Gudang Customer].RG = Regional.RG WHERE     ([Gudang Customer].PartnerID = N'" & MyDDE.GetFieldByName("PartnerID") & "') ORDER BY [Gudang Customer].[GDG ID]", CNN, lckLockReadOnly
            Else
               RcPartner.DBOpen "SELECT [Gudang Customer].[GDG ID], [Gudang Customer].[Nama Gudang], [Gudang Customer].Alamat, Regional.[RG Name] FROM [Gudang Customer] INNER JOIN  Regional ON [Gudang Customer].RG = Regional.RG WHERE     ([Gudang Customer].PartnerID = N'XXXXX') ORDER BY [Gudang Customer].[GDG ID]", CNN, lckLockReadOnly
            End If
       Case 3:
                          
'             RcPartner.DBOpen " SELECT dbo.Inventory.NoItem as [No Barang], dbo.Inventory.ItemName as [Nama Barang], dbo.Inventory.UOM, dbo.Inventory.PriceIn As Harga, dbo.[QTY Availlable].[QTY On Hand] as [On Hand]" & _
'                              " FROM   dbo.Inventory INNER JOIN " & _
'                              " dbo.[QTY Availlable] ON dbo.Inventory.NoItem = dbo.[QTY Availlable].NoItem " & _
'                              " where (Manufacture = '1') order by dbo.Inventory.NoItem", CNN, lckLockReadOnly
             
             RcPartner.DBOpen " SELECT     dbo.Inventory.NoItem as [No Barang], dbo.Inventory.ItemName as [Nama Barang], dbo.Inventory.UOM as Satuan, dbo.Inventory.PriceIn as Harga, dbo.inventory_qty.qty_actual as [On Hand]" & _
                              " FROM   dbo.Inventory  LEFT OUTER JOIN" & _
                              " dbo.inventory_qty ON dbo.Inventory.NoItem = dbo.inventory_qty.item_no" & _
                              " WHERE     (dbo.Inventory.Manufacture = '1')", CNN, lckLockReadOnly
             
'             RcPartner.DBOpen " SELECT dbo.Inventory.NoItem as [No Barang], dbo.Inventory.ItemName as [Nama Barang], dbo.Inventory.UOM, dbo.Inventory.PriceIn As Harga" & _
'                              " FROM   dbo.Inventory " & _
'                              " where (Manufacture = '1') order by dbo.Inventory.NoItem", CNN, lckLockReadOnly
             
             mFirstCaller = True
        Case 4:
            RcPartner.DBOpen "Select Code as Kode, Description as Keterangan,  [Bal_ Account Type], [Bal_ Account No_] from TermMethod ", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "MASTER CUSTOMER"
            mCall.CaptionLink = "Customer"
          Case 1:
            mCall.FromTagActive = "MASTER BANK"
            mCall.CaptionLink = "Bank"
            mCall.txtCari = lblBank(1)
          Case 2:
            mCall.FromTagActive = "GUDANG"
            mCall.CaptionLink = "Gudang"
            mCall.txtCari = lblBank(1)
          Case 3:
            mCall.FromTagActive = "MASTER BARANG"
            If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
          Case 4:
            mCall.FromTagActive = "TERM METHOD"
          
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgInfo
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If

End If
Exit Sub
Hell:
    Err.Clear
End Sub


Private Sub SimpanDetail(ByVal Tipical As Boolean)
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [Detail PO] WHERE   (PurchaseID = N'" & txtBox(0) & "')") = True Then
           Do
              If .EOF = True Then Exit Do
              SendDataToServer " INSERT INTO [Detail PO] ( PurchaseID, NoItem, QTYPO, ItemSupplierID, POPrice, ScheduleDate, VAT,QtyTemp,TDebet,Hpp,typeTrans)" & _
                               " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "', " & .Fields("QTYPO") & ", N'" & .Fields("ItemSupplierID") & "', " & CDbl(.Fields("POPrice")) & ", convert(Datetime,'" & Format(.Fields("ScheduleDate"), "dd/mm/yy") & "',3), " & CDbl(.Fields("VAT")) & ", " & .Fields("QTYPO") & " ," & CCur(LblAmount(3)) & "," & HppProce(.Fields("NoItem")) & ",'" & txtBox(5) & "' )"
              .MoveNext
           Loop
           End If
           .MoveLast
           DGPurchase.Refresh
     End If
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("PurchaseID")
CekBankName IIf(Not IsNull(MyDDE.GetFieldByName("Partnerid")), MyDDE.GetFieldByName("Partnerid"), "xxxxx"), IIf(Not IsNull(MyDDE.GetFieldByName("Account")), MyDDE.GetFieldByName("Account"), "xxxxx")
HitungTotal
ListTotalDeliver MyDDE.GetFieldByName("PurchaseID")
lblBank(2).Text = MyDDE.GetFieldByName("termmethod")  'untuk tampilkan TermMethod
'With MyDDE
'     If .ActiveRecordset.Recordcount <> 0 Then
'        If IIf(Not IsNull(.GetFieldByName("TypeTrans")), .GetFieldByName("TypeTrans"), "QUOTE") = "QUOTE" Then
'           'cboType.ListIndex = 0
'        Else
'          'cboType.ListIndex = 1
'        End If
'     Else
'       ' cboType.ListIndex = 0
'     End If
'End With
'If MyDDE.ChildRecordset.Recordcount <> 0 Then
'    LbLInfo(0).Caption = Format(CekStock(MyDDE.ChildRecordset.Fields("NoItem")), QtyForm) 'Availibility
'End If
If MyDDE.GetFieldByName("currid") <> "" Then
    CekCurrency MyDDE.GetFieldByName("currid")   'digunkaan untuk menampillkan currency
'   CekTypeBayar MyDDE.GetFieldByName("typeloco") 'digunakan menampilkan franco
End If
MEdit = False
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
If ParameterString = "" Then ParameterString = "xxxxxxxx"
RcDetail.DBOpen " SELECT     [Detail PO].NoItem, Inventory.ItemName, [Detail PO].ItemSupplierID, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate, [Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2) " & _
                " + ([Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2))   * ROUND([Detail PO].VAT / 100, 2) AS FldTotal, [Detail PO].POPrice AS TMP, [Detail PO].PurchaseID, [Detail PO].QTYTemp, [Detail PO].StatusTrans FROM         [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem INNER JOIN  [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE     ([Detail PO].PurchaseID = N'" & ParameterString & "') ORDER BY [Detail PO].NoItem", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
DGPurchase.Columns(6).Visible = True
DGPurchase.Columns(7).Visible = False
End Sub


Private Sub CekBankName(ByVal PartnerId As String, ByVal NoRekening As String)
Dim RcBnk As New DBQuick
'RcBnk.DBOpen "SELECT     Account, [Bank Name] FROM         [Bank Partner] WHERE     (PartnerID = N'" & PartnerId & "') AND (Account = N'" & NoRekening & "')", CNN, lckLockReadOnly
RcBnk.DBOpen "SELECT     Account, [Bank Name] FROM         [Bank Partner] WHERE      (Account = N'" & NoRekening & "')", CNN, lckLockReadOnly
With RcBnk
     If .Recordcount <> 0 Then
         lblBank(1) = .Fields(1)
     Else
         lblBank(1) = ""
     End If
End With
RcBnk.CloseDB
End Sub


Private Sub HitungTotal()
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mTotal, mPPn, mDisc, mStDisc As Variant
Dim mTmpDisc As Byte
Dim I As Integer
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
               mStDisc = mStDisc + (Avdata(3, I) * Avdata(4, I))
               mDisc = mDisc + 0
            End If
            If Avdata(5, I) > 0 Then
               mPPn = mPPn + ((((Avdata(3, I) * Avdata(4, I)) - ((Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100))) * (Avdata(5, I) / 100)))
            Else
               mPPn = mPPn + 0
            End If
            mTotal = mTotal + Avdata(3, I) * Avdata(4, I)
        Next I
     Else
        mTotal = 0
     End If
End With
LblAmount(0) = FormatNumber(mTotal, 0)
LblAmount(1) = FormatNumber(mPPn, 0)
LblAmount(2) = FormatNumber(mDisc, 0)
LblAmount(3) = FormatNumber((mTotal - mDisc) + mPPn, 0)
Set Avdata = Nothing
End Sub

Private Sub ListTotalDeliver(ByVal ParamString As String)
Dim RcDN As New DBQuick
If ParamString = "" Then ParamString = "XXXXX"
RcDN.DBOpen "SELECT DateTrans FROM TransData GROUP BY DateTrans, PurchaseID HAVING      (PurchaseID = N'" & ParamString & "')", CNN, lckLockReadOnly
With RcDN
     If .Recordcount <> 0 Then
        LblDeliVer = Abs(CDate(Format(DTPicker1.Value, "dd/mm/yyyy")) - CDate(Format(.Fields(0), "dd/mm/yyyy")))
     Else
        LblDeliVer = 0
     End If
End With
RcDN.CloseDB
End Sub

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT     SUM([Inventory Tabel].QTY_IN) - SUM([Inventory Tabel].QTY_OUT) AS QTY FROM         [Inventory Tabel] INNER JOIN                       Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
End With
RcCek.CloseDB
End Function

Private Sub PrepareQuery()
On Error Resume Next
Dim mPoSc As String
mPoSc = "QUOTE"

With MyDDE
    .PrepareAppend = " INSERT INTO  [PO Order] ( PurchaseID, PartnerID,  DatePurchase, TermPayment,  Periode,Kurs, TypeTrans,Account,CurrID,[CUSTOMER PO],Discount,TERMMETHOD,person) " & _
                     " VALUES ('" & txtBox(0).Text & "', N'" & MyDDE.GetFieldByName("PartnerID") & "',convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3) , " & txtBox(1) & ", " & Val(Month(DTPicker1.Value)) & "," & CDbl(txtBox(4)) & ", N'" & mPoSc & "','0' ,'" & CboUang.BoundText & "',N'" & ValidString(txtBox(2)) & "'," & txtBox(3) & ",'" & lblBank(2) & "','" & MainMenu.StatusBar1.Panels(1).Text & "' )"
    
                 '  MsgBox .PrepareAppend
    .PrepareUpdate = " UPDATE [PO Order]" & _
                     " Set discount=" & txtBox(3) & " ,PartnerID = N'" & MyDDE.GetFieldByName("PartnerID") & "', Kurs = " & CDbl(txtBox(4)) & ", DatePurchase = convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3), TermPayment = " & CDbl(txtBox(1)) & ", Periode = " & Val(Month(DTPicker1.Value)) & ", TypeTrans = N'" & mPoSc & "',Account='0'," & _
                     " [CUSTOMER PO] =N'" & ValidString(txtBox(2)) & "',Currid='" & CboUang.BoundText & "', termmethod ='" & lblBank(2) & "', person ='" & MainMenu.StatusBar1.Panels(1).Text & "'  WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (Status = 0)"
'
    .PrepareDelete = " DELETE FROM  [PO Order] WHERE (PurchaseID = N'" & txtBox(0) & "')"
End With
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = CBool(IsHeaderOk(txtBox(0)))
               If MyDDE.CancelTrans = True Then MessageBox "Transaksi SQ Tidak Bisa Diedit.Karena Transaksi SC Sudah Valid.", "Larangan", msgOkOnly, msgCrtical
            End If
            DTPicker1.Enabled = False
       Case tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = CBool(IsHeaderOk(txtBox(0)))
               If MyDDE.CancelTrans = True Then MessageBox "Transaksi SQ Tidak Bisa Diedit.Karena Transaksi SC Sudah Valid.", "Larangan", msgOkOnly, msgCrtical
            End If
            DTPicker1.Enabled = False
       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               If MyData.CheckGridKosong(MyDDE.ChildRecordset, "fldtotal") = True Then
'                   MyDDE.CancelTrans = True
'                   MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly
'               End If
'               'MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly
'            Else
'               MyDDE.CancelTrans = mFirstCaller
'            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Function IsHeaderOk(ByVal NoPo As String) As Boolean
Dim RcIs As New DBQuick
RcIs.DBOpen "SELECT     StatusSJ FROM  [PO Order] WHERE     (PurchaseID = N'" & NoPo & "')", CNN, lckLockReadOnly
IsHeaderOk = False
With RcIs
     If .Recordcount <> 0 Then If CBool(.Fields(0)) = True Or CBool(.Fields(1)) = True Then IsHeaderOk = True
End With
RcIs.CloseDB
End Function

Private Function CekGridKosong() As Boolean
Dim RcKsg As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcKsg
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            If Val(Avdata(3, I)) = 0 Or Val(Avdata(4, I)) = 0 Then
               MessageBox "Data Item Untuk Quantity Atau Harga Ada Yang Berisi Nol.Harap Dicek Dulu", "Peringatan", msgOkOnly, msgExclamation
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

Public Property Let IDParams(vData As String)
   pWhere = vData
End Property

Private Sub Form_Unload(Cancel As Integer)
   Set frmSalesQuote = Nothing
   pWhere = ""
End Sub


Private Function HppProce(ByVal NoItem As String) As Double
Dim RcHpp As New DBQuick
RcHpp.DBOpen "SELECT     PriceIn FROM         [Inventory Tabel] WHERE     (LockFIFO = 0) AND (QTY_IN <> 0) AND (StockTmp <> 0) AND (NoItem = N'" & NoItem & "') GROUP BY PriceIn, DateTrans ORDER BY DateTrans", CNN, lckLockReadOnly
With RcHpp
     If .Recordcount <> 0 Then
        HppProce = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        HppProce = 0
     End If
End With
RcHpp.CloseDB
End Function

Private Sub CekCurrency(ByVal CurrID As String)
Dim RcCurrency As New DBQuick

RcCurrency.DBOpen "SELECT     CuRRID, [currency Name] FROM         [Currency Setup] WHERE     (CuRRID = N'" & CurrID & "') ", CNN, lckLockReadOnly
With RcCurrency
     If .Recordcount <> 0 Then
         CboUang.Text = .Fields(1)
     Else
         CboUang.Text = ""
     End If
End With
RcCurrency.CloseDB

End Sub

Private Sub CekTypeBayar(ByVal typeLoco As String)
Dim RcTBayar As New DBQuick
RcTBayar.DBOpen "SELECT     TypeLoco,typefreight FROM         [type bayar] WHERE     (typeloco = N'" & typeLoco & "') ", CNN, lckLockReadOnly
With RcTBayar
     If .Recordcount <> 0 Then
'         CboBayar.Text = .Fields(1)
     Else
'         CboBayar.Text = ""
     End If
End With
RcTBayar.CloseDB
End Sub



