VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmArTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Penjualan"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmARTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10320
   Tag             =   "Invoicing"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   10320
      TabIndex        =   19
      Top             =   0
      Width           =   10320
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TransID"
         Height          =   315
         Index           =   0
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "RN"
         Top             =   105
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "DNID"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1455
         TabIndex        =   3
         Tag             =   "RN"
         Top             =   810
         Width           =   3090
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
         Height          =   1080
         Index           =   2
         Left            =   180
         MaxLength       =   200
         TabIndex        =   13
         Tag             =   "RN"
         Top             =   4320
         Width           =   4890
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   4545
         Picture         =   "frmARTrans.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   810
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "frmARTrans.frx":6BDC
         Height          =   2415
         Left            =   195
         TabIndex        =   18
         Tag             =   "Partner"
         Top             =   1575
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   4260
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NoItem"
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
            DataField       =   "QTY_OUT"
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
         BeginProperty Column03 
            DataField       =   "Price"
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
         BeginProperty Column04 
            DataField       =   "VAT"
            Caption         =   "PPn"
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
            DataField       =   "TotalA"
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
         BeginProperty Column06 
            DataField       =   "TotalB"
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
            DataField       =   "StatusItem"
            Caption         =   "Status"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         Height          =   330
         Left            =   1455
         TabIndex        =   2
         Tag             =   "RN"
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
         Format          =   71630851
         CurrentDate     =   38272
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   7305
         TabIndex        =   14
         Top             =   4095
         Width           =   2820
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
         Height          =   315
         Index           =   3
         Left            =   7305
         TabIndex        =   17
         Top             =   5085
         Width           =   2820
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
         Height          =   315
         Index           =   2
         Left            =   7305
         TabIndex        =   16
         Top             =   4755
         Width           =   2820
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
         Height          =   315
         Index           =   1
         Left            =   7305
         TabIndex        =   15
         Top             =   4425
         Width           =   2820
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   6375
         TabIndex        =   9
         Top             =   810
         Width           =   1740
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   6375
         TabIndex        =   8
         Top             =   465
         Width           =   3720
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Term"
         DataField       =   "Term"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   9240
         TabIndex        =   10
         Top             =   795
         Width           =   855
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kurs"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   6375
         TabIndex        =   11
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Term"
         DataField       =   "PurchaseID"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   1455
         TabIndex        =   5
         Tag             =   "RN"
         Top             =   1155
         Width           =   3450
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Perusahaan"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   6375
         TabIndex        =   6
         Tag             =   "RN"
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblSupplier 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "Discount"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   9240
         TabIndex        =   12
         Tag             =   "RN"
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "CompanyName"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   7515
         TabIndex        =   7
         Tag             =   "RN"
         Top             =   120
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   11
         Left            =   8415
         TabIndex        =   37
         Top             =   1185
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   36
         Top             =   4140
         Width           =   675
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner Name"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   5760
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO Date"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   5070
         TabIndex        =   34
         Top             =   855
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   5070
         TabIndex        =   33
         Top             =   165
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   32
         Top             =   495
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   31
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO No."
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   30
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   5070
         TabIndex        =   29
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Pay"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   8385
         TabIndex        =   28
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   27
         Top             =   4080
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order No."
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   26
         Top             =   1185
         Width           =   1140
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         DataField       =   "CURRID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   7335
         TabIndex        =   25
         Top             =   1200
         Width           =   315
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   195
         X2              =   1440
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   195
         X2              =   1560
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   195
         X2              =   1560
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5070
         X2              =   6375
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5070
         X2              =   6375
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   5070
         X2              =   6375
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5070
         X2              =   6375
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   8415
         X2              =   9495
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   195
         X2              =   1560
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   5070
         TabIndex        =   24
         Top             =   495
         Width           =   585
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   8400
         X2              =   9480
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   5400
         X2              =   7320
         Y1              =   5385
         Y2              =   5385
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   5400
         X2              =   7320
         Y1              =   5055
         Y2              =   5055
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   5400
         X2              =   7320
         Y1              =   4725
         Y2              =   4725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   12
         Left            =   5400
         TabIndex        =   22
         Top             =   5130
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   13
         Left            =   5400
         TabIndex        =   21
         Top             =   4815
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   14
         Left            =   5400
         TabIndex        =   20
         Top             =   4470
         Width           =   285
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   5400
         X2              =   7320
         Y1              =   4395
         Y2              =   4395
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5670
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmArTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcPartner As New DBQuick
Dim MyData As New clsTransaksi
Dim MEdit As Boolean
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private pWhere As String
Dim IDGen As New IDGenerator
Dim SQLInit As String

Public Property Let IDParams(vData As String)
   pWhere = vData
   
End Property

Private Sub cmdLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then OpenPartner Index
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
'If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
On Error GoTo 1
SQLInit = " SELECT TransID, EmpID, ID AS DeliverID, DateTrans, DateIssued, TermPayment, " & _
    " Kurs, Status, PurchaseID, RefNotes, TypeTrans, DNID, CurrID as [Mata Uang], Discount " & _
    " FROM TransData WHERE "
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
With MyDDE
     .EditModeReplace = False
     Set .BindForm = frmArTrans
     .SetPermissions = UserEditDeleteDenied
     .BindFormTAG = "RN"
     Set .ActiveConnection = CNN
     If Trim(pWhere) = "" Then
      .PrepareQuery = SQLInit & "(TypeTrans = N'AR') AND (StatusInvoice = 0) ORDER BY TransID"
     Else
      .PrepareQuery = SQLInit & " TransID='" & pWhere & "'"
     End If
End With
Set mCall = New frmCaller
Exit Sub
1:
MessageBox Err.Description, "frmartrans_form_load" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
Set mCall = Nothing
End Sub

Private Sub Form_Resize()

Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmArTrans = Nothing
pWhere = ""
End Sub

Private Sub mCall_BeforeUnload()
'Select Case mCall.FromTagActive
'       Case "Data Order Penjualan":
'            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
'       Case "MASTER BARANG":
'
'End Select
If DGPurchase.Enabled = True Then DGPurchase.SetFocus
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
Select Case TagForm:
       Case "Data Order Penjualan":
            MyDDE.GetFieldByName("dnid") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("PurchaseID") = mCall.GetFieldByName(1)
            lblSupplier(0) = mCall.GetFieldByName(3)
            lblSupplier(1) = mCall.GetFieldByName(4)
            lblSupplier(2) = mCall.GetFieldByName(5)
            lblSupplier(3) = Format(mCall.GetFieldByName(2), "dd mmmm yyyy")
            lblSupplier(4) = mCall.GetFieldByName("Kurs")
            lblSupplier(5) = FormatNumber(mCall.GetFieldByName(8), 0)
            lblSupplier(6) = mCall.GetFieldByName(1)
            lblSupplier(7) = mCall.GetFieldByName("Mata Uang")
            lblSupplier(8) = mCall.GetFieldByName("Perusahaan")
            lblSupplier(9) = mCall.GetFieldByName("Discount")
            MyDDE.GetFieldByName("Discount") = mCall.GetFieldByName("Discount")
            MyDDE.GetFieldByName("Kurs") = CDbl(lblSupplier(4))
            MyDDE.GetFieldByName("Mata Uang") = lblSupplier(7)
            MyDDE.GetFieldByName("TermPayment") = CDbl(lblSupplier(5))
            OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DNID")), MyDDE.GetFieldByName("DNID"), "XXXXXXX")
            HitungTotal
           ' IsiDetail MyDDE.GetFieldByName("PurchaseID")
       Case "BANK":
            txtBox(3) = mCall.GetFieldByName(0)
            lblSupplier(2) = mCall.GetFieldByName(1)
       Case "MASTER BARANG":
            MyDDE.ChildRecordset.Fields(0) = mCall.GetFieldByName(0)
            MyDDE.ChildRecordset.Fields(1) = mCall.GetFieldByName(1)
            MyDDE.ChildRecordset.Fields(2) = mCall.GetFieldByName(2)
End Select
Exit Sub
1:
MessageBox Err.Description, "frmartrans_mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

'Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
'Select Case ColIndex
'       Case 2:
'            mydde.childrecordset.Fields("TotalB") = ((mydde.childrecordset.Fields("qty_Out") * mydde.childrecordset.Fields("Price")) * (mydde.childrecordset.Fields("Vat") / 100)) + (mydde.childrecordset.Fields("qty_Out") * mydde.childrecordset.Fields("Price"))
'End Select
'HitungTotal
'End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 2
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
                  PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Exit Sub
2:
MessageBox Err.Description, "frmartrans_mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
txtBox(0).Enabled = False
txtBox(3).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            cmdLink(1).Enabled = MEdit
            DTPicker1.SetFocus
       Case tmbAddNew:
            MEdit = True
            'messagebox CDate(Format(dDateBegin, "dd/mm/yy"))
            DTPicker1.Value = Date
            MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
            MyDDE.GetFieldByName("TransID") = IDGen.GetID("IP") 'MyData.PrepareIndex(tmbTransaksiAR, 5, "1", TglIndex)
            MyDDE.GetFieldByName("RefNotes") = "-"
            DTPicker1.SetFocus
            cmdLink(1).Enabled = MEdit
            DGPurchase.Columns(5).Visible = False
            DGPurchase.Columns(6).Visible = True
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail

               MEdit = False
               cmdLink(1).Enabled = MEdit
            End If
       Case tmbPrint:
            CallRPTReport "INVOICE.rpt", "SELECT  * From Invoice Where TransID='" & txtBox(0) & "'"
       Case tmbCancel:
            MEdit = False
            cmdLink(1).Enabled = MEdit
End Select
Exit Sub
1:
MessageBox Err.Description, "frmartrans_mydde_afterpreparedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DNID")), MyDDE.GetFieldByName("DNID"), "XXXXXXX")
OpenDetailPart IIf(Not IsNull(MyDDE.GetFieldByName("PurchaseID")), MyDDE.GetFieldByName("PurchaseID"), "XXXXXXX")
HitungTotal
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1:
            RcPartner.DBOpen " SELECT     TransData.TransID AS [SC NUMBER], [PO Order].PurchaseID AS [SC NUMBER], [PO Order].DatePurchase AS [TGL. PO],                        TransData.PartnerId AS [PARTNER ID], PartnerDB.CompanyName AS Perusahaan, PartnerDB.Address AS ALAMAT, PartnerDB.City AS Kota,                        [PO Order].Kurs, [PO Order].TermPayment AS Term, TransData.Status, [PO Order].CurrID AS [Mata Uang], [PO Order].Discount FROM         TransData INNER JOIN                       PartnerDB ON TransData.PartnerId = PartnerDB.PartnerID INNER JOIN                       [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID WHERE     (TransData.TypeTrans = N'DN') AND (TransData.Status = 0)", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT Inventory.NoItem, Inventory.ItemName, Inventory.UOM, Inventory.PPn, MAX([Inventory Tabel].PriceIn) * (Inventory.PPn / 100)  + MAX([Inventory Tabel].PriceIn) * (Inventory.Markup / 100) + MAX([Inventory Tabel].PriceIn) AS Harga, SUM([Inventory Tabel].QTY_IN) AS QTY FROM Inventory LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE     ([Inventory Tabel].LockFIFO = 0) GROUP BY Inventory.NoItem, Inventory.ItemName, Inventory.PPn, Inventory.Markup, Inventory.UOM HAVING      (SUM([Inventory Tabel].QTY_IN) <> 0)", CNN, lckLockReadOnly
            DGPurchase.Columns(6).Visible = False
            DGPurchase.Columns(7).Visible = True
            
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1:
            mCall.FromTagActive = "Data Order Penjualan"
            mCall.txtCari = txtBox(3)
          Case 2:
            mCall.FromTagActive = "MASTER BARANG"
            mCall.txtCari = txtBox(2)
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Surat Jalan(Delivery Note) Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetailPart(ByVal ParamKocok As String)
On Error GoTo 8
Dim RcKu As New Recordset
RcKu.CursorLocation = adUseClient
RcKu.Open "SELECT [PO Order].PurchaseID AS [PO Number], [PO Order].PartnerID AS [Partner ID], [PO Order].DatePurchase AS [Tgl. PO], PartnerDB.Address AS Alamat, PartnerDB.City AS Kota , [PO Order].Kurs, [PO Order].TermPayment as Term,currid ,PartnerDB.CompanyName FROM [PO Order] INNER JOIN PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE   ([PO Order].PurchaseID =N'" & ParamKocok & "') ORDER BY [PO Order].PurchaseID", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcKu
     If .Recordcount <> 0 Then
        lblSupplier(0) = .Fields(1)
        lblSupplier(1) = .Fields(3)
        lblSupplier(2) = .Fields(4)
        lblSupplier(3) = Format(.Fields(2), "dd mmmm yyyy")
        lblSupplier(4) = FormatNumber(.Fields(5), 0)
        lblSupplier(5) = FormatNumber(.Fields(6), 0)
        lblSupplier(7) = .Fields("Currid")
        lblSupplier(8) = .Fields(1)
        lblSupplier(10) = .Fields(8)
     Else
        lblSupplier(0) = ""
        lblSupplier(1) = ""
        lblSupplier(2) = ""
        lblSupplier(3) = Format(Date, "dd mmmm yyyy")
        lblSupplier(4) = "0"
        lblSupplier(5) = "0"
        lblSupplier(7) = "IDR"
        lblSupplier(8) = ""
        lblSupplier(10) = ""
     End If
End With
CloseDB RcKu
Exit Sub
8:
MessageBox Err.Description, "frmartrans_OPENdetailpart & Err.Number, msgOkOnly, msgExclamation"
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
On Error GoTo 7
Dim rs As New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
rs.DBOpen " SELECT   [Detail TransData].NoItem, Inventory.ItemName, [Detail TransData].QTY_Receive AS QTY_OUT, [Detail TransData].Price, [Detail TransData].VAT,  [Detail TransData].Price * [Detail TransData].QTY_Receive - [Detail TransData].Price * [Detail TransData].QTY_Receive * ROUND(TransData.Discount / 100, 2)  + ([Detail TransData].Price * [Detail TransData].QTY_Receive - [Detail TransData].Price * [Detail TransData].QTY_Receive * ROUND(TransData.Discount / 100, 2)) * ROUND([Detail TransData].VAT / 100, 2) AS TOTALA, TransData.Discount" & _
          " FROM         [Detail TransData] INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem WHERE     (TransData.TransID = N'" & ParameterString & "') ORDER BY [Detail TransData].NoItem", CNN, lckLockBatch
Set MyDDE.ChildRecordset = rs.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
DGPurchase.Columns(5).Visible = True
DGPurchase.Columns(6).Visible = False
Exit Sub
7:
MessageBox Err.Description, "frmartrans_OPENDETAIL" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_Change(Index As Integer)
RefresDB Index
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
RefresDB Index
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub RefresDB(ByVal Index As Integer)
If MEdit = True Then
   If txtBox(Index).DataField <> "" And txtBox(Index).Tag <> "" Then
      MyDDE.GetFieldByName(txtBox(Index).DataField) = txtBox(Index)
   End If
End If
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "AR-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub HitungTotal()
On Error GoTo 5
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mTotal As Variant
Dim mPPn As Variant
Dim mDisc As Variant
Dim mStDisc As Variant
Dim mTmpDisc As Byte
Dim I As Integer
mTotal = 0
mDisc = 0
mPPn = 0
mStDisc = 0
LblAmount(0) = 0
LblAmount(1) = 0
LblAmount(2) = 0
LblAmount(3) = 0
mTmpDisc = IIf(Not IsNull(MyDDE.GetFieldByName("Discount")), MyDDE.GetFieldByName("Discount"), 0)
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcTotal
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        ' 2 = QTY  3 = Harga 4 = Vat
        For I = 0 To UBound(Avdata, 2)
            If mTmpDisc > 0 Then
               mDisc = mDisc + (Avdata(2, I) * Avdata(3, I)) * (mTmpDisc / 100)
               mStDisc = mStDisc + ((Avdata(2, I) * Avdata(3, I)) - ((Avdata(2, I) * Avdata(3, I)) * (mTmpDisc / 100)))
            Else
               mStDisc = mStDisc + (Avdata(2, I) * Avdata(3, I))
               mDisc = mDisc + 0
            End If
            If Avdata(4, I) > 0 Then
               mPPn = mPPn + ((((Avdata(2, I) * Avdata(3, I)) - ((Avdata(2, I) * Avdata(3, I)) * (mTmpDisc / 100))) * (Avdata(4, I) / 100)))
            Else
               mPPn = mPPn + 0
            End If
            mTotal = mTotal + Avdata(2, I) * Avdata(3, I)
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
Set mTotal = Nothing
Set mPPn = Nothing
Set mDisc = Nothing
Exit Sub
5:
MessageBox Err.Description, "frmartrans_hitungtotal" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Function BukaQTY(ByVal NoItem As String) As Long
On Error GoTo 1
Dim RcBuka As New Recordset
RcBuka.CursorLocation = adUseClient
RcBuka.Open "SELECT SUM([Inventory Tabel].StockTmp) - Inventory.MinStock AS QTY FROM [Inventory Tabel] INNER JOIN Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY Inventory.MinStock, LEFT([Inventory Tabel].RefTrans, 2), [Inventory Tabel].NoItem HAVING (LEFT([Inventory Tabel].RefTrans, 2) = N'RN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcBuka
     If .Recordcount <> 0 Then
        BukaQTY = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        If BukaQTY < 0 Then BukaQTY = 0
     Else
        BukaQTY = 0
     End If
     .Close
End With
Set RcBuka = Nothing
Exit Function
1:
MessageBox Err.Description, "frmartrans_bukaqty" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
     .PrepareAppend = " INSERT INTO TransData  (TransID,Discount,Kurs,CurrID, PurchaseID, DNID, DateTrans, DateIssued, RefNotes,PartnerID,TypeTrans,status)" & _
                      " VALUES (N'" & txtBox(0) & "'," & CDbl(lblSupplier(9)) & "," & MyDDE.GetFieldByName("Kurs") & ",N'" & MyDDE.GetFieldByName("Mata Uang") & "', N'" & lblSupplier(6) & "',N'" & txtBox(3) & "',  CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), getdate(), N'" & txtBox(2) & "','" & lblSupplier(0) & "','AR',1 )"
                      
     .PrepareUpdate = " UPDATE  TransData" & _
                      " Set Discount =" & CDbl(lblSupplier(9)) & ", Kurs = " & MyDDE.GetFieldByName("Kurs") & ",CurrID = N'" & MyDDE.GetFieldByName("Mata Uang") & "',DNID = N '" & lblSupplier(6) & "',PurchaseID = N '" & txtBox(3) & "',  DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), DateIssued = GETDATE(),PartnerID=N'" & lblSupplier(0) & "',TypeTrans='AR', RefNotes = N'" & txtBox(2) & "'" & _
                      " WHERE (TransID = N'" & txtBox(0) & "') AND (Status = 0)"

     .PrepareDelete = " DELETE FROM  TransData WHERE (TransID = N'" & txtBox(0) & "')"
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Function CekStock(ByVal NoItem As String) As Long
On Error GoTo 3
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
RcCek.Open "SELECT SUM(StockTmp) AS QTY FROM         [Inventory Tabel] GROUP BY NoItem, LEFT(RefTrans, 2) HAVING      (NoItem = N'" & NoItem & "') AND (LEFT(RefTrans, 2) = N'RN')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
     .Close
End With
Set RcCek = Nothing
Exit Function
3:
MessageBox Err.Description, "frmartrans_cekstock" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub TutupInvoice()
SendDataToServer ("UPDATE TransData SET Status = 1 WHERE     (TransID = N'" & txtBox(3) & "')")
SendDataToServer ("UPDATE [po Order] SET StatusSJ = 1 WHERE     (PurchaseID = N'" & lblSupplier(6) & "')")
End Sub

Private Sub SimpanDetail()
On Error GoTo xErr
Dim MyJournal            As New clsJournal
Dim mVarTunai, StrPartic As String

     If MyDDE.ChildRecordset.Recordcount <> 0 Then
        'mVarTunai = "BKMPTP" mVarTunai = "BPJK"
'        StrPartic = "Penjualan Persediaan Ke "
'        If MyJournal.CiptaKaryaHeaderJournal("", txtBox(0), txtBox(3), "", "", lblSupplier(0), lblSupplier(7), DTPicker1.Value, mVarPeriode, "BPJK") = True Then
'           'Piutang Dagang DR
'           MyJournal.CiptaKaryaDetailJournal "", CariTypeJournal(39), lblSupplier(0), CDbl(LblAmount(3)), 0
'           StrPartic = StrPartic & "," & lblSupplier(0) & " "
'           'Discount DR
'           MyJournal.CiptaKaryaDetailJournal "", CariTypeJournal(40), "xxx", CDbl(LblAmount(2)), 0
'           StrPartic = StrPartic & "," & "Diskon "
'           'Penjualan CR
'           MyJournal.CiptaKaryaDetailJournal "", CariTypeJournal(64), txtBox(0), 0, CDbl(LblAmount(0))
'           StrPartic = StrPartic & "," & "Penjualan "
'           'PPN KELUARAN CR
'           MyJournal.CiptaKaryaDetailJournal "", CariTypeJournal(41), "xxx", 0, CDbl(LblAmount(1))
'           StrPartic = StrPartic & "," & "PPN Keluaran "
           MyDDE.ChildRecordset.MoveFirst
           Do
              If MyDDE.ChildRecordset.EOF = True Then Exit Do
                 'FIFOBegun mydde.childrecordset.Fields("Noitem")
                  SendDataToServer " INSERT INTO [Detail TransData](TransID, NoItem, QTY_OUT, QTY_Receive, Price,Vat,Hpp)" & _
                                   " VALUES (N'" & txtBox(0) & "', N'" & MyDDE.ChildRecordset.Fields("NoItem") & "', " & MyDDE.ChildRecordset.Fields("QTY_OUT") & ",0, " & MyDDE.ChildRecordset.Fields("Price") & "," & MyDDE.ChildRecordset.Fields("VAT") & "," & HppProce(lblSupplier(0), MyDDE.ChildRecordset.Fields("NoItem")) & ")"
                  SendARItem MyDDE.ChildRecordset.Fields("NoItem"), CDbl(MyDDE.ChildRecordset.Fields("QTY_OUT")), CDbl(MyDDE.ChildRecordset.Fields("Price")), txtBox(3), DTPicker1.Value, HppProce(lblSupplier(6), MyDDE.ChildRecordset.Fields("NoItem")), "AR"
'                  'HPP DR
'                   MyJournal.CiptaKaryaDetailJournal "", CariTypeJournal(23), MyDDE.ChildRecordset.Fields("NoItem"), CDbl(MyDDE.ChildRecordset.Fields("QTY_OUT")) * HppProce(lblSupplier(6), MyDDE.ChildRecordset.Fields("NoItem")), 0
'                   StrPartic = StrPartic & "," & "Hpp "
'                  'Persediaan CR
'                   MyJournal.CiptaKaryaDetailJournal "", CariAkunItem(MyDDE.ChildRecordset.Fields("NoItem")), MyDDE.ChildRecordset.Fields("NoItem"), 0, CDbl(MyDDE.ChildRecordset.Fields("QTY_OUT")) * HppProce(lblSupplier(6), MyDDE.ChildRecordset.Fields("NoItem"))
'                   StrPartic = StrPartic & "," & "Persediaan " & MyDDE.ChildRecordset.Fields("NoItem")
                  MyDDE.ChildRecordset.MoveNext
           Loop
           MyDDE.ChildRecordset.MoveLast
'           MyJournal.CreateRefNotes StrPartic
           SendVoucher lblSupplier(6), lblSupplier(0), txtBox(2), DTPicker1.Value, CDbl(LblAmount(3)), 0, lblSupplier(6), "AR"
           TutupInvoice
'        End If
     End If

Set MyJournal = Nothing
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub FIFOBegun(ByVal NoItem As String)
On Error GoTo 4
Dim RcFifo As New Recordset
Dim mSisa As Long
On Error GoTo xErr
RcFifo.CursorLocation = adUseClient
RcFifo.Open " SELECT StockTmp FROM [Inventory Tabel] WHERE (NoItem = N'" & NoItem & "') AND (LockFIFO = 0) AND (LEFT(RefTrans, 2) = N'RN') ORDER BY NoIdx, DateIssued", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcFifo
     If .Recordcount <> 0 Then
        mSisa = 0
        Do
          If .EOF Then Exit Do
             mSisa = .Fields(0) - MyDDE.ChildRecordset.Fields("QTY_OUT")
             Select Case mSisa
                    Case Is > 0:
                         SendDataToServer (" INSERT INTO [Inventory Tabel](NoIdx, NoItem, QTY_OUT, PriceOut, RefTrans, DateTrans, StockTmp,DateIssued)" & _
                                           " VALUES (NEWID(), N'" & NoItem & "', " & MyDDE.ChildRecordset.Fields("QTY_OUT") & ", " & MyDDE.ChildRecordset.Fields("Price") & ",'" & txtBox(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3)," & mSisa & ", CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3))")
                         SendDataToServer (" UPDATE [Inventory Tabel] SET  StockTmp =" & mSisa & " WHERE     (LockFIFO = 0) AND (NoItem = N'" & NoItem & "')")
                         Exit Do
                    Case Is = 0:
                         SendDataToServer (" INSERT INTO [Inventory Tabel](NoIdx, NoItem, QTY_OUT, PriceOut, RefTrans, DateTrans, StockTmp,DateIssued)" & _
                                           " VALUES (NEWID(), N'" & NoItem & "', " & MyDDE.ChildRecordset.Fields("QTY_OUT") & ", " & MyDDE.ChildRecordset.Fields("Price") & ",'" & txtBox(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3)," & MyDDE.ChildRecordset.Fields("QTY_OUT") & ", CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3))")
                         SendDataToServer (" UPDATE  [Inventory Tabel] SET  StockTmp =" & mSisa & ",LockFifo=1 WHERE  (LockFIFO = 0) AND (NoItem = N'" & NoItem & "')")
                    Case Is < 0:
                         SendDataToServer (" INSERT INTO [Inventory Tabel](NoIdx, NoItem, QTY_OUT, PriceOut, RefTrans, DateTrans, StockTmp,DateIssued)" & _
                                           " VALUES (NEWID(), N'" & NoItem & "', " & MyDDE.ChildRecordset.Fields("QTY_OUT") & ", " & MyDDE.ChildRecordset.Fields("Price") & ",'" & txtBox(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3)," & MyDDE.ChildRecordset.Fields("QTY_OUT") & ", CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3))")
                    
                         SendDataToServer (" UPDATE    [Inventory Tabel] SET  StockTmp =" & mSisa & ",LockFifo = 1 WHERE  (LockFIFO = 1) AND (NoItem = N'" & NoItem & "')")
             End Select
          .MoveNext
        Loop
     Else
     End If
     .Close
End With
Set RcFifo = Nothing
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
Exit Sub
4:
MessageBox Err.Description, "frmartrans_fifobegun" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Function HppProce(ByVal NoPurchaseID As String, ByVal NoItem As String) As Double
On Error GoTo 6
Dim RcHpp As New DBQuick
RcHpp.DBOpen "SELECT     HPP FROM         [Detail PO] GROUP BY PurchaseID, NoItem, HPP HAVING      (PurchaseID = N'" & NoPurchaseID & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcHpp
     If .Recordcount <> 0 Then
        HppProce = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        HppProce = 0
     End If
End With
Exit Function
6:
MessageBox Err.Description, "frmartrans_hppproce" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function CariAkunItem(ByVal NoItem As String) As String
On Error GoTo 2
Dim Rc As DBQuick
Set Rc = New DBQuick
Rc.DBOpen "SELECT     NoAccount FROM         Inventory WHERE     (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
CariAkunItem = ""
With Rc
     If .Recordcount <> 0 Then
        CariAkunItem = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
Exit Function
2:
MessageBox Err.Description, "frmartrans_cariakunitem" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub GridLayout()
DGPurchase.Columns(0).width = 2025.071
DGPurchase.Columns(1).width = 2564.788
DGPurchase.Columns(2).width = 794.8347
DGPurchase.Columns(3).width = 1860.095
DGPurchase.Columns(4).width = 569.7638
DGPurchase.Columns(5).width = 1514.835
DGPurchase.Columns(6).width = 1514.835
DGPurchase.Columns(7).width = 1514.835
End Sub
