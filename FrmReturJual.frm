VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmReturJual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retur Jual"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmReturJual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Tag             =   "Sales Return"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   10980
      TabIndex        =   6
      Top             =   0
      Width           =   10980
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ReturID"
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
         Left            =   1335
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   240
         Width           =   3345
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TransID"
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
         Index           =   1
         Left            =   1335
         TabIndex        =   7
         Tag             =   "Partner"
         Top             =   915
         Width           =   2895
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
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
         Left            =   1350
         MaxLength       =   200
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   5040
         Width           =   4425
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4230
         Picture         =   "FrmReturJual.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   915
         Width           =   405
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   570
         Width           =   3345
         _ExtentX        =   5900
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
         Format          =   139329539
         CurrentDate     =   38272
      End
      Begin MSDataListLib.DataCombo cboGudang 
         DataField       =   "WareHouse"
         Height          =   330
         Left            =   1335
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1995
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nama Gudang"
         BoundColumn     =   "WareHouse"
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
         Height          =   2535
         Left            =   135
         TabIndex        =   8
         Top             =   2385
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4471
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
         ColumnCount     =   7
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
            DataField       =   "UOM"
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
         BeginProperty Column03 
            DataField       =   "QTYPO"
            Caption         =   "QTY Jual"
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
            DataField       =   "Retur Beli"
            Caption         =   "QTY Retur"
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
         BeginProperty Column06 
            DataField       =   "VAT"
            Caption         =   "Ppn"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang"
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
         Left            =   330
         TabIndex        =   27
         Top             =   2025
         Width           =   630
      End
      Begin VB.Label lblRN 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "DatePurchase"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
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
         Height          =   270
         Index           =   1
         Left            =   1335
         TabIndex        =   26
         Tag             =   "Partner"
         Top             =   1680
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. SC"
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
         Index           =   6
         Left            =   330
         TabIndex        =   25
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblRN 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         ForeColor       =   &H80000005&
         Height          =   270
         Index           =   0
         Left            =   1335
         TabIndex        =   24
         Tag             =   "Partner"
         Top             =   1305
         Width           =   3345
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "Phone"
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
         Height          =   225
         Index           =   4
         Left            =   5820
         TabIndex        =   23
         Tag             =   "Partner"
         Top             =   1380
         Width           =   3210
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "City"
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
         Height          =   225
         Index           =   3
         Left            =   5820
         TabIndex        =   22
         Tag             =   "Partner"
         Top             =   1140
         Width           =   3210
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         Height          =   225
         Index           =   2
         Left            =   5820
         TabIndex        =   21
         Tag             =   "Partner"
         Top             =   825
         Width           =   3210
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         Height          =   225
         Index           =   1
         Left            =   5820
         TabIndex        =   20
         Tag             =   "Partner"
         Top             =   555
         Width           =   3210
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "PartnerID"
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
         Height          =   225
         Index           =   0
         Left            =   5820
         TabIndex        =   19
         Tag             =   "Partner"
         Top             =   270
         Width           =   3210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Left            =   4890
         TabIndex        =   18
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO No."
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
         Left            =   330
         TabIndex        =   17
         Top             =   975
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order"
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
         Left            =   330
         TabIndex        =   16
         Top             =   1335
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Retur ID"
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
         Index           =   4
         Left            =   330
         TabIndex        =   15
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Index           =   5
         Left            =   330
         TabIndex        =   14
         Top             =   615
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   180
         TabIndex        =   13
         Top             =   5100
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   330
         X2              =   1575
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   315
         X2              =   1560
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   300
         X2              =   1545
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   315
         X2              =   1560
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   150
         X2              =   1395
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Label LblAmount 
         Caption         =   "Label2"
         Height          =   300
         Index           =   0
         Left            =   7845
         TabIndex        =   12
         Top             =   3210
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label LblAmount 
         Caption         =   "Label2"
         Height          =   300
         Index           =   1
         Left            =   7845
         TabIndex        =   11
         Top             =   3540
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label LblAmount 
         Caption         =   "Label2"
         Height          =   300
         Index           =   2
         Left            =   7845
         TabIndex        =   10
         Top             =   3870
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label LblAmount 
         Caption         =   "Label2"
         Height          =   300
         Index           =   3
         Left            =   7845
         TabIndex        =   9
         Top             =   4200
         Visible         =   0   'False
         Width           =   2130
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5700
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1005
      BindFormTAG     =   "RN"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmReturJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall                 As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner                        As New DBQuick
Private RcDetail                         As New DBQuick
Private MyData                           As New clsTransaksi
Private MEdit, mFirstCaller              As Boolean
Private RcGudang                         As New DBQuick
Private pWhere As String
Dim SQLInit As String

Public Property Let IDParams(vData As String)
   pWhere = vData
   
End Property

Private Sub cboGudang_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim mStok As Variant
Select Case ColIndex
       Case 4:
            If TotalRetur(MyDDE.ChildRecordset.Fields("NoItem")) = True And CDbl(DGPurchase.Columns(ColIndex).Value) <= CDbl(DGPurchase.Columns(3).Value) Then
                mStok = CekStock(MyDDE.ChildRecordset.Fields("NoItem")) - MyDDE.ChildRecordset.Fields("Retur Beli")
                If mStok < 0 Then
                   MessageBox "Stock Tidak Cukup Untuk Melakukan Transaksi." & vbCrLf & "Stok Kurang -> " & mStok & " Untuk Memenuhi Transaksi SC", "Peringatan", msgOkOnly
                   MyDDE.ChildRecordset.Fields("Retur Beli") = 0
                End If
             Else
                MessageBox "QTY Retur tidak boleh melebihi QTY Jual.", "Peringatan", msgOkOnly
                MyDDE.ChildRecordset.Fields("Retur Beli") = 0
             End If
             TotalTrans
End Select
'HitungTotal
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If MEdit = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
'If Shift = 2 And KeyCode = vbKeyF3 Then
'    DGPurchase.Columns(3) = 0
'    DGPurchase.Columns(4) = 0
'    DGPurchase.Columns(5) = 0
'    DGPurchase.Columns(6) = 0
'    OpenPartner 1
'ElseIf Shift = 2 And KeyCode = vbKeyF2 Then
'   If MyDDE.CheckEmptyControl = False Then
'     OpenPartner 1
'     DGPurchase.SetFocus
'   Else
'      MessageBox "Data Transaksi Belum Ada.Harap Diisi Dulu.", "Peringatan"
'   End If
'End If
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
   Exit Sub
End If
With DGPurchase
     Select Case .col
            Case 0, 1, 2, 3, 5, 6:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = False
            Case Else:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
     End Select
End With
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
RcGudang.DBOpen " SELECT     WareHouse , [WareHouse Name] as [Nama Gudang] FROM  WareHouse ", CNN, lckLockReadOnly
Set cboGudang.RowSource = RcGudang.DBRecordset
SQLInit = " SELECT ReturData.ReturID, TransData.TransID, ReturData.DateTrans, ReturData.DateIssued, ReturData.RefNotes, ReturData.WareHouse,                       [PO Order].PurchaseID, [PO Order].PartnerID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, PartnerDB.Phone, " & _
          " [PO Order].DatePurchase, [PO Order].Discount FROM         ReturData INNER JOIN                       TransData ON ReturData.TransID = TransData.TransID INNER JOIN                       [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN                       PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE "
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmReturJual
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    If Trim(pWhere) = "" Then
      .PrepareQuery = SQLInit & " (ReturData.TypeTrans = N'RJ') ORDER BY ReturData.ReturID"
    Else
      .PrepareQuery = SQLInit & " ReturData.ReturID='" & pWhere & "'"
    End If
    .SetPermissions = aksess.MayDo("Penerimaan Retur Customer")
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set mCall = Nothing
RcPartner.CloseDB
RcDetail.CloseDB
RcGudang.CloseDB
Set MyData = Nothing
End Sub

Private Sub Form_Resize()
  Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set FrmReturJual = Nothing
   pWhere = ""
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT TransData.DNID AS [NO RN], [PO Order].PurchaseID AS [PO Order], TransData.DateTrans AS [Tgl Bukti], [PO Order].PartnerID AS [Kode Customer], " & _
                             " PartnerDB.CompanyName AS Perusahaan, PartnerDB.Address AS Alamat, PartnerDB.City AS Kota, PartnerDB.Phone AS Telepon, [PO Order].Discount FROM         TransData INNER JOIN                       [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN                       PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE     (TransData.TypeTrans = N'AR') ORDER BY TransData.DNID", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen "SELECT Inventory.NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit, [Detail TransData].QTY_Receive AS [QTY Beli],                        [Detail TransData].Price AS Harga, [Detail TransData].VAT AS Ppn FROM         [Detail TransData] INNER JOIN                       Inventory ON [Detail TransData].NoItem = Inventory.NoItem WHERE     ([Detail TransData].TransID = N'" & txtBox(1) & "')", CNN, lckLockBatch
            mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "MASTER CUSTOMER"
           Case 1: mCall.FromTagActive = "DETAIL PEMBELIAN":
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.FromTagActive
       Case "DETAIL PEMBELIAN":
            If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            End If
            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
            mFirstCaller = False
       Case "MASTER CUSTOMER": If cboGudang.Enabled = True Then cboGudang.SetFocus
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm:
       Case "MASTER CUSTOMER":
            With MyDDE
                 .GetFieldByName("TransID") = mCall.GetFieldByName(0)
                 .GetFieldByName("PurchaseID") = mCall.GetFieldByName(1)
                 .GetFieldByName("DatePurchase") = mCall.GetFieldByName(2)
                 .GetFieldByName("PartnerID") = mCall.GetFieldByName(3)
                 .GetFieldByName("CompanyName") = mCall.GetFieldByName(4)
                 .GetFieldByName("Address") = mCall.GetFieldByName(5)
                 .GetFieldByName("City") = mCall.GetFieldByName(6)
                 .GetFieldByName("Phone") = mCall.GetFieldByName(7)
                 .GetFieldByName("Discount") = mCall.GetFieldByName("Discount")
            End With
       Case "DETAIL PEMBELIAN":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = mCall.GetFieldByName(2)
                 .Fields(3) = mCall.GetFieldByName(3)
                 .Fields(5) = mCall.GetFieldByName(4)
                 .Fields(6) = mCall.GetFieldByName(5)
                 .Fields("Retur Beli") = 0
            End With
       
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            cmdLink(0).Enabled = False
            DTPicker1.SetFocus
       Case tmbAddNew:
            Dim IDGen As New IDGenerator
            DTPicker1.Value = CDate(Format(dDateBegin, "dd/mm/yy"))
            MEdit = True
            txtBox(0).Enabled = False
            cmdLink(0).Enabled = True
            MyDDE.GetFieldByName("ReturID") = IDGen.GetID("SR")    'MyData.PrepareIndex(tmbTransaksiReturjual, 5, "1", TglIndex)
            MyDDE.GetFieldByName("RefNotes") = "-"
            DTPicker1.SetFocus
       Case tmbCancel:
            'mEdit = True
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               cmdLink(0).Enabled = True
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               MEdit = False
            End If
       Case tmbDetail:
            If mFirstCaller = False Then
               OpenPartner 1
               MEdit = True
            End If
       Case tmbPrint:
            CallRPTReport "RETUR JUAL.RPT", "SELECT * FROM [RETUR JUAL] WHERE [No Retur] =N'" & txtBox(0) & "'"
End Select
txtBox(0).Enabled = False
txtBox(1).Enabled = False
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("ReturID")
TotalTrans
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
       Case tmbAddNew:
       Case tmbDetail: MyDDE.CancelTrans = mFirstCaller
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If CekGridKosong = False Then
                     MyDDE.IsChildMemberReady = True
                  Else
                     MyDDE.IsChildMemberReady = False
                  End If
               Else
                  MessageBox "Data detail belum ada. Harap diisi dulu.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "RJ/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub OpenDetail(ByVal ParameterString As String)
If ParameterString = "" Then ParameterString = "xxxxxxxx"
RcDetail.DBOpen " SELECT     [Detail Retur].NoItem, Inventory.ItemName, Inventory.UOM, [Detail Retur].[QTY BJ] AS QTYPO, [Detail Retur].[Retur Jual] AS [Retur Beli], [Detail Retur].Price,   [Detail Retur].VAT FROM         [Detail Retur] INNER JOIN Inventory ON [Detail Retur].NoItem = Inventory.NoItem INNER JOIN ReturData ON [Detail Retur].ReturID = ReturData.ReturID WHERE     (ReturData.ReturID = N'" & ParameterString & "') ORDER BY [Detail Retur].NoItem", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) <> N'DN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
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
With MyDDE
    .PrepareAppend = " INSERT INTO ReturData" & _
                     " (ReturID, TransID, DateTrans, RefNotes,TypeTrans,WareHouse)" & _
                     " VALUES     (N'" & txtBox(0) & "', N'" & txtBox(1) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & ValidString(txtBox(2)) & "',N'RJ',N'" & cboGudang.BoundText & "')"
                     
    .PrepareUpdate = " UPDATE    ReturData" & _
                     " Set Warehouse=N'" & cboGudang.BoundText & "', TransID = N'" & txtBox(1) & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), RefNotes = N'" & ValidString(txtBox(2)) & "' WHERE     (ReturID = N'" & txtBox(0) & "')"
                     
    .PrepareDelete = " DELETE FROM  [ReturData] WHERE (ReturID = N'" & txtBox(0) & "')"
End With
Err.Clear
End Sub

Private Sub SimpanDetail()
Dim MyJournal As New clsJournal
Dim StrPartic As String
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [Detail Retur] WHERE     (ReturID = N'" & txtBox(0) & "')") = True Then
'              If SendDataToServer("DELETE From [Table Journal] where TransID =N'" & txtBox(0) & "' and TypeTrans=N'BRPJ'") = True Then
                 'If MyJournal.CiptaKaryaHeaderJournal("", txtBox(0), txtBox(1), "", "", lblSupplier(0), "IDR", DTPicker1.Value, mVarPeriode, "BRPJ") = True Then
                    StrPartic = "Retur Pembelian "
                    Do
                       If .EOF = True Then Exit Do
                          SendDataToServer " INSERT INTO [Detail Retur]" & _
                                           " (ReturID, NoItem, [QTY BJ],[Retur Jual],  Price,hpp, VAT)" & _
                                           " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "'," & .Fields("QTYPO") & ", " & .Fields("Retur Beli") & ",  " & .Fields("Price") & "," & HppProce(lblRN(0), .Fields("NoItem")) & ", " & .Fields("Vat") & ")"
                          'SendDataToServer (" DELETE FROM  [Inventory Tabel] WHERE (RefTrans = N'" & txtBox(0) & "') and (Noitem=N'" & .Fields("NoItem") & "')")
                          SendARItem .Fields("NoItem"), CCur(.Fields("Retur Beli")), CDbl(.Fields("Price")), txtBox(0), DTPicker1.Value, HppProce(lblRN(0), .Fields("NoItem")), "RJ"
                          SendDataToServer ("Update [Detail PO] Set QTYRetur =" & TotalReturbeli(.Fields("NoItem")) & " where Purchaseid=N'" & lblRN(0) & "' and NoItem=N'" & .Fields("NoItem") & "'")
'                          'Persediaan
'                          MyJournal.CiptaKaryaDetailJournal "", CariAkunItem(.Fields("NoItem")), .Fields("NoItem"), .Fields("Retur Beli") * HppProce(lblRN(0), .Fields("NoItem")), 0
'                          StrPartic = StrPartic & "," & .Fields("NoItem")
'                          'Hpp
'                          MyJournal.CiptaKaryaDetailJournal "", CariTypeAccount(23), .Fields("NoItem"), 0, .Fields("Retur Beli") * HppProce(lblRN(0), .Fields("NoItem"))
'                          StrPartic = StrPartic & ", Hpp " & .Fields("NoItem")
                          .MoveNext
                    Loop
                    .MoveLast
'                    'Retur Penjualan
'                    MyJournal.CiptaKaryaDetailJournal "", CariTypeAccount(43), txtBox(0), CDbl(LblAmount(3)), 0
'                    StrPartic = StrPartic & ", Retur Penj. " & txtBox(0)
'                    'Piutang Usaha
'                    MyJournal.CiptaKaryaDetailJournal "", CariTypeAccount(39), lblSupplier(0), 0, CDbl(LblAmount(3))
'                    StrPartic = StrPartic & ", Piut. Usaha " & lblSupplier(0)
'                    'Diskon
'                    MyJournal.CiptaKaryaDetailJournal "", CariTypeAccount(40), .Fields("NoItem"), 0, CDbl(LblAmount(2))
'                    StrPartic = StrPartic & ", Diskon"
'                    'PPn Keluaran
'                    MyJournal.CiptaKaryaDetailJournal "", CariTypeAccount(41), .Fields("NoItem"), CDbl(LblAmount(1)), 0
'                    StrPartic = StrPartic & ", Ppn Keluaran"
'                 End If
'              End If
              
           End If
           .MoveLast
           MyJournal.CreateRefNotes StrPartic
           'DGPurchase.Refresh
     End If
End With
Set MyJournal = Nothing
End Sub

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
               MessageBox "Data item untuk QTY Beli atau QTY Retur ada yang berisi NOL", "Peringatan", msgOkOnly
               CekGridKosong = True
               MyDDE.CancelTrans = True
               Exit For
            End If
        Next I
     Else
        CekGridKosong = True
     End If
End With
RcKsg.CloseDB
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Function HppProce(ByVal NoPurchaseID As String, ByVal NoItem As String) As Double
Dim RcHpp As New DBQuick
RcHpp.DBOpen "SELECT     HPP FROM         [Detail PO] GROUP BY PurchaseID, NoItem, HPP HAVING      (PurchaseID = N'" & NoPurchaseID & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcHpp
     If .Recordcount <> 0 Then
        HppProce = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        HppProce = 0
     End If
End With
RcHpp.CloseDB
End Function

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GLAccount.NoAccount, AccType.ID, GLAccount.AccountName FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

Private Sub TotalTrans()
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Dim mDis As Integer
Set Rc.DBRecordset = MyDDE.ChildRecordset.Clone(adLockBatchOptimistic)
LblAmount(0) = 0
LblAmount(1) = 0
LblAmount(2) = 0
LblAmount(3) = 0
mDis = IIf(Not IsNull(MyDDE.GetFieldByName("DISCOUNT")), MyDDE.GetFieldByName("DISCOUNT"), 0)
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        ' 4 = QTY Retur 5 = Harga 6 = PPn
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            LblAmount(0) = FormatNumber(LblAmount(0) + (Avdata(4, I) * Avdata(5, I)), 0)
            LblAmount(1) = FormatNumber(LblAmount(1) + (Avdata(5, I) * (Avdata(6, I) / 100)) * (Avdata(4, I)), 0)
            If mDis <> 0 Then
               LblAmount(2) = FormatNumber(CDbl(LblAmount(0)) * CDbl(mDis / 100), 0)
            Else
               LblAmount(2) = 0
            End If
            LblAmount(3) = FormatNumber((CDbl(LblAmount(0)) + CDbl(LblAmount(1)) - CDbl(LblAmount(2))), 0)
        Next I
     End If
End With
Set Avdata = Nothing
End Sub

Private Function CariAkunItem(ByVal NoItem As String) As String
Dim Rc As DBQuick
Set Rc = New DBQuick
Rc.DBOpen "SELECT     NoAccount FROM         Inventory WHERE     (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
CariAkunItem = ""
With Rc
     If .Recordcount <> 0 Then
        CariAkunItem = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Function TotalRetur(ByVal NoItem As String) As Boolean
Dim RcRetur As New DBQuick
RcRetur.DBOpen " SELECT [Detail Retur].[QTY BJ] - SUM([Detail Retur].[Retur Jual]) AS [QTY Retur] FROM [Detail Retur] INNER JOIN ReturData ON [Detail Retur].ReturID = ReturData.ReturID WHERE     (ReturData.TransID = N'" & txtBox(1) & "') AND ([Detail Retur].NoItem = N'" & NoItem & "') GROUP BY [Detail Retur].[QTY BJ]", CNN, lckLockReadOnly
With RcRetur.DBRecordset
     If .Recordcount <> 0 Then
        If .Fields(0) = 0 Then
           TotalRetur = False
        Else
           TotalRetur = True
        End If
     Else
        TotalRetur = True
     End If
     .Close
End With
Set RcRetur = Nothing
End Function

Private Function TotalReturbeli(ByVal NoItem As String) As Long
Dim RcRetur As New DBQuick
RcRetur.DBOpen "SELECT      SUM([Detail Retur].[Retur Jual]) AS [QTY Retur] FROM         [Detail Retur] INNER JOIN ReturData ON [Detail Retur].ReturID = ReturData.ReturID WHERE     (ReturData.TransID = N'" & txtBox(1) & "') AND ([Detail Retur].NoItem = N'" & NoItem & "') GROUP BY [Detail Retur].[QTY BJ]", CNN, lckLockReadOnly
With RcRetur.DBRecordset
     If .Recordcount <> 0 Then
        TotalReturbeli = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        TotalReturbeli = 0
     End If
     .Close
End With
Set RcRetur = Nothing
End Function

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1950.236
DGPurchase.Columns(1).width = 3644.788
DGPurchase.Columns(2).width = 1514.835
DGPurchase.Columns(3).width = 1514.835
DGPurchase.Columns(4).width = 1514.835
DGPurchase.Columns(5).width = 1514.835
DGPurchase.Columns(6).width = 1514.835
End Sub
