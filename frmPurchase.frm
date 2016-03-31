VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPurchase 
   AutoRedraw      =   -1  'True
   Caption         =   "P.O Transaksi"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPurchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   10590
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   5010
      Left            =   60
      ScaleHeight     =   4950
      ScaleWidth      =   10275
      TabIndex        =   11
      Top             =   300
      Width           =   10335
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
         Left            =   6450
         MaxLength       =   5
         TabIndex        =   8
         Tag             =   "PO"
         Top             =   2610
         Width           =   1440
      End
      Begin VB.TextBox txtBox 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   6450
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1950
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         DataField       =   "PartnerID"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   6450
         TabIndex        =   1
         Tag             =   "PO"
         Top             =   930
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         DataField       =   "TermPayment"
         Height          =   315
         Index           =   1
         Left            =   1665
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "PO"
         Top             =   1950
         Width           =   690
      End
      Begin VB.TextBox txtBox 
         DataField       =   "PurchaseID"
         Height          =   315
         Index           =   0
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "PO"
         Top             =   930
         Width           =   3450
      End
      Begin VB.CheckBox chkPo 
         BackColor       =   &H00F2E2D5&
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
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1680
         TabIndex        =   7
         Top             =   2460
         Width           =   1995
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "frmPurchase.frx":030A
         Height          =   2265
         Left            =   90
         TabIndex        =   9
         Tag             =   "Partner"
         Top             =   3015
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   3995
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         BorderStyle     =   0
         HeadLines       =   1
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
            Caption         =   "Sup. Id Code"
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2610.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin SemeruDC.SemeruButton cmdLink 
         Height          =   315
         Index           =   1
         Left            =   9915
         TabIndex        =   5
         Top             =   1935
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPurchase.frx":031F
         PICN            =   "frmPurchase.frx":033B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SemeruDC.SemeruButton cmdLink 
         Height          =   315
         Index           =   0
         Left            =   9930
         TabIndex        =   2
         Top             =   930
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPurchase.frx":15BD
         PICN            =   "frmPurchase.frx":15D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DatePurchase"
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Tag             =   "PO"
         Top             =   1275
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
         Format          =   57999363
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   2
         Left            =   75
         TabIndex        =   25
         Top             =   1680
         Width           =   2790
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   8
         Left            =   5925
         TabIndex        =   24
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   7995
         TabIndex        =   23
         Top             =   5415
         Width           =   2445
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   7
         Left            =   7500
         TabIndex        =   22
         Top             =   5430
         Width           =   525
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   2
         Left            =   6450
         TabIndex        =   21
         Top             =   2325
         Width           =   555
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner Name"
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   1
         Left            =   6450
         TabIndex        =   20
         Top             =   1665
         Width           =   1125
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   0
         Left            =   6450
         TabIndex        =   19
         Top             =   1320
         Width           =   840
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   6
         Left            =   5520
         TabIndex        =   18
         Top             =   2310
         Width           =   870
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   5
         Left            =   5130
         TabIndex        =   17
         Top             =   2010
         Width           =   1260
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   4
         Left            =   5370
         TabIndex        =   16
         Top             =   975
         Width           =   1020
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   3
         Left            =   540
         TabIndex        =   15
         Top             =   2010
         Width           =   2325
      End
      Begin VB.Label LblDeliVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   1665
         TabIndex        =   14
         Top             =   1695
         Width           =   105
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   1
         Left            =   900
         TabIndex        =   13
         Top             =   1320
         Width           =   705
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   0
         Left            =   915
         TabIndex        =   12
         Top             =   975
         Width           =   705
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   90
         Picture         =   "frmPurchase.frx":285B
         Stretch         =   -1  'True
         Top             =   60
         Width           =   960
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   10
      Top             =   5400
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   820
      BindFormTAG     =   "Partner"
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents RcDetail As Recordset
Attribute RcDetail.VB_VarHelpID = -1
Dim RcPartner As Recordset
Dim RcRemind As New Recordset
Dim MyData As New clsTransaksi
Dim mEdit, mEditPO As Boolean
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim mAccount As String

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

'Private Sub Form_Activate()
'Dim i As Long
'lblCaption = Me.Caption
'If mEdit = False Then
'   i = MyDDE.SavedPointer
'   If Me.Caption = "P.O Transaksi" Then MyData.PrepareTransaksi tmbTransaksiPO Else MyData.PrepareTransaksi tmbTransaksiSC
'   If i <> 0 Then MyDDE.SavedPointer = i
'End If
'If Me.Caption <> "P.O Transaksi" Then
'   Label1(0) = "S.C. ID:"
'   Label1(1) = "Tgl. SC:"
'End If
'End Sub
'
'Private Sub Form_Initialize()
'
'End Sub

Private Sub Form_Load()
DTPicker1.Value = dDateBegin
'If Me.Caption = "P.O Transaksi" Then MyData.PrepareTransaksi tmbTransaksiPO Else MyData.PrepareTransaksi tmbTransaksiSC
Set mCall = New frmCaller
'MyDDE.EditModeReplace = False
MyDDE.SetPermissions = UserDeleteDenied
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
If Me.Caption = "P.O Transaksi" Then
   IsFrmPo = False
Else
   IsfrmSc = False
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
SweetForm Me
Me.Height = 6270
Me.Width = 10710
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPurchase = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
If pRecordset.RecordCount <> 0 Then
Select Case TagForm:
       Case "PARTNER", "CUSTOMER":
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
            lblSupplier(0) = mCall.GetFieldByName(1)
            lblSupplier(1) = mCall.GetFieldByName(2)
       Case "BANK":
            mAccount = mCall.GetFieldByName(0)
            txtBox(3) = mCall.GetFieldByName(1)
            lblSupplier(2) = mCall.GetFieldByName(2)
'       Case "ITEM":
'            RcDetail.Fields(0) = mCall.GetFieldByName(0)
'            RcDetail.Fields(1) = mCall.GetFieldByName(1)
'            RcDetail.Fields(2) = mCall.GetFieldByName(2)
       Case "ITEM", "ITEMSC", "REMINDER":
            RcDetail.Fields(0) = mCall.GetFieldByName(0)
            RcDetail.Fields(1) = mCall.GetFieldByName(1)
            RcDetail.Fields(2) = mCall.GetFieldByName(2)
            RcDetail.Fields(4) = mCall.GetFieldByName(5)
            RcDetail.Fields(5) = mCall.GetFieldByName(4)
            DGPurchase.Columns(7).Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
            If chkPo.Value = 0 Then
               RcDetail.Fields(3) = 0
            Else
               RcDetail.Fields(3) = mCall.GetFieldByName(3)
               If CDbl(DGPurchase.Columns(3).Value) <> 0 Then
                  RcDetail.Fields("tmp") = CDbl((DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4)))
               Else
                  DGPurchase.Columns("tmp").Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100)
               End If
            End If
End Select
End If
End Sub

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim i As Integer
Dim mStok As Long
Select Case ColIndex
       Case 3, 4, 5:
            If ColIndex = 3 And Me.Caption <> "P.O Transaksi" Then
               mStok = CekStock(RcDetail.Fields("NoItem")) - RcDetail.Fields("QtyPo")
               If mStok < 0 Then
                  i = MessageBox("Stock Tidak Cukup Untuk Melakukan Transaksi." & vbCrLf & "Stok Kurang -> " & mStok & " Untuk Memenuhi Transaksi SC" & vbCrLf & vbCrLf & "Tekan YES Untuk Tranfer Ke PO Reminder" & vbCrLf & "Tekan NO Untuk Membatalkan", "Peringatan", msgYesNo)
                  If i = 1 Then
                     MyDDE.SendDataToServer (" DELETE FROM [Remainder PO] WHERE     (NoItem = N'" & RcDetail.Fields("NoItem") & "') AND (SCNo = N'" & txtBox(0) & "') ")
                     MyDDE.SendDataToServer (" INSERT INTO [Remainder PO]  (Idx, NoItem, QTYOrder, SCNo)" & _
                                             " VALUES     (NEWID(), N'" & RcDetail.Fields("NoItem") & "', " & CDbl(Abs(mStok)) & ", N'" & txtBox(0) & "')")
                  
                  Else
                     DGPurchase.Columns(3).Value = 0
                  End If
               End If
            End If
            If RcDetail.Fields("qtyTemp") = 0 Then
               If CDbl(DGPurchase.Columns(ColIndex).Value) <> 0 Then
                  DGPurchase.Columns(7).Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
               Else
                  DGPurchase.Columns(7).Value = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100)
               End If
            Else
               MessageBox "Data Tidak Bisa Diedit Karena Digunakan Oleh Receive Notes Transaksi", "Peringatan", msgOkOnly
               RcDetail.CancelBatch adAffectCurrent
            End If
End Select
HitungTotal
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If mEdit = False Then Exit Sub

If Shift = 2 And KeyCode = vbKeyF3 Then
   mEdit = True
   RcDetail.AddNew
   DGPurchase.Columns(3) = 0
   DGPurchase.Columns(4) = 0
   DGPurchase.Columns(5) = 0
   DGPurchase.Columns(7) = 0
   RcDetail.Fields("QtyTemp") = 0
   DGPurchase.Columns(8) = Format(DTPicker1.Value + Val(txtBox(1)), "dd/mm/yyyy")
   If chkPo.Value = 1 Then
      OpenPartner 2
   Else
      OpenPartner 3
   End If
End If
If Shift = 2 And KeyCode = vbKeyF2 Then
   mEdit = True
   If chkPo.Value = 1 Then
      OpenPartner 2
   Else
      OpenPartner 3
   End If
End If

'ScanGrid DGPurchase
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
   Exit Sub
End If
With DGPurchase
     Select Case .Col
            Case 0, 1, 6:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = False
            Case Else:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
     End Select
End With
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            If IsHeaderOk(txtBox(0)) = False Then
               MyDDE.CancelTrans = False
            Else
               MyDDE.CancelTrans = True
               MessageBox "Transaksi PO Tidak Bisa Diedit...........!Sudah Dikunci Oleh RN Transaksi."
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If ScanGrid(DGPurchase) = False And RcDetail.RecordCount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                  If Me.Caption = "P.O Transaksi" Then
                     PrepareQuery
                  Else
                     MyData.PrepareQuery tmbTransaksiReceive
                  End If
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
'MyDDE.CancelTrans = False
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
txtBox(0).Enabled = False
txtBox(2).Enabled = False

Select Case AdReasonActiveDb
       Case tmbEdit:
            mEdit = True
            mEditPO = True
            cmdLink(0).Enabled = mEdit
            cmdLink(1).Enabled = mEdit
            If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = mEdit
       Case tmbAddNew:
            mEdit = True
            MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
            MyDDE.GetFieldByName("TermPayment") = 0
            MyDDE.GetFieldByName("Kurs") = 1
            If Me.Caption = "P.O Transaksi" Then
               MyDDE.GetFieldByName("PurchaseID") = MyData.PrepareIndex(tmbTransaksiPO, 5, "1", TglIndex)
            Else
               MyDDE.GetFieldByName("PurchaseID") = MyData.PrepareIndex(tmbTransaksiSC, 5, "1", TglIndex)
            End If
            DGPurchase.Columns(6).Visible = False
            DGPurchase.Columns(7).Visible = True
            DTPicker1.SetFocus
            cmdLink(0).Enabled = mEdit
            cmdLink(1).Enabled = mEdit
            If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = mEdit
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail mEditPO
               mEdit = False
               cmdLink(0).Enabled = mEdit
               cmdLink(1).Enabled = mEdit
               If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = mEdit
               mEditPO = False
            End If
       Case tmbCancel:
            mEdit = False
            cmdLink(0).Enabled = mEdit
            cmdLink(1).Enabled = mEdit
            If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = mEdit
       Case tmbPrint:
            Dim Mprint As New frmReportView
            With Mprint
                 If Me.Caption = "P.O Transaksi" Then
                    .QuerySource = "Select * from [Purchase Order] where PurchaseID='" & txtBox(0) & "'"
                    .ReportName = "Purchase Order.rpt"
                 Else
                    .QuerySource = "Select * from [Sales Contract] where PurchaseID='" & txtBox(0) & "'"
                    .ReportName = "Sales Contract.rpt"
                 End If
                 .Show
            End With
       Case tmbQuit:
            Unload Me
            Set MainMenu.DataDC.BindForm = Nothing
            
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("PurchaseID")
OpenDetailPart IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "")
HitungTotal
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
CloseDB RcPartner
Set RcPartner = New Recordset
RcPartner.CursorLocation = adUseClient
Set mCall = New frmCaller
Select Case Index
       Case 0:
            If Me.Caption = "P.O Transaksi" Then
               RcPartner.Open " SELECT PartnerID AS [Partner ID],CompanyName, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'SUPPLIER') ORDER BY PartnerID", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
               mCall.FromTagActive = "PARTNER"
            Else
               RcPartner.Open "SELECT PartnerID AS [Partner ID],CompanyName, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'CUSTOMER') ORDER BY PartnerID", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
               mCall.FromTagActive = "CUSTOMER"
            End If
            mCall.txtCari = txtBox(2)
       Case 1:
            RcPartner.Open "SELECT Account,[Bank Name], Currency, [Default] FROM [Bank Partner] WHERE (PartnerID = N'" & txtBox(2) & "') ORDER BY [Default], [Bank Name]", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
            mCall.FromTagActive = "BANK"
            mCall.txtCari = txtBox(3)
       Case 2:
            RcPartner.Open "SELECT [Remainder PO].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100)   + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
            mCall.FromTagActive = "REMINDER"
            mCall.txtCari = txtBox(3)
       Case 3:
            If Me.Caption = "P.O Transaksi" Then
               RcPartner.Open "SELECT NoItem, ItemName, [Serial Supplier], UOM, PPn, PriceIn * (Markup / 100) + PriceIn AS Harga FROM Inventory GROUP BY NoItem, ItemName, PPn, Markup, [Serial Supplier], UOM, PriceIn", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
               'RcPartner.Open "SELECT NoItem, ItemName, [Serial Supplier], Merk FROM Inventory ORDER BY NoItem", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
               mCall.FromTagActive = "ITEM"
            Else
               RcPartner.Open "SELECT NoItem, ItemName, [Serial Supplier], UOM, PPn, PriceIn * (Markup / 100) + PriceIn AS Harga FROM Inventory GROUP BY NoItem, ItemName, PPn, Markup, [Serial Supplier], UOM, PriceIn", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
               mCall.FromTagActive = "ITEMSC"
            End If
            mCall.txtCari = txtBox(2)
            DGPurchase.Columns(6).Visible = False
            DGPurchase.Columns(7).Visible = True
End Select
Set RcPartner.ActiveConnection = Nothing
If RcPartner.RecordCount <> 0 Then
    Set mCall.FormData = RcPartner
    mCall.Show vbModal
    Set mCall = Nothing
    If FindOwnRecordset(RcDetail, "NoItem = '" & RcDetail.Fields("NoItem") & "'") = True Then
       MessageBox "Record -> " & RcDetail.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
       RcDetail.CancelBatch adAffectCurrent
    End If
End If
    Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
CloseDB RcDetail
Set RcDetail = New Recordset
RcDetail.CursorLocation = adUseClient
If ParameterString = "" Then ParameterString = "xxxxxxxx"
With RcDetail
     .Open "SELECT [Detail PO].NoItem, Inventory.ItemName, [Detail PO].ItemSupplierID, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate, ( ([Detail PO].QTYPO * [Detail PO].POPrice) * ([Detail PO].VAT / 100)) + ([Detail PO].QTYPO * [Detail PO].POPrice) AS FldTotal, [Detail PO].POPrice AS TMP, [Detail PO].PurchaseID,[Detail PO].QTYTemp,[Detail PO].StatusTrans FROM [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem WHERE     ([Detail PO].PurchaseID = N'" & ParameterString & "') ORDER BY [Detail PO].NoItem", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
     Set .ActiveConnection = Nothing
     DGPurchase.Columns(6).Visible = True
     DGPurchase.Columns(7).Visible = False
     Set DGPurchase.DataSource = RcDetail
End With
End Sub

Private Sub OpenDetailPart(ByVal Param As String)
On Error Resume Next
Dim rcP As New Recordset
Dim rcD As New Recordset
rcP.CursorLocation = adUseClient
rcP.Open "Shape{SELECT  PartnerID, Address, CompanyName FROM PartnerDB WHERE (PartnerID = N'" & Param & "')} Append ({SELECT account,[Bank Name], Currency,PartnerID FROM [Bank Partner] WHERE (PartnerID = N'" & Param & "') AND ([Default] = 1)} as Cap relate PartnerID to PartnerID)", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
Set rcP.ActiveConnection = Nothing
With rcP
     If .RecordCount <> 0 Then
        lblSupplier(0) = .Fields(2)
        lblSupplier(1) = .Fields(1)
        Set rcD = rcP("Cap").UnderlyingValue
        If Not rcD.EOF Then
           mAccount = rcD.Fields(0)
           txtBox(3) = rcD.Fields("Bank Name")
           lblSupplier(2) = rcD.Fields("Currency")
        Else
           mAccount = ""
           txtBox(3) = ""
           lblSupplier(2) = "IDR"
        End If
     Else
        mAccount = ""
        lblSupplier(0) = ""
        lblSupplier(1) = ""
        txtBox(3) = ""
        lblSupplier(2) = "IDR"
     End If
     CloseDB rcP
End With
Err.Clear
End Sub

Private Sub SimpanDetail(ByVal Tipical As Boolean)
With RcDetail
     If .RecordCount <> 0 Then
           .MoveFirst
           Do
              
              If .EOF = True Then Exit Sub
                 'If Tipical = False Then
                    MyDDE.SendDataToServer ("DELETE FROM [Detail PO] WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields("NoItem") & "')")
                    MyDDE.SendDataToServer " INSERT INTO [Detail PO] ( PurchaseID, NoItem, QTYPO, ItemSupplierID, POPrice, ScheduleDate, VAT,QtyTemp)" & _
                                           " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "', " & .Fields("QTYPO") & ", N'" & .Fields("ItemSupplierID") & "', " & CDbl(.Fields("POPrice")) & ", convert(Datetime,'" & Format(.Fields("ScheduleDate"), "dd/mm/yy") & "',3), " & CDbl(.Fields("VAT")) & ", " & .Fields("QTYPO") & "  )"
                 'Else
                   ' MyDDE.SendDataToServer " UPDATE [Detail PO]" & _
                                           " SET QTYPO =  " & .Fields("QTYPO") & ", QTYReceive =  " & .Fields("QTYReceive") & ", QTYTemp = " & .Fields("QTYTemp") & ", ItemSupplierID =  N'" & .Fields("ItemSupplierID") & "', POPrice = 1" & CDbl(.Fields("POPrice")) & ", ScheduleDate = convert(Datetime,'" & Format(.Fields("ScheduleDate"), "dd/mm/yy") & "',3), VAT = " & CDbl(.Fields("VAT")) & _
                                           " WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields("NoItem") & "') AND (StatusTrans = 0)"
                ' End If
              .MoveNext
           Loop
           .MoveLast
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
If Me.Caption = "P.O Transaksi" Then
   TglIndex = "PO/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
Else
   TglIndex = "SC/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End If
End Function

Private Sub HitungTotal()
Dim rcTotal As New Recordset
Dim AvData As Variant
Dim mTmp, mTotal As Currency
Dim i As Long
rcTotal.CursorLocation = adUseClient
Set rcTotal = RcDetail.Clone(adLockReadOnly)
mTotal = 0
mTmp = 0
With rcTotal
     If .RecordCount <> 0 Then
        AvData = .GetRows(.RecordCount, adBookmarkFirst)
        For i = 0 To UBound(AvData, 2)
            mTmp = ((AvData(3, i) * AvData(4, i)) * (AvData(5, i) / 100)) + (AvData(3, i) * AvData(4, i))
            mTotal = mTotal + mTmp
        Next i
     Else
        mTotal = 0
     End If
End With
LblAmount = FormatNumber(mTotal, 2)
Set AvData = Nothing
CloseDB rcTotal
End Sub

Private Sub PrepareQuery()
With MyDDE
    If Me.Caption = "P.O Transaksi" Then
       .PrepareAppend = " INSERT INTO  [PO Order] ( PurchaseID, PartnerID,  DatePurchase, TermPayment,  Periode,Kurs, TypeTrans,Account) " & _
                        " VALUES (N'" & txtBox(0) & "', N'" & txtBox(2) & "',convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3) , " & txtBox(1) & ", " & Val(Month(DTPicker1.Value)) & "," & CDbl(txtBox(4)) & ", N'PO',N' & mAccount & ' )"
                        
       .PrepareUpdate = " UPDATE    [PO Order]" & _
                        " Set PartnerID = N'" & txtBox(2) & "', Kurs = " & CDbl(txtBox(4)) & ", DatePurchase = convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3), TermPayment = " & CDbl(txtBox(1)) & ", Periode = " & Val(Month(DTPicker1.Value)) & ", TypeTrans = N'PO',Account=N'" & mAccount & "'" & _
                        " WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (Status = 0)"
                        
       .PrepareDelete = " DELETE FROM  [PO Order] WHERE (TypeTrans = N'PO') AND (PurchaseID = N'" & txtBox(0) & "')"
    Else
       .PrepareAppend = " INSERT INTO  [PO Order] ( PurchaseID, PartnerID,  DatePurchase, TermPayment,  Periode,Kurs, TypeTrans) " & _
                        " VALUES (N'" & txtBox(0) & "', N'" & txtBox(2) & "',convert(Datetime, '" & Format(DTPicker1.Value, "dd/mm/yy") & "',3) , " & txtBox(1) & ", " & Val(Month(DTPicker1.Value)) & "," & CDbl(txtBox(4)) & ", N'SC' )"
       .PrepareDelete = " DELETE FROM  [PO Order] WHERE (TypeTrans = N'SC') AND (PurchaseID = N'" & txtBox(0) & "')"
    End If
End With
End Sub

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
RcCek.Open "SELECT     SUM(StockTmp) AS QTY FROM         [Inventory Tabel] GROUP BY NoItem, LEFT(RefTrans, 2) HAVING      (NoItem = N'" & NoItem & "') AND (LEFT(RefTrans, 2) = N'RN')", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCek
     If .RecordCount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
     .Close
End With
Set RcCek = Nothing
End Function

Private Function IsHeaderOk(ByVal NoPo As String) As Boolean
Dim RcIs As New Recordset
RcIs.CursorLocation = adUseClient
RcIs.Open "SELECT StatusTrans FROM  [Detail PO] GROUP BY PurchaseID, NoItem, StatusTrans, PurchaseID, NoItem HAVING      (PurchaseID = N'" & NoPo & "') AND (StatusTrans = 1)", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
IsHeaderOk = False
With RcIs
     If .RecordCount <> 0 Then
        IsHeaderOk = CBool(IIf(Not IsNull(.Fields(0)), .Fields(0), False))
     End If
End With
CloseDB RcIs
End Function
