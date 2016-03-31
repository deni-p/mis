VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmReceiveNotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penerimaan Barang"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReceiveNotes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Tag             =   "Goods Receive Note"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   10170
      TabIndex        =   7
      Top             =   0
      Width           =   10170
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TransID"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   0
         Left            =   1455
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "RN"
         Top             =   135
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1455
         MaxLength       =   200
         TabIndex        =   6
         Tag             =   "RN"
         Top             =   4770
         Width           =   4365
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   4470
         Picture         =   "frmReceiveNotes.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   855
         Width           =   405
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "No Pol"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "RN"
         Top             =   1557
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PurchaseID"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1455
         TabIndex        =   3
         Tag             =   "RN"
         Top             =   840
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2385
         Left            =   135
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   2325
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   4207
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
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "InternalName"
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
            DataField       =   "kondisi"
            Caption         =   "Kondisi"
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
            DataField       =   "UOM"
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
         BeginProperty Column04 
            DataField       =   "QTY_IN"
            Caption         =   "Qty"
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
         BeginProperty Column05 
            DataField       =   "QTY_Receive"
            Caption         =   "Qty Received"
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
         BeginProperty Column06 
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
               ColumnWidth     =   3240
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1455
         TabIndex        =   2
         Tag             =   "RN"
         Top             =   480
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
         Format          =   65601539
         CurrentDate     =   38272
      End
      Begin MSDataListLib.DataCombo CboID 
         DataField       =   "ID"
         Height          =   330
         Left            =   1455
         TabIndex        =   9
         Tag             =   "RN"
         Top             =   1200
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Expedisi"
         BoundColumn     =   "ID"
         Text            =   ""
      End
      Begin VB.Label lblSupplier 
         BackColor       =   &H00C0FFFF&
         Caption         =   "lblSupplier"
         DataField       =   "CompanyName"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   6585
         TabIndex        =   11
         Tag             =   "RN"
         Top             =   225
         Width           =   3285
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alamat"
         DataField       =   "DatePurchase"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   6585
         TabIndex        =   10
         Tag             =   "RN"
         Top             =   1230
         Width           =   3285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batas Bayar              ( Hari )"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   8
         Left            =   5400
         TabIndex        =   29
         Top             =   1560
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         DataField       =   "Kurs"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   4
         Left            =   6660
         TabIndex        =   28
         Tag             =   "RN"
         Top             =   1995
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   7545
         TabIndex        =   27
         Top             =   4815
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   7
         Left            =   6720
         TabIndex        =   26
         Top             =   4815
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblSupplier 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alamat"
         DataField       =   "city"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   6585
         TabIndex        =   25
         Tag             =   "RN"
         Top             =   765
         Width           =   3285
      End
      Begin VB.Label lblSupplier 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Partner Name"
         DataField       =   "Address"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   6585
         TabIndex        =   24
         Tag             =   "RN"
         Top             =   495
         Width           =   3285
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         DataField       =   "PartnerID"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   5235
         TabIndex        =   23
         Tag             =   "RN"
         Top             =   1995
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Date"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   5
         Left            =   5700
         TabIndex        =   22
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   4
         Left            =   5700
         TabIndex        =   21
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Penerimaan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   20
         Top             =   570
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Penerimaan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   19
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   5
         Left            =   6600
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   10
         Left            =   165
         TabIndex        =   17
         Top             =   4830
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   135
         X2              =   1700
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   150
         X2              =   1700
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   165
         X2              =   1700
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   150
         X2              =   4905
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   150
         X2              =   5835
         Y1              =   5085
         Y2              =   5085
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   150
         X2              =   1575
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   9
         Left            =   165
         TabIndex        =   16
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P.O Number"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   165
         TabIndex        =   15
         Top             =   930
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pengirim"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   165
         TabIndex        =   14
         Top             =   1275
         Width           =   690
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         DataField       =   "Mata Uang"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   7
         Left            =   8025
         TabIndex        =   13
         Tag             =   "RN"
         Top             =   1995
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblSupplier 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs"
         DataField       =   "Discount"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   8
         Left            =   8985
         TabIndex        =   12
         Tag             =   "RN"
         Top             =   1980
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5685
         X2              =   6700
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   5670
         X2              =   6685
         Y1              =   1470
         Y2              =   1470
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5250
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmReceiveNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents RcDetail             As Recordset
Attribute RcDetail.VB_VarHelpID = -1
Private WithEvents mCall                As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcEx                            As Recordset
Private RcPartner                       As New DBQuick
Private mVarPPn, mVarDisc, mVarHutang   As Variant
Private MyData                          As New clsTransaksi
Private MEdit                           As Boolean
Private pWhere As String
Dim SQLInit As String

Private Sub cboGudang_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub CboID_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then OpenPartner Index

End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0#
Response = 0
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Public Property Let IDParams(vData As String)
   pWhere = vData
   
End Property

Private Sub Form_Load()
SQLInit = " SELECT TransData.TransID, TransData.EmpID, TransData.ID, TransData.DateTrans, TransData.DateIssued, TransData.TermPayment, TransData.Kurs, TransData.Status, TransData.PurchaseID, TransData.RefNotes, TransData.TypeTrans, TransData.[No Pol], WareHouse.WareHouse, " & _
                    " WareHouse.[WareHouse Name], [PO Order].PartnerID, PartnerDB.CompanyName, PartnerDB.Address,PartnerDB.City, [PO Order].DatePurchase,[PO Order].Discount, [PO Order].CurrID AS [Mata Uang] FROM TransData LEFT OUTER JOIN WareHouse ON TransData.WareHouse = WareHouse.WareHouse INNER JOIN [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE "
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
OpenEx
DTPicker1.Value = Date
With MyDDE
    .EditModeReplace = False
    Select Case aksess.MayDo("Penerimaan Bahan Penunjang")
      Case 2, 6, 0, 5:
         .SetPermissions = 6
      Case Else
         .SetPermissions = aksess.MayDo("Penerimaan Bahan Penunjang")
    End Select
    
    Set .BindForm = frmReceiveNotes
    .BindFormTAG = "RN"
    Set .ActiveConnection = CNN
    If Trim(pWhere) = "" Then
    .PrepareQuery = SQLInit & "(TransData.TypeTrans = N'AP') AND (TransData.StatusInvoice=0) ORDER BY TransData.TransID"
    Else
    .PrepareQuery = SQLInit & " TransData.TransID ='" & pWhere & "'"
    End If
End With
Set mCall = New frmCaller
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
Set MyDDE.BindForm = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pWhere = ""
End Sub

Private Sub mCall_BeforeUnload()
If CboID.Enabled = True Then CboID.SetFocus
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm:
       Case "Purchase List":
            MyDDE.GetFieldByName("PurchaseID") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(2)
            MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName(3)
            MyDDE.GetFieldByName("Address") = mCall.GetFieldByName(4)
            MyDDE.GetFieldByName("City") = mCall.GetFieldByName("Kota")
            MyDDE.GetFieldByName("DatePurchase") = Format(mCall.GetFieldByName("Tgl. PO"), "dd mmmm yyyy")
            MyDDE.GetFieldByName("Discount") = mCall.GetFieldByName("Diskon")
            lblSupplier(4) = FormatNumber(mCall.GetFieldByName(6), 2)
            lblSupplier(5) = FormatNumber(mCall.GetFieldByName(7), 2)
            MyDDE.GetFieldByName("Kurs") = CDbl(lblSupplier(4))
            MyDDE.GetFieldByName("TermPayment") = CDbl(lblSupplier(5))
            MyDDE.GetFieldByName("Mata Uang") = mCall.GetFieldByName("Mata Uang")
            'IsiDetail MyDDE.GetFieldByName("PurchaseID")
            
       Case "BANK":
            txtBox(3) = mCall.GetFieldByName(0)
            lblSupplier(2) = mCall.GetFieldByName(1)
            
       Case "Master File Inventory":
            RcDetail.Fields(0) = mCall.GetFieldByName(0)
            RcDetail.Fields(1) = mCall.GetFieldByName(1)
            RcDetail.Fields(2) = mCall.GetFieldByName(2)
            
      Case "Barang yang Dipesan":
            MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("Kode")
            MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("Nama Barang")
            MyDDE.ChildRecordset.Fields("InternalName") = mCall.GetFieldByName("Nama Barang")
            MyDDE.ChildRecordset.Fields("kondisi") = "Kering/Basah"
            MyDDE.ChildRecordset.Fields("UOM") = mCall.GetFieldByName("Satuan")
            MyDDE.ChildRecordset.Fields("QTY_in") = mCall.GetFieldByName("QTYPO")
            MyDDE.ChildRecordset.Fields("Price") = mCall.GetFieldByName("POPrice")
            MyDDE.ChildRecordset.Fields("VAT") = mCall.GetFieldByName("VAT")
            MyDDE.ChildRecordset.Fields("Qty_receive") = 0
            MyDDE.ChildRecordset.Fields("statusItem") = "OutStanding"
            MyDDE.ChildRecordset.Fields("UOMKonversi") = mCall.GetFieldByName("UOMKonversi")
            MyDDE.ChildRecordset.Fields("warehouse") = mCall.GetFieldByName("warehouse")

End Select
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
If MEdit = True Then
   If DGPurchase.Columns(ColIndex) = "" Or IsNull(DGPurchase.Columns(ColIndex)) Then DGPurchase.Columns(ColIndex) = 0
   If CDbl(DGPurchase.Columns(5)) > CDbl(DGPurchase.Columns(4)) Then
      MessageBox "Perimaan Barang Lebih Besar Dari Barang Yang Dipesan.", "Lebih Barang", msgOkOnly
      DGPurchase.Columns(5) = 0
      DGPurchase.Columns(6) = "OutStanding"
      Exit Sub
   End If
End If

Select Case ColIndex
       Case 5:
            If RcDetail.Fields("QTY_In") <> RcDetail.Fields("QTY_Receive") Then
               RcDetail.Fields("StatusItem") = "OutStanding"
            Else
               RcDetail.Fields("StatusItem") = "Fixed"
            End If
End Select
'HitungTotal
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)

If MEdit = False Then Exit Sub
If Shift = 2 And KeyCode = vbKeyF3 Then
   MEdit = True
   RcDetail.AddNew
   DGPurchase.Columns(4) = 0
   DGPurchase.Columns(5) = 0
   DGPurchase.Columns(6) = 0
   DGPurchase.Columns(8) = 0
   DGPurchase.Columns(9) = Format(Date, "dd mmm yyyy")
   OpenPartner 2
ElseIf Shift = 2 And KeyCode = vbKeyF2 Then
   MEdit = True
   OpenPartner 2
Else
   Call Form_KeyDown(KeyCode, Shift)
End If
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
   Exit Sub
End If
With DGPurchase
     Select Case .col
            Case 0, 1, 3, 4:
                DGPurchase.MarqueeStyle = dbgHighlightRow
                .AllowUpdate = False
            Case Else:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
     End Select
End With
End Sub
Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If ScanGrid(DGPurchase) = False Then
                 If CekIsiGrid = False Then
                    MyDDE.IsChildMemberReady = True
                    MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
                    PrepareQuery
                    MEdit = True
                 End If
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
txtBox(0).Enabled = False
txtBox(3).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            cmdLink(1).Enabled = MEdit
       Case tmbAddNew:
            MEdit = True
            DTPicker1.Value = Date 'CDate(Format(dDateBegin, "dd/mm/yyyy"))
            MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
            Dim IDGen As New IDGenerator
            MyDDE.GetFieldByName("TransID") = IDGen.GetID("RN")    'MyData.PrepareIndex(tmbTransaksiReceive, 5, "1", TglIndex)
            Set IDGen = Nothing
            MyDDE.GetFieldByName("RefNotes") = "-"
            MyDDE.GetFieldByName("No Pol") = "-"
            DTPicker1.SetFocus
            cmdLink(1).Enabled = MEdit
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               MEdit = False
               cmdLink(1).Enabled = MEdit
            End If
       Case tmbCancel:
            MEdit = False
            cmdLink(1).Enabled = MEdit
       Case tmbQuit:
            Set MyDDE.BindForm = Nothing
            Unload Me
       Case tmbDetail:
            OpenPartner 3
       Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "Select * from [Receive_Notes] Where TransID='" & txtBox(0) & "'", "Receive Notes.rpt", ReportPath, "Penerimaan Barang"
            Set aReport = Nothing
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("TransID")), MyDDE.GetFieldByName("Transid"), "XXXXXXX") 'MyDDE.GetFieldByName("PurchaseID")
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1:
'            RcPartner.DBOpen "SELECT [PO Order].PurchaseID AS [PO Number], [PO Order].DatePurchase AS [Tgl. PO], " & _
'            " [PO Order].PartnerID AS [Partner ID], PartnerDB.CompanyName AS Perusahaan, " & _
'            " PartnerDB.Address AS Alamat, PartnerDB.City AS Kota, [PO Order].Kurs, " & _
'            " [PO Order].TermPayment AS Term, [PO Order].Discount AS Diskon, " & _
'            " [PO Order].CurrID AS [Mata Uang] FROM [PO Order] INNER JOIN " & _
'            " PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID WHERE " & _
'            " ([PO Order].Status = 2) " & _
'            " ORDER BY [PO Order].PurchaseID", CNN, lckLockReadOnly
             RcPartner.DBOpen "SELECT [PO Order].PurchaseID AS [PO Number], PartnerDB.CompanyName AS Perusahaan,[PO Order].DatePurchase AS [Tgl. PO], [PO Order].PartnerID AS [Partner ID]," & _
                                     "  PartnerDB.Address AS Alamat, PartnerDB.City AS Kota, [PO Order].Kurs," & _
                                     " [PO Order].TermPayment AS Term, [PO Order].Discount AS Diskon, [PO Order].CurrID AS [Mata Uang] " & _
                               " FROM  [PO Order] INNER JOIN " & _
                                     " PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID INNER JOIN " & _
                                     " [Detail PO] ON [PO Order].PurchaseID = [Detail PO].PurchaseID " & _
                              " Where ([Detail PO].StatusTrans < 2) and ([PO Order].status > 0) AND ([Detail PO].NoItem NOT LIKE 'BB%') and [po order].typetrans <> 'SO' " & _
                              " GROUP BY [PO Order].PurchaseID, [PO Order].DatePurchase, [PO Order].PartnerID, PartnerDB.CompanyName, PartnerDB.Address," & _
                                    " PartnerDB.City , [PO Order].Kurs, [PO Order].TermPayment, [PO Order].Discount, [PO Order].CurrID " & _
                              " ORDER BY [PO Order].PurchaseID ", CNN, lckLockReadOnly

       Case 2:
            RcPartner.DBOpen "SELECT NoItem, ItemName, [Serial Supplier], Merk FROM Inventory ORDER BY NoItem", CNN, lckLockReadOnly
       Case 3:
            Dim NoPo As String
            NoPo = MyDDE.GetFieldByName("purchaseID")
 
                             
            RcPartner.DBOpen "SELECT [Detail PO].NoItem as Kode, Inventory.InternalName as [Nama Barang], Inventory.UOM as satuan, Inventory.UOMKonversi, [Detail PO].QTYTemp AS QtyPo, " & _
                                   " [Detail PO].POPrice, [Detail PO].VAT, [PO Order].Discount, Inventory.WareHouse, WareHouse.[WareHouse Name] " & _
                             "FROM Inventory INNER JOIN [Detail PO] ON Inventory.NoItem = [Detail PO].NoItem " & _
                                   " INNER JOIN  [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID " & _
                                   " INNER JOIN WareHouse ON Inventory.WareHouse = WareHouse.WareHouse " & _
                             "WHERE  ([detail PO].statusTrans <= 2) AND ([Detail PO].QTYTemp <> 0) AND ([Detail PO].PurchaseID = N'" & NoPo & "') " & _
                             "ORDER BY [Detail PO].NoItem", CNN, lckLockBatch
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 1:
               mCall.FromTagActive = "Purchase List"
               mCall.txtCari = txtBox(3)
           Case 2:
               mCall.FromTagActive = "Master File Inventory"
               mCall.txtCari = txtBox(2)
               DGPurchase.Columns(7).Visible = False
               DGPurchase.Columns(8).Visible = True
           Case 3:
               mCall.FromTagActive = "Barang yang Dipesan"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If

Exit Sub
Hell:
    'MsgBox Err.Description
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
Dim rs As New Recordset
CloseDB RcDetail
Set RcDetail = New Recordset

RcDetail.CursorLocation = adUseClient
If ParameterString = "" Then ParameterString = "xxxxxxxx"
With RcDetail
     .Open " SELECT [Detail TransData].NoItem, Inventory.itemName, Inventory.InternalName, Inventory.UOM,[Detail TransData].QTY_IN, [Detail TransData].QTY_Receive, [Detail TransData].StatusItem,  TransData.PurchaseID, [Detail TransData].Price, [Detail TransData].VAT,[Detail TransData].kondisi,Inventory.UOMKonversi,inventory.warehouse FROM [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE ([Detail TransData].TransID = N'" & ParameterString & "') ORDER BY [Detail TransData].NoItem", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
     Set .ActiveConnection = Nothing
     Set DGPurchase.DataSource = RcDetail
     Set MyDDE.ChildRecordset = RcDetail
End With
End Sub

Private Sub IsiDetail(ByVal NoPo As String)
Dim rs As New Recordset
rs.CursorLocation = adUseClient
rs.Open "SELECT [Detail PO].NoItem, Inventory.ItemName, Inventory.UOM, [Detail PO].QTYTemp AS QtyPo, [Detail PO].POPrice, [Detail PO].VAT, [PO Order].Discount, Inventory.WareHouse, WareHouse.[WareHouse Name] FROM         Inventory INNER JOIN                       [Detail PO] ON Inventory.NoItem = [Detail PO].NoItem INNER JOIN                       [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID INNER JOIN                       WareHouse ON Inventory.WareHouse = WareHouse.WareHouse WHERE     ([Detail PO].QTYTemp <> 0) AND ([Detail PO].PurchaseID = N'" & NoPo & "') ORDER BY [Detail PO].NoItem", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With rs
If .Recordcount <> 0 Then
   If RcDetail.Recordcount > 0 Then

      RcDetail.MoveFirst
      Do
        If RcDetail.EOF Then Exit Do
           RcDetail.Delete adAffectCurrent
           RcDetail.MoveNext
      Loop
   End If
   Do
      If .EOF Then Exit Sub
         RcDetail.AddNew (0), .Fields(0)
         RcDetail.Fields(1) = .Fields(1)
         RcDetail.Fields("Uom") = .Fields(2)
         RcDetail.Fields("Price") = .Fields("PoPrice")
         RcDetail.Fields("Qty_In") = .Fields("QtyPo")
         RcDetail.Fields("Qty_Receive") = 0
         RcDetail.Fields("VAT") = .Fields("VAT")
         RcDetail.Fields("StatusItem") = "OutStanding"
         .MoveNext
   Loop
   .MoveFirst
End If
End With
CloseDB rs
End Sub

Private Sub SimpanDetail()
Dim MyJournal                       As New clsJournal
Dim mSisa, mDK, mTotal              As Variant
Dim mvarStatusItem, StrPartic       As String

If ScanGrid(DGPurchase) = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
    With MyDDE.ChildRecordset
         If .Recordcount <> 0 Then
            If MyDDE.SendDataToServer("DELETE FROM [Detail TransData] WHERE (TransID = N'" & txtBox(0) & "') ") = True Then
            
               '*** Update status po Order = 1 (close) ***
               SendDataToServer "update [po Order] set status=1 where purchaseID='" & MyDDE.GetFieldByName("purchaseID") & "'"
               
               .MoveFirst
               mDK = 0
               mTotal = 0
               
               While Not .EOF
                  If .Fields("QTY_IN") <> .Fields("QTY_Receive") Then
                     mvarStatusItem = "OutStanding"
                  Else
                     mvarStatusItem = "Fixed"
                  End If
                  mSisa = 0
                  'insert detail data transaksi
                  SendDataToServer " INSERT INTO [Detail TransData](TransID, NoItem, QTY_IN, QTY_Receive, Price,Vat, StatusItem,dnid,Referense,kondisi)" & _
                                   " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "', " & .Fields("QTY_IN") & ", " & .Fields("QTY_Receive") & ", " & .Fields("Price") & "," & .Fields("VAT") & ", N'" & mvarStatusItem & "',N'" & txtBox(3) & "',N'" & txtBox(3) & "','" & .Fields("kondisi") & "')"
                                   
                  If CDbl(.Fields("QTY_IN")) = CDbl(.Fields("QTY_Receive")) Then
                     mSisa = CDbl(.Fields("QTY_IN"))
                     SendAPItem .Fields("NoItem"), mSisa, CDbl(.Fields("Price")), txtBox(0), DTPicker1.Value, "RN", CDbl(lblSupplier(8)), CDbl(.Fields("VAT")), , .Fields("warehouse")
                     'update status detail po
                     SendDataToServer ("UPDATE [Detail PO] SET  StatusTrans = 2,QtyTemp = 0,QtyReceive=" & .Fields("QTY_Receive") & " ,receive_date = '" & Format(Now, "yyyy-MM-dd") & "',DNID='" & txtBox(0) & "' WHERE (PurchaseID = N'" & txtBox(3) & "') AND (NoItem = N'" & .Fields("NoItem") & "')")
                  Else
                     mSisa = CDbl(.Fields("QTY_IN")) - CDbl(.Fields("QTY_Receive"))
                     SendAPItem .Fields("NoItem"), CDbl(.Fields("QTY_Receive")), CDbl(.Fields("Price")), txtBox(0), DTPicker1.Value, "RN", CDbl(lblSupplier(8)), CDbl(.Fields("VAT")), , .Fields("warehouse")
                     'update status detail po
                     SendDataToServer ("UPDATE [Detail PO] SET  StatusTrans = 2,QtyTemp = " & mSisa & ",QtyReceive=" & .Fields("QTY_Receive") & ",receive_date = '" & Format(Now, "yyyy-MM-dd") & "' WHERE (PurchaseID = N'" & txtBox(3) & "') AND (NoItem = N'" & .Fields("NoItem") & "')")
                  End If
                  
                  
                  'Update stok gudang
                  'SendDataToServer " insert into [inventory tabel] (noItem,Qty_in,priceIn,Qty_out,priceOut,RefTrans,DateTrans,TypeTrans,lokasiGdg,stockTmp) values ('" & _
                  '                 .Fields("noItem") & "'," & Val(.Fields("Qty_receive")) * Val(.Fields("UOMKonversi")) & ",0,0,0,'" & txtBox(0).Text & "','" & Format(Now, "yyyy-MM-dd") & "','AP','" & .Fields("warehouse") & "'," & Val(.Fields("Qty_receive")) * Val(.Fields("UOMKonversi")) & ")"
                  
                  mTotal = (CDbl(.Fields("QTY_Receive")) * .Fields("Price"))
                  mDK = mDK + (mTotal - (mTotal * Round(MyDDE.GetFieldByName("Discount") / 100, 2))) + (mTotal - (mTotal * Round(MyDDE.GetFieldByName("Discount") / 100, 2))) * Round((.Fields("vat") / 100), 2)
                  ListTotalDeliver .Fields("NoItem"), txtBox(3)
                  
                  'Update SPP Doc
                  SendDataToServer "update SPP_line set receive_date='" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' where PO='" & txtBox(3).Text & "' and noItem ='" & .Fields("NoItem") & "'"
                  
                  .MoveNext
               Wend
               
               .MoveLast
               ClosePO txtBox(3)
               'SendVoucher txtBox(0), lblSupplier(0), "-", DTPicker1.value, mDK, 0, txtBox(3), "AP"
                                       
            End If
         End If
    End With
Else
   'MessageBox ""
End If
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
'If mEdit = True Then
'   If txtBox(Index).DataField <> "" And txtBox(Index).Tag <> "" Then
'      MyDDE.GetFieldByName(txtBox(Index).DataField) = txtBox(Index)
'   End If
'End If
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "RN-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub PrepareQuery()
On Error GoTo Hell:
With MyDDE
                      
     .PrepareAppend = " INSERT INTO TransData (TransID,PartnerID ,Currid, ID, DateTrans,  TermPayment,   PurchaseID, RefNotes, TypeTrans,Kurs,[no pol],warehouse,Discount) " & _
                      " VALUES (N'" & txtBox(0) & "',N'" & lblSupplier(0) & "',N'" & lblSupplier(7) & "', N'" & CboID.BoundText & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & CDbl(lblSupplier(5)) & ",   N'" & txtBox(3) & "', N'" & ValidString(txtBox(2)) & "', N'AP'," & CDbl(lblSupplier(4)) & ",N'" & ValidString(txtBox(1)) & "','" & MyDDE.ChildRecordset.Fields("Warehouse") & "'," & CDbl(lblSupplier(8)) & ")"
     .PrepareDelete = " DELETE FROM TransData WHERE (TransID = N'" & txtBox(0) & "')"
End With
Hell:
    Err.Clear
End Sub

Private Sub ClosePO(ByVal NoPo As String)
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
RcCek.Open " SELECT SUM(QTYTemp) AS QTYTemp FROM         [Detail PO] WHERE     (PurchaseID = N'" & NoPo & "') HAVING      (SUM(QTYTemp) <> 0)", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set RcCek.ActiveConnection = Nothing
With RcCek
     If .Recordcount = 0 Then
        SendDataToServer ("UPDATE [PO Order] SET Status = 1 WHERE (PurchaseID = N'" & NoPo & "')")
     End If
End With
CloseDB RcCek
End Sub

Private Sub OpenEx()
CloseDB RcEx
Set RcEx = New Recordset
RcEx.CursorLocation = adUseClient
RcEx.Open "SELECT ID, Expedisi FROM Transport ORDER BY Expedisi", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set CboID.RowSource = RcEx
End Sub

Private Function Myterm(ByVal NoPurchaseID As String) As Integer
Dim RcHpp As New DBQuick
RcHpp.DBOpen "SELECT     TermPayment FROM         [PO Order] WHERE     (PurchaseID = N'" & NoPurchaseID & "')", CNN, lckLockReadOnly
With RcHpp
     If .Recordcount <> 0 Then
        Myterm = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        Myterm = 0
     End If
End With
RcHpp.CloseDB
End Function

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

Private Sub TotalJournal()
Dim Rc              As New DBQuick
Dim I               As Integer
Dim Avdata, mTmp    As Variant
Set Rc.DBRecordset = RcDetail.Clone(adLockReadOnly)
mVarPPn = 0
mVarDisc = 0
mVarHutang = 0
mVarDisc = IIf(Not IsNull(MyDDE.GetFieldByName("Discount")), MyDDE.GetFieldByName("Discount"), 0)
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        '4 = QTY Receive 7 = Harga 8 = Vat
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            mVarHutang = mVarHutang + Avdata(4, I) * Avdata(7, I)
            If Avdata(8, I) <> 0 Then
               mVarPPn = mVarPPn + (Avdata(4, I) * Avdata(7, I)) * (Avdata(8, I) / 100)
            Else
               mVarPPn = mVarPPn + (Avdata(4, I) * Avdata(7, I))
            End If
        Next I
        If mVarDisc > 0 Then
           mVarDisc = mVarHutang * (mVarDisc / 100)
        Else
           mVarDisc = 0
        End If
        mVarHutang = (mVarHutang + mVarPPn) - mVarDisc
     End If
End With
End Sub

Private Function CekIsiGrid() As Boolean
Dim Rc              As New DBQuick
Dim I               As Integer
Dim Avdata          As Variant
Set Rc.DBRecordset = RcDetail.Clone(adLockReadOnly)
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst, "QTY_Receive")
        For I = 0 To UBound(Avdata, 2)
            If IsNull(Avdata(0, I)) = True Or IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), 0) = 0 Then
               MessageBox "Data belum diisi nilai. Harap diperiksa dulu", "Peringatan"
               CekIsiGrid = True
               Exit For
            End If
        Next I
     Else
        MessageBox "Data masih kosong.", "Peringatan"
        CekIsiGrid = True
     End If
End With
End Function

Private Sub ListTotalDeliver(ByVal NoItemData As String, ByVal ParamString As String)
Dim RcDN As New DBQuick
Dim mVarLead As Integer
If ParamString = "" Then ParamString = "XXXXX"
RcDN.DBOpen "SELECT     [PO Order].DatePurchase FROM [Detail TransData] INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID INNER JOIN [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID WHERE     ([Detail TransData].NoItem = N'" & NoItemData & "') AND ([PO Order].PurchaseID = N'" & ParamString & "') GROUP BY [PO Order].DatePurchase", CNN, lckLockReadOnly
With RcDN
     If .Recordcount <> 0 Then
        mVarLead = Abs(CDate(Format(DTPicker1.Value, "dd/mm/yyyy")) - CDate(Format(.Fields(0), "dd/mm/yyyy")))
     Else
        mVarLead = 0
     End If
     SendDataToServer ("UPDATE Inventory SET  LeadTimeDays =" & mVarLead & " WHERE  (NoItem = N'" & NoItemData & "')")
End With
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1709.858
DGPurchase.Columns(1).width = 3240
DGPurchase.Columns(2).width = 1214.929
DGPurchase.Columns(3).width = 1214.929
DGPurchase.Columns(4).width = 1214.929
DGPurchase.Columns(5).width = 1019.906
End Sub

