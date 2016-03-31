VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmDO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surat Jalan"
   ClientHeight    =   6090
   ClientLeft      =   360
   ClientTop       =   1365
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
   Icon            =   "FrmDO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Tag             =   "Delivery Order"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5520
      Left            =   0
      ScaleHeight     =   5520
      ScaleWidth      =   10170
      TabIndex        =   15
      Top             =   0
      Width           =   10170
      Begin VB.ComboBox cboTruck 
         Appearance      =   0  'Flat
         DataField       =   "TypeTruck"
         Height          =   330
         ItemData        =   "FrmDO.frx":6852
         Left            =   1635
         List            =   "FrmDO.frx":6877
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "RN"
         Top             =   1650
         Width           =   3345
      End
      Begin MSDataListLib.DataCombo CboID 
         DataField       =   "ID"
         Height          =   330
         Left            =   1635
         TabIndex        =   5
         Tag             =   "RN"
         Top             =   1290
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Expedisi"
         BoundColumn     =   "ID"
         Text            =   ""
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PurchaseID"
         Height          =   330
         Index           =   0
         Left            =   1635
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "RN"
         Top             =   930
         Width           =   3015
      End
      Begin VB.TextBox TxtDN 
         Appearance      =   0  'Flat
         DataField       =   "No Pol"
         Height          =   330
         Index           =   2
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "RN"
         Top             =   2010
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DN Date"
         Height          =   330
         Left            =   1635
         TabIndex        =   2
         Tag             =   "RN"
         Top             =   570
         Width           =   3345
         _ExtentX        =   5900
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
      Begin VB.TextBox TxtDN 
         Appearance      =   0  'Flat
         DataField       =   "DNID"
         Height          =   330
         Index           =   0
         Left            =   1635
         TabIndex        =   1
         Tag             =   "RN"
         Top             =   210
         Width           =   3345
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   6540
         TabIndex        =   9
         Tag             =   "RN"
         Top             =   570
         Width           =   3345
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PartnerID"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   6540
         TabIndex        =   8
         Tag             =   "RN"
         Top             =   210
         Width           =   3345
      End
      Begin VB.TextBox TxtDN 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
         Height          =   315
         Index           =   1
         Left            =   1335
         MaxLength       =   200
         TabIndex        =   13
         Tag             =   "RN"
         Top             =   4800
         Width           =   4425
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "DatePurchase"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, MMMM dd, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   6540
         TabIndex        =   10
         Tag             =   "RN"
         Top             =   930
         Width           =   3345
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4650
         Picture         =   "FrmDO.frx":6914
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   945
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   9555
         Picture         =   "FrmDO.frx":6C9E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1298
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2295
         Left            =   285
         TabIndex        =   14
         Tag             =   "Partner"
         Top             =   2445
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4048
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
         ColumnCount     =   5
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
            DataField       =   "qty_out"
            Caption         =   "QTY. SO"
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
            DataField       =   "Uom"
            Caption         =   "Satuan"
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
            DataField       =   "QTY_RECEIVE"
            Caption         =   "QTY. Kirim"
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
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label lblGudang 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         DataField       =   "GDG ID"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   6540
         TabIndex        =   11
         Tag             =   "RN"
         Top             =   1290
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   5205
         TabIndex        =   29
         Top             =   278
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order No."
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   28
         Top             =   998
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Order"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   5205
         TabIndex        =   27
         Top             =   998
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   26
         Top             =   285
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   315
         TabIndex        =   25
         Top             =   645
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Courier Services"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   24
         Top             =   1358
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   315
         TabIndex        =   23
         Top             =   2085
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kendaraan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   315
         TabIndex        =   22
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   21
         Top             =   4845
         Width           =   945
      End
      Begin VB.Label lblGudang 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "Nama Gudang"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   5820
         TabIndex        =   20
         Tag             =   "RN"
         Top             =   1635
         Width           =   3345
      End
      Begin VB.Label lblGudang 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "Alamat"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   5820
         TabIndex        =   19
         Tag             =   "RN"
         Top             =   1890
         Width           =   3345
      End
      Begin VB.Label lblGudang 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "Kota"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   5820
         TabIndex        =   18
         Tag             =   "RN"
         Top             =   2145
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   5205
         TabIndex        =   17
         Top             =   1358
         Width           =   555
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   315
         X2              =   1700
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   315
         X2              =   1700
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   315
         X2              =   1700
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   315
         X2              =   1700
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   315
         X2              =   1700
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   315
         X2              =   1700
         Y1              =   2325
         Y2              =   2325
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   5190
         X2              =   6600
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5190
         X2              =   6600
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   5190
         X2              =   6600
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   315
         X2              =   1545
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama customer"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   5205
         TabIndex        =   16
         Top             =   638
         Width           =   1110
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   5190
         X2              =   6600
         Y1              =   885
         Y2              =   885
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1005
      BindFormTAG     =   "RN"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDetail As New DBQuick
Private RcPartner As New DBQuick
Private RcEx As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData As New clsTransaksi
Private MEdit As Boolean, mFirstCaller As Boolean
Private Irow As Long
Private mVarBtn As ButtonTransDB
Private pWhere As String
Dim IDGen As New IDGenerator
Dim SQLInit As String

Private Sub CboID_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cboTruck_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo 1
Select Case ColIndex
       Case 4:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
               If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = "0"
               If CDbl(DGPurchase.Columns(ColIndex)) > OpenQTY(MyDDE.ChildRecordset.Fields(0)) Then
                  MessageBox "Data Yang Dimasukan Lebih Besar Dari QTY. Yang Seharusnya Dikirim.", "Peringatan", msgOkOnly, msgCrtical
                  MyDDE.ChildRecordset.Fields("QTY_RECEIVE") = 0
               End If
            End If
End Select
Exit Sub
1:
MessageBox Err.Description, "frmdo:dgpurchase_aftercoledit" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
'If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
'Call DGPurchase_KeyDown(KeyCode, Shift)
End Sub

Public Property Let IDParams(vData As String)
   pWhere = vData
End Property


Private Sub Form_Load()
SQLInit = " SELECT TransData.TransID AS DNID, TransData.PurchaseID, Transport.ID, Transport.Expedisi, TransData.DateTrans AS [DN DATE], TransData.[No Pol],  TransData.TypeTruck, TransData.Status, PartnerDB.CompanyName, [PO Order].DatePurchase, TransData.PartnerId, TransData.RefNotes,  [Gudang Customer].[GDG ID], [Gudang Customer].[Nama Gudang], [Gudang Customer].Alamat, Regional.[RG Name] AS Kota,[PO Order].Discount, [PO Order].Kurs, [PO Order].CurrID AS [Mata Uang]" & _
                     " FROM Regional INNER JOIN [Gudang Customer] ON Regional.RG = [Gudang Customer].RG RIGHT OUTER JOIN Transport INNER JOIN TransData ON Transport.ID = TransData.ID INNER JOIN" & _
                     " PartnerDB INNER JOIN [PO Order] ON PartnerDB.PartnerID = [PO Order].PartnerID ON TransData.PurchaseID = [PO Order].PurchaseID ON  [Gudang Customer].[GDG ID] = TransData.[GDG ID] WHERE  "
                     
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
OpenEx
With MyDDE
     .EditModeReplace = False
     Set .BindForm = FrmDO
     .BindFormTAG = "RN"
     Set .ActiveConnection = CNN
   If Trim(pWhere) = "" Then
      .PrepareQuery = SQLInit & "  (TransData.TypeTrans = N'DN') AND (TransData.StatusInvoice = 0)"
   Else
      .PrepareQuery = SQLInit & "TransData.TransID ='" & pWhere & "'"
   End If
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set mCall = Nothing
Set MyData = Nothing
Set mCall = Nothing
RcDetail.CloseDB
RcPartner.CloseDB
RcEx.CloseDB
MyDDE.ClearRecordset
End Sub

Private Sub Form_Resize()

Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmDO = Nothing
pWhere = ""
End Sub

Private Sub mCall_BeforeUnload()
On Error GoTo 1
Select Case mCall.FromTagActive
       Case "DAFTAR PENJUALAN":
            If CboID.Enabled = True Then CboID.SetFocus
       Case "MASTER GUDANG"
            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
       Case "DETAIL PENJUALAN":
            If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            End If
            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
            mFirstCaller = False
End Select
Exit Sub
1:
MessageBox Err.Description, "frmdo:mcall_beforeunload" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub mCall_CallLinkForm()
If mCall.FromTagActive = "MASTER GUDANG" Then
   frmWareHouse.SetFocus
   frmWareHouse.ZOrder (0)
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 3
Select Case TagForm:
       Case "PO":
            MyDDE.GetFieldByName("ID") = mCall.GetFieldByName(0)
            MyDDE.GetFieldByName("Expedisi") = mCall.GetFieldByName(1) 'TxtDN(1)
       Case "DAFTAR PENJUALAN":
            MyDDE.GetFieldByName("PurchaseID") = mCall.GetFieldByName(0) 'txtBox(0)
            MyDDE.GetFieldByName("DatePurchase") = mCall.GetFieldByName(1) 'txtBox(1)
            MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName("Kode Customer")
            MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName("Nama Perusahaan")
            MyDDE.GetFieldByName("Discount") = mCall.GetFieldByName("DIskon")
            MyDDE.GetFieldByName("Kurs") = mCall.GetFieldByName("Kurs")
            MyDDE.GetFieldByName("Mata Uang") = mCall.GetFieldByName("Mata Uang")
            OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DNID")), MyDDE.GetFieldByName("DNID"), "XXXXXXX"), False
       Case "MASTER GUDANG":
            MyDDE.GetFieldByName("GDG ID") = mCall.GetFieldByName(0) 'txtBox(0)
            MyDDE.GetFieldByName("Nama Gudang") = mCall.GetFieldByName(1) 'txtBox(1)
            MyDDE.GetFieldByName("Alamat") = mCall.GetFieldByName(2)
            MyDDE.GetFieldByName("Kota") = mCall.GetFieldByName(3)
       Case "DETAIL PENJUALAN":
            MyDDE.ChildRecordset.Fields(0) = mCall.GetFieldByName(0)
            MyDDE.ChildRecordset.Fields(1) = mCall.GetFieldByName(1)
            MyDDE.ChildRecordset.Fields(2) = mCall.GetFieldByName(2)
            MyDDE.ChildRecordset.Fields(3) = mCall.GetFieldByName(3)
            MyDDE.ChildRecordset.Fields("qty_out") = mCall.GetFieldByName(4)
            MyDDE.ChildRecordset.Fields("VAT") = mCall.GetFieldByName("VAT")
            MyDDE.ChildRecordset.Fields("qty_Receive") = 0
            MyDDE.ChildRecordset.Fields("Price") = mCall.GetFieldByName("harga")
            MyDDE.ChildRecordset.Fields("StatusItem") = "OutStanding"
            MyDDE.ChildRecordset.Fields("lokasigdg") = mCall.GetFieldByName("warehouse")
End Select
Exit Sub
3:
MessageBox Err.Description, "frmdo:mcall_rowcallchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If MEdit = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
''If mEdit = False Then Exit Sub
'If MyDDE.CheckEmptyControl = False Then
'   If Shift = 2 And KeyCode = vbKeyF3 Then
'     mEdit = True
'      DGPurchase.Columns(3) = 0
'      DGPurchase.Columns(4) = 0
'      OpenPartner 2
'      If MyDDE.ChildRecordset.Recordcount <> 0 Then Irow = MyDDE.ChildRecordset.AbsolutePosition
'   ElseIf Shift = 2 And KeyCode = vbKeyF2 Then
'      mEdit = True
'      OpenPartner 2
'   End If
'End If
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo 2
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
Else
    With DGPurchase
         Select Case .col
                Case 0, 1, 2, 3:
                    DGPurchase.MarqueeStyle = dbgFloatingEditor
                    .AllowUpdate = False
                Case Else:
                    DGPurchase.MarqueeStyle = dbgFloatingEditor
                    .AllowUpdate = True
         End Select
    End With
End If
Exit Sub
2:
MessageBox Err.Description, "frmdo:dgpurchase_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbEdit:
            MyDDE.CancelTrans = CekInvoiceClosed
            If MyDDE.CancelTrans = True Then
               MessageBox "Data Sudah Valid/Closed Oleh Transaksi Invoice", "Peringatan", msgOkOnly, msgCrtical
            End If
       Case tmbDetail:
            If MyData.CheckGridKosong(MyDDE.ChildRecordset, "fldtotal") = True Then
               MyDDE.CancelTrans = True
               MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly, msgCrtical
            Else
               MyDDE.CancelTrans = mFirstCaller
            End If
       Case tmbDelete:
            If RcDetail.Recordcount <> 0 Then
                DetailDelete
                MEdit = False
                mVarBtn = BtnNone
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If ScanGrid(DGPurchase) = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
                  PrepareQuery
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete: PrepareQuery
End Select
Exit Sub
1:
MessageBox Err.Description, "frmdo:mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
txtBox(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            cmdLink(0).Enabled = False
            cmdLink(1).Enabled = False
            TxtDN(0).Enabled = False
            txtBox(1).Enabled = False
            txtBox(2).Enabled = False
            txtBox(3).Enabled = False
            mVarBtn = BtnEdit
            'DGPurchase.Columns(0).Button = True
       Case tmbAddNew:
            MEdit = True
            DTPicker1.Value = CDate(Format(dDateBegin, "dd/mm/yyyy"))
            MyDDE.GetFieldByName("DN DATE") = DTPicker1.Value
            MyDDE.GetFieldByName("DNID") = IDGen.GetID("DO") 'MyData.PrepareIndex(tmbDeliveryNotes, 5, "", TglIndex)
            MyDDE.GetFieldByName("RefNotes") = "-"
            MyDDE.GetFieldByName("No Pol") = "-"
            TxtDN(0).Enabled = False
            DTPicker1.SetFocus
            txtBox(1).Enabled = False
            txtBox(2).Enabled = False
            txtBox(3).Enabled = False
            mVarBtn = BtnAddnew
            'DGPurchase.Columns(0).Button = True
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               MEdit = False
               mVarBtn = BtnNone
               'DGPurchase.Columns(0).Button = False
            End If
       Case tmbDetail:
            If mFirstCaller = False Then
               OpenPartner 2
               MEdit = True
            End If
'            Call DGPurchase_KeyDown(vbKeyF3, 2)
              
       Case tmbPrint:
            Dim aReport As New utility
            aReport.CallReportView "select * from [form SJ] where sj='" & TxtDN(0) & "'", "surat jalan.rpt", ReportPath, "Surat Jalan"
            Set aReport = Nothing
       Case tmbCancel:
            mVarBtn = BtnNone
            MEdit = False
            'DGPurchase.Columns(0).Button = False
End Select
'If mVarBtn <> BtnEdit Then
   cmdLink(0).Enabled = MEdit
   cmdLink(1).Enabled = MEdit
'End If
Exit Sub
1:
MessageBox Err.Description, "frmdo:mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("DNID")), MyDDE.GetFieldByName("DNID"), "XXXXXXX"), True 'MyDDE.GetFieldByName("DNID")
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo 5
Select Case Index
       Case 0:
            'RcPartner.DBOpen "SELECT     [PO Order].PurchaseID AS [No SO], [PO Order].DatePurchase AS Tanggal, PartnerDB.CompanyName AS [Nama Perusahaan],                        PartnerDB.PartnerID AS [Kode Customer], [PO Order].Discount AS Diskon,[PO Order].Kurs, [PO Order].CurrID AS [Mata Uang] FROM         PartnerDB INNER JOIN                       [PO Order] ON PartnerDB.PartnerID = [PO Order].PartnerID WHERE     ([PO Order].TypeTrans = N'SO') AND ([PO Order].Status = 0 or [PO Order].Status = 2) ORDER BY [PO Order].PurchaseID", CNN, lckLockReadOnly
            RcPartner.DBOpen "SELECT     [PO Order].PurchaseID AS [No SO], [PO Order].DatePurchase AS Tanggal, PartnerDB.CompanyName AS [Nama Perusahaan], PartnerDB.PartnerID AS [Kode Customer], [PO Order].Discount AS Diskon,[PO Order].Kurs, [PO Order].CurrID AS [Mata Uang],[PO Order].approved_by FROM  PartnerDB INNER JOIN  [PO Order] ON PartnerDB.PartnerID = [PO Order].PartnerID WHERE     ([PO Order].TypeTrans = N'SO') AND ([PO Order].Status = 0 or [PO Order].Status = 2) and ([PO Order].approved_by <>'') ORDER BY [PO Order].PurchaseID", CNN, lckLockReadOnly
       Case 1:
            If Not IsNull(lblGudang(0)) Then
               RcPartner.DBOpen "SELECT [Gudang Customer].[GDG ID], [Gudang Customer].[Nama Gudang], [Gudang Customer].Alamat, Regional.[RG Name] as Kota FROM [Gudang Customer] INNER JOIN  Regional ON [Gudang Customer].RG = Regional.RG WHERE     ([Gudang Customer].PartnerID = N'" & txtBox(2) & "') ORDER BY [Gudang Customer].[GDG ID]", CNN, lckLockReadOnly
            Else
               RcPartner.DBOpen "SELECT [Gudang Customer].[GDG ID], [Gudang Customer].[Nama Gudang], [Gudang Customer].Alamat, Regional.[RG Name] FROM [Gudang Customer] INNER JOIN  Regional ON [Gudang Customer].RG = Regional.RG WHERE     ([Gudang Customer].PartnerID = N'XXXXX') ORDER BY [Gudang Customer].[GDG ID]", CNN, lckLockReadOnly
            End If
       Case 2:
           ' RcPartner.DBOpen "SELECT  [Detail PO].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.[Serial Supplier], Inventory.UOM AS Unit,  [Detail PO].QTYTemp AS [QTY Blm Terkirim], [Detail PO].POPrice AS Harga, [Detail PO].VAT, [Detail PO].QTYPO AS [QTY Order] FROM [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem WHERE ([Detail PO].QTYTemp > 0) AND ([Detail PO].PurchaseID = N'" & txtBox(0) & "') ORDER BY [Detail PO].NoItem", CNN, lckLockReadOnly
            RcPartner.DBOpen "SELECT  [Detail PO].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.[Serial Supplier], Inventory.UOM AS Unit,  [Detail PO].QTYTemp AS [QTY Blm Terkirim], [Detail PO].POPrice AS Harga, [Detail PO].VAT, [Detail PO].QTYPO AS [QTY Order],Inventory.warehouse FROM [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem WHERE ([Detail PO].QTYTemp > 0) AND ([Detail PO].PurchaseID = N'" & txtBox(0) & "') ORDER BY [Detail PO].NoItem", CNN, lckLockReadOnly
            mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "DAFTAR PENJUALAN"
            mCall.txtCari = txtBox(2)
          Case 1:
            mCall.FromTagActive = "MASTER GUDANG"
            mCall.CaptionLink = "Gudang"
          Case 2:
            mCall.FromTagActive = "DETAIL PENJUALAN"
            'mCall.txtCari = TxtDN(1)
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data masih Kosong Atau Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
    Err.Clear
Exit Sub
5:
MessageBox Err.Description, "frmdo:openpartner" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDetail(ByVal ParameterString As String, Optional ByVal Tipical As Boolean)
On Error GoTo 4
If ParameterString = "" Then ParameterString = "xxxxxxxx"
With RcDetail
     .DBOpen "SELECT [Detail TransData].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], Inventory.UOM, [Detail TransData].QTY_OUT , [Detail TransData].Price,  [Detail TransData].VAT, [Detail TransData].QTY_Receive, [Detail TransData].StatusItem,[Detail TransData].LokasiGdg FROM [Detail TransData] INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem WHERE ([Detail TransData].TransID = N'" & ParameterString & "') ORDER BY [Detail TransData].NoItem", CNN
     Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone
     Set DGPurchase.DataSource = MyDDE.ChildRecordset
     
End With
RcDetail.CloseDB
Exit Sub
4:
MessageBox Err.Description, "frmdo:opendetail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "DN/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub PrepareQuery()
On Error GoTo Hell:
With MyDDE
     .PrepareAppend = " INSERT INTO TransData (empID,TransID,Kurs,CurrID , ID, DateTrans,     PurchaseID, RefNotes, TypeTrans,PartnerID,[No Pol],TypeTruck,[GDG ID],TermPayment,Discount) " & _
                      " VALUES (N'" & mVarLoginActive & "',N'" & TxtDN(0) & "'," & MyDDE.GetFieldByName("Kurs") & ",N'" & MyDDE.GetFieldByName("Mata Uang") & "', N'" & CboID.BoundText & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3),N'" & txtBox(0) & "', N'" & TxtDN(1) & "', N'DN',N'" & txtBox(2) & "',N'" & TxtDN(2) & "',N'" & cboTruck.Text & "',N'" & lblGudang(0) & "'," & Myterm(txtBox(0)) & "," & MyDDE.GetFieldByName("Discount") & ")"
          '  MessageBox .PrepareAppend
     .PrepareUpdate = " UPDATE TransData" & _
                      " Set empID = N'" & mVarLoginActive & "', discount=" & MyDDE.GetFieldByName("Discount") & ",Kurs=" & MyDDE.GetFieldByName("Kurs") & ",CurrID=N'" & MyDDE.GetFieldByName("Mata Uang") & "',ID = N'" & CboID.BoundText & "', DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & _
                      " PurchaseID = N'" & txtBox(0) & "', RefNotes = N'" & TxtDN(1) & "', TypeTrans = N'DN', PartnerID = N'" & txtBox(2) & "',[GDG ID] =N'" & lblGudang(0) & "',TermPayment=" & Myterm(txtBox(0)) & " WHERE     (TransID = N'" & TxtDN(0) & "') "
                      
     .PrepareDelete = " DELETE FROM TransData WHERE (TransID = N'" & TxtDN(0) & "')"
End With
Exit Sub
Hell:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub DetailDelete()
On Error GoTo xErr
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        .MoveFirst
        Do
          If .EOF = True Then Exit Do
             SendDataToServer ("UPDATE [Detail PO]" & _
                               " SET   QTYTemp = QTYTemp + " & .Fields("QTY_Receive") & _
                               " WHERE (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & RcDetail.Fields("NoItem") & "')")
             
             'Mengembalikan stock setelah di hapus
             KembaliStock RcDetail.Fields("NoItem"), MyDDE.ChildRecordset("lokasigdg"), .Fields("QTY_Receive")
             CekItemStatus
             .MoveNext
        Loop
        .MoveLast
        CekDOClosed
     End If
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub SimpanDetail()
Dim mSisa As Long
Dim mDK As Currency
Dim mvarStatusItem As String
On Error GoTo xErr
If ScanGrid(DGPurchase) = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
    With MyDDE.ChildRecordset
         If .Recordcount <> 0 Then
            If MyDDE.SendDataToServer("DELETE FROM [Detail TransData] WHERE (Transid = N'" & TxtDN(0) & "') ") = True Then
               .MoveFirst
               mDK = 0
               
               Do
                  If .EOF = True Then Exit Do
                     If .Fields("QTY_OUT") <> .Fields("QTY_Receive") Then
                        mvarStatusItem = "OutStanding"
                     Else
                        mvarStatusItem = "Fixed"
                     End If
                     mSisa = 0
                     SendDataToServer " INSERT INTO [Detail TransData](TransID, NoItem, QTY_OUT, QTY_Receive, Price,Vat, StatusItem,Hpp,DNID,lokasigdg)" & _
                                      " VALUES (N'" & TxtDN(0) & "', N'" & .Fields("NoItem") & "', " & OpenQTY(.Fields("NoItem")) & ", " & .Fields("QTY_Receive") & ", " & .Fields("Price") & "," & .Fields("VAT") & ", N'" & mvarStatusItem & "'," & HppProce(txtBox(0), .Fields("NoItem")) & ",N'" & txtBox(0) & "','" & MyDDE.ChildRecordset("lokasigdg") & "')"
                     
                     'digunakan mengurangi stock di inventory tabel
                              KurangStock .Fields("noitem"), MyDDE.ChildRecordset("lokasigdg"), .Fields("QTY_Receive")
                                                             
                     Select Case mVarBtn
                            Case BtnEdit:
                                 If Not IsNull(.Fields("QTY_Receive").OriginalValue) Then
                                    If .Fields("QTY_Receive").OriginalValue <> .Fields("QTY_Receive") Then
                                       SendDataToServer (" UPDATE [Detail PO]" & _
                                                         " SET   QTYTemp = QTYTemp + " & IIf(Not IsNull(.Fields("QTY_Receive").OriginalValue), .Fields("QTY_Receive").OriginalValue, 0) - .Fields("QTY_Receive") & _
                                                         " WHERE (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields("NoItem") & "')")
                                    End If
                                 End If
                            Case BtnAddnew:
                                 SendDataToServer (" UPDATE [Detail PO]" & _
                                                   " SET   QTYTemp = QTYTemp - " & .Fields("QTY_Receive") & _
                                                   " WHERE (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields("NoItem") & "')")
                                                   
                             
                                                                 
                     End Select
                     CekItemStatus
                     mDK = mDK + (mSisa * .Fields("Price")) * (.Fields("vat") / 100) + (mSisa * .Fields("Price"))
                     .MoveNext
               Loop
               .MoveLast
               CekDOClosed
            End If
         End If
    End With
Else
   'MessageBox ""
End If
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Function OpenQTY(ByVal NoItem As String) As Long
On Error GoTo 6
Dim rsQty As New DBQuick
rsQty.DBOpen "SELECT QTYTEMP FROM [Detail PO] WHERE (NoItem = N'" & NoItem & "') AND (PurchaseID = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
With rsQty
     If .Recordcount <> 0 Then
        OpenQTY = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        OpenQTY = 0
     End If
End With
rsQty.CloseDB
Exit Function
6:
MessageBox Err.Description, "frmdo:openqty" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub OpenEx()
RcEx.DBOpen "SELECT ID, Expedisi FROM Transport ORDER BY Expedisi", CNN, lckLockReadOnly
Set CboID.RowSource = RcEx.DBRecordset
End Sub

Private Sub CekItemStatus()
Dim rcq As New DBQuick
On Error GoTo xErr
rcq.DBOpen "SELECT QTYPO, QTYTemp FROM [Detail PO] WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & MyDDE.ChildRecordset.Fields("NoItem") & "')", CNN, lckLockReadOnly
With rcq
     If .Recordcount <> 0 Then
        If .Fields(1) = .Fields(0) Then
           SendDataToServer ("UPDATE    [Detail PO] SET StatusTrans = 0 WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & MyDDE.ChildRecordset.Fields("NoItem") & "')")
        ElseIf .Fields(1) <= .Fields(0) And .Fields(1) > 0 Then
           SendDataToServer ("UPDATE    [Detail PO] SET StatusTrans = 1 WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & MyDDE.ChildRecordset.Fields("NoItem") & "')")
        ElseIf .Fields(1) = 0 Then
           SendDataToServer ("UPDATE    [Detail PO] SET StatusTrans = 0 WHERE     (PurchaseID = N'" & txtBox(0) & "') AND (NoItem = N'" & MyDDE.ChildRecordset.Fields("NoItem") & "')")
        End If
     End If
End With
rcq.CloseDB
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Function CekInvoiceClosed() As Boolean
On Error GoTo 2
Dim RcInvoice As New DBQuick
RcInvoice.DBOpen " SELECT     Status FROM   TransData WHERE     (DNID = N'" & TxtDN(0) & "') AND (PurchaseID = N'" & txtBox(0) & "') ", CNN, lckLockReadOnly
With RcInvoice
     If .Recordcount <> 0 Then
        CekInvoiceClosed = CBool(IIf(Not IsNull(.Fields(0)), .Fields(0), False))
     Else
        CekInvoiceClosed = False
     End If
End With
RcInvoice.CloseDB
Exit Function
2:
MessageBox Err.Description, "frmdo:cekinvoiseclosed" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub CekDOClosed()
On Error GoTo 1
Dim RcInvoice As New DBQuick
RcInvoice.DBOpen " SELECT     SUM(QTYTemp) AS QTYPO FROM         [Detail PO] WHERE     (PurchaseID = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
With RcInvoice
     If .Recordcount <> 0 Then
        If .Fields(0) = 0 Then
           SendDataToServer ("UPDATE [PO Order] SET  Status =1 WHERE     (PurchaseID = N'" & txtBox(0) & "')")
        Else
           SendDataToServer ("UPDATE [PO Order] SET  Status =0 WHERE     (PurchaseID = N'" & txtBox(0) & "')")
        End If
     End If
End With
RcInvoice.CloseDB
Exit Sub
1:
MessageBox Err.Description, "frmdo:checkdoclose" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Function HppProce(ByVal NoPurchaseID As String, ByVal NoItem As String) As Double
On Error GoTo 4
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
Exit Function
4:
MessageBox Err.Description, "frmdo:hppproce" & Err.Number, msgOkOnly, msgExclamation
End Function

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

Private Sub TxtDN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1860.095
DGPurchase.Columns(1).width = 2865.26
DGPurchase.Columns(2).width = 1514.835
DGPurchase.Columns(3).width = 1184.882
DGPurchase.Columns(4).width = 1514.835
End Sub

Private Sub KurangStock(ByVal NoItem As String, ByVal gdg As String, ByVal qtyout As Integer)
On Error Resume Next
Dim RcQty As New DBQuick
Dim StockTmp As Integer
RcQty.DBOpen "SELECT    noitem,qty_in,qty_out,stocktmp   from [inventory tabel] WHERE noitem = N'" & NoItem & "' and lokasigdg=N'" & gdg & "'", CNN, lckLockReadOnly
If RcQty.Recordcount > 0 Then
With RcQty
     If .Recordcount <> 0 Then
         StockTmp = RcQty.Fields("stocktmp") - qtyout
         SendDataToServer "UPDATE [inventory tabel] SET stocktmp =" & StockTmp & ", qty_out=" & qtyout & "   WHERE  noitem = N'" & NoItem & "' and lokasigdg='" & gdg & "'"
     End If
End With
End If
RcQty.CloseDB
End Sub

Private Sub KembaliStock(ByVal NoItem As String, ByVal gdg As String, ByVal qtyout As Integer)
'On Error Resume Next
Dim RcQty As New DBQuick
Dim StockTmp As Integer
RcQty.DBOpen "SELECT    noitem,qty_in,qty_out,stocktmp   from [inventory tabel] WHERE noitem = N'" & NoItem & "' and lokasigdg=N'" & gdg & "'", CNN, lckLockReadOnly '
With RcQty
     If .Recordcount <> 0 Then
         StockTmp = RcQty.Fields("stocktmp") + qtyout
         SendDataToServer "UPDATE [inventory tabel] SET stocktmp =" & StockTmp & ", qty_out=" & qtyout & "   WHERE  noitem = N'" & NoItem & "' and lokasigdg='" & gdg & "'"
     End If
End With
RcQty.CloseDB
End Sub

Private Sub TutupSO()
SendDataToServer ("UPDATE [po Order] SET StatusSJ = 1 WHERE     (PurchaseID = N'" & txtBox(0) & "')")
End Sub
