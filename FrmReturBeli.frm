VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmReturBeli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retur Pembelian"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmReturBeli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Return"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   24
      Top             =   5370
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   1005
      BindFormTAG     =   "RETUR"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5520
      Left            =   0
      ScaleHeight     =   5520
      ScaleWidth      =   11010
      TabIndex        =   23
      Top             =   0
      Width           =   11010
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Alasan"
         Height          =   1035
         Index           =   2
         Left            =   1665
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   14
         Tag             =   "RETUR"
         Top             =   4155
         Width           =   3465
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TransID"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "RETUR"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ReturID"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "RETUR"
         Top             =   255
         Width           =   3225
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4560
         Picture         =   "FrmReturBeli.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "RETUR"
         Top             =   608
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "typeTruck"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1665
         TabIndex        =   12
         Tag             =   "RETUR"
         Top             =   3465
         Width           =   2895
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "nopol"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1665
         TabIndex        =   13
         Tag             =   "RETUR"
         Top             =   3810
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataField       =   "jenisRL"
         DataSource      =   "MyDDE"
         Height          =   330
         ItemData        =   "FrmReturBeli.frx":6BDC
         Left            =   1665
         List            =   "FrmReturBeli.frx":6BE9
         TabIndex        =   7
         Tag             =   "RETUR"
         Text            =   "-"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         DataField       =   "kondisi"
         DataSource      =   "MyDDE"
         Height          =   330
         ItemData        =   "FrmReturBeli.frx":6C06
         Left            =   1665
         List            =   "FrmReturBeli.frx":6C13
         TabIndex        =   8
         Tag             =   "RETUR"
         Text            =   "-"
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Retur Beli"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1665
         TabIndex        =   10
         Tag             =   "RETUR"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "sak"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   3165
         TabIndex        =   11
         Tag             =   "RETUR"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PurchaseID"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   7
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "RETUR"
         Top             =   255
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "datepurchase"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   8
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "RETUR"
         Top             =   615
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "partnerID"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   9
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "RETUR"
         Top             =   975
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   10
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "RETUR"
         Top             =   1335
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Address"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   11
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "RETUR"
         Top             =   1695
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   12
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "RETUR"
         Top             =   2055
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Phone"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   13
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "RETUR"
         Top             =   2415
         Width           =   3075
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "noItem"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   14
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "RETUR"
         Top             =   2775
         Width           =   3075
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4560
         Picture         =   "FrmReturBeli.frx":6C29
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "RETUR"
         Top             =   968
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "itemName"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   330
         Index           =   15
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "RETUR"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "QtyPO"
         DataSource      =   "MyDDE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   17
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "RETUR"
         Top             =   2775
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1665
         TabIndex        =   5
         Tag             =   "RETUR"
         Top             =   1320
         Width           =   2145
         _ExtentX        =   3784
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
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "DateTrans"
         DataSource      =   "MyDDE"
         Height          =   315
         Left            =   1665
         TabIndex        =   6
         Tag             =   "RETUR"
         Top             =   1680
         Width           =   2025
         _ExtentX        =   3572
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
         Format          =   139329538
         CurrentDate     =   38272
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   6435
         X2              =   7905
         Y1              =   3075
         Y2              =   3075
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   6435
         X2              =   7905
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   6435
         X2              =   7905
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   6435
         X2              =   7905
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alasan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   43
         Top             =   4200
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   285
         TabIndex        =   42
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. &Retur"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   41
         Top             =   315
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Order"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   6495
         TabIndex        =   40
         Top             =   315
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Penerimaan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   39
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   6495
         TabIndex        =   38
         Top             =   1035
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal PO"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   6495
         TabIndex        =   37
         Top             =   675
         Width           =   825
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   255
         X2              =   1725
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   255
         X2              =   1770
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   255
         X2              =   1770
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   255
         X2              =   1935
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   285
         TabIndex        =   36
         Top             =   1740
         Width           =   285
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   255
         X2              =   1695
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   285
         TabIndex        =   35
         Top             =   2115
         Width           =   360
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   255
         X2              =   1695
         Y1              =   2715
         Y2              =   2715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kondisi"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   285
         TabIndex        =   34
         Top             =   2475
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   255
         X2              =   1680
         Y1              =   3765
         Y2              =   3765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kendaraan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   285
         TabIndex        =   33
         Top             =   3525
         Width           =   1185
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   255
         X2              =   1800
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Pol"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   12
         Left            =   285
         TabIndex        =   32
         Top             =   3870
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Retur"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   13
         Left            =   285
         TabIndex        =   31
         Top             =   3180
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   255
         X2              =   1680
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   14
         Left            =   2640
         TabIndex        =   30
         Top             =   3180
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sak"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   15
         Left            =   4275
         TabIndex        =   29
         Top             =   3180
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   11
         Left            =   6495
         TabIndex        =   28
         Top             =   2835
         Width           =   915
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   225
         X2              =   1740
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   16
         Left            =   285
         TabIndex        =   27
         Top             =   1035
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   18
         Left            =   2625
         TabIndex        =   26
         Top             =   2820
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah diterima"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   19
         Left            =   285
         TabIndex        =   25
         Top             =   2835
         Width           =   1110
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   255
         X2              =   1695
         Y1              =   3075
         Y2              =   3075
      End
   End
End
Attribute VB_Name = "FrmReturBeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall             As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner                    As New DBQuick
Private RcDetail                     As New DBQuick
Private MyData                       As New clsTransaksi
Private MEdit, mFirstCaller          As Boolean
Private pWhere As String
Dim SQLInit As String
Dim IDGen As New IDGenerator

Public Property Let IDParams(vData As String)
   pWhere = vData
   
End Property

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 4:
            If MyDDE.ChildRecordset.Fields("QTYPO") > 0 Then
                If MyDDE.ChildRecordset.Fields("Retur Beli") > MyDDE.ChildRecordset.Fields("QTYPO") Then
                   MessageBox "Stock Tidak Cukup Untuk Melakukan Transaksi Retur.", "Peringatan", msgOkOnly
                   MyDDE.ChildRecordset.Fields("Retur Beli") = 0
                End If
             Else
                 MyDDE.ChildRecordset.Fields("Retur Beli") = 0
             End If
             TotalTrans
End Select
'HitungTotal
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If MEdit = False Then Exit Sub
'Call Form_KeyDown(KeyCode, Shift)
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub


Private Sub Form_Load()
'SQLInit = " SELECT ReturData.ReturID, TransData.TransID, ReturData.DateTrans, ReturData.DateIssued, ReturData.RefNotes, ReturData.WareHouse,                        [PO Order].PurchaseID, [PO Order].PartnerID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, PartnerDB.Phone,                        [PO Order].DatePurchase, TransData.WareHouse AS [KOde Gudang], WareHouse.[WareHouse Name] as [Nama Gudang] ,[PO Order].Discount FROM         ReturData INNER JOIN TransData ON ReturData.TransID = TransData.TransID INNER JOIN [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID INNER JOIN  WareHouse ON TransData.WareHouse = WareHouse.WareHouse WHERE "
'SQLInit = "SELECT     ReturData.ReturID,  ReturData.TransID,  ReturData.EmpID,  ReturData.DateTrans,  PartnerDB.CompanyName, " & _
'                    " ReturData.RefNotes,  ReturData.jam,  ReturData.JenisRL,  ReturData.KondisiRL,  ReturData.TypeTruck,  ReturData.NoPol," & _
'                    " ReturData.berat,  ReturData.sak,  TransData.PurchaseID,  [PO Order].DatePurchase,  PartnerDB.Address,  PartnerDB.City, " & _
'                    " PartnerDB.Phone, partnerDB.partnerID " & _
'          "FROM        TransData INNER JOIN " & _
'                    "  ReturData ON  TransData.TransID =  ReturData.TransID INNER JOIN " & _
'                    "  PartnerDB ON  TransData.PartnerId =  PartnerDB.PartnerID INNER JOIN" & _
'                    "  [PO Order] ON  TransData.PurchaseID =  [PO Order].PurchaseID "

SQLInit = "select * from QueryReturBeli"
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me

Set mCall = New frmCaller
With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    Set .ActiveConnection = CNN
    
    If Trim(pWhere) = "" Then
      .PrepareQuery = SQLInit
    Else
      .PrepareQuery = SQLInit & " where TransID ='" & pWhere & "'"
    End If
   .SetPermissions = aksess.MayDo("Retur Supplier")
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set mCall = Nothing
RcPartner.CloseDB
RcDetail.CloseDB
Set MyData = Nothing
End Sub

Private Sub Form_Resize()
'On Error Resume Next
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmReturBeli = Nothing
   pWhere = ""
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
    Case 0:
        RcPartner.DBOpen "SELECT TransData.TransID, [PO Order].PurchaseID AS [PO Order], [PO Order].DatePurchase, [PO Order].PartnerID AS [Kode Supplier], " & _
                               " PartnerDB.CompanyName AS Perusahaan, PartnerDB.Address AS Alamat, PartnerDB.City AS Kota, PartnerDB.Phone AS Telepon, [PO Order].Discount " & _
                         "FROM [Detail PO] INNER JOIN " & _
                              "[PO Order] INNER JOIN " & _
                              "PartnerDB ON [PO Order].PartnerID = PartnerDB.PartnerID INNER JOIN " & _
                              "TransData ON [PO Order].PurchaseID = TransData.PurchaseID ON [Detail PO].PurchaseID = TransData.PurchaseID " & _
                         "WHERE ([Detail PO].StatusTrans = 2) OR ([Detail PO].StatusTrans = 4) AND (TransData.TypeTrans = 'RN') " & _
                         "GROUP BY [PO Order].PurchaseID, [PO Order].PartnerID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, " & _
                                  "PartnerDB.Phone, [PO Order].Discount, TransData.TransID, [PO Order].DatePurchase order by [PO Order].DatePurchase desc", CNN, lckLockReadOnly
    Case 1:
        RcPartner.DBOpen "SELECT Inventory.NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit, " & _
                                "[Detail TransData].QTY_Receive AS [QTY Beli],[Detail TransData].Price AS Harga, [Detail TransData].VAT AS Ppn " & _
                         "FROM  [Detail TransData] " & _
                                "INNER JOIN Inventory ON [Detail TransData].NoItem = Inventory.NoItem " & _
                         "WHERE ([Detail TransData].TransID = N'" & txtBox(1) & "')", CNN, lckLockBatch
        mFirstCaller = True
End Select

If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "MASTER SUPPLIER"
           Case 1: mCall.FromTagActive = "DETAIL PEMBELIAN"
                   cmdLink(0).Enabled = False
                   cmdLink(1).Enabled = False
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
'Select Case mCall.FromTagActive
'       Case "DETAIL PEMBELIAN":
'            If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
'               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'            End If
'            mFirstCaller = False
'End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm:
       Case "MASTER SUPPLIER":
            With MyDDE
                 .GetFieldByName("TransID") = mCall.GetFieldByName("TransID")
                 .GetFieldByName("PurchaseID") = mCall.GetFieldByName("PO Order")
                 .GetFieldByName("DatePurchase") = mCall.GetFieldByName("tgl Bukti")
                 .GetFieldByName("PartnerID") = mCall.GetFieldByName("Kode Supplier")
                 .GetFieldByName("companyName") = mCall.GetFieldByName("Perusahaan")
                 .GetFieldByName("Address") = mCall.GetFieldByName("Alamat")
                 .GetFieldByName("City") = mCall.GetFieldByName("kota")
                 .GetFieldByName("phone") = mCall.GetFieldByName("telepon")
                 .GetFieldByName("jenisRL") = mCall.GetFieldByName("jenis")
                 .GetFieldByName("kondisiRL") = mCall.GetFieldByName("kondisi")
                 .GetFieldByName("berat") = mCall.GetFieldByName("berat")
                 .GetFieldByName("sak") = mCall.GetFieldByName("jml sak")
                 txtBox(14).Text = mCall.GetFieldByName("NoItem")
            End With
       Case "DETAIL PEMBELIAN":
            With MyDDE
                 .GetFieldByName("noItem") = mCall.GetFieldByName("Kode Barang")
                 txtBox(15).Text = mCall.GetFieldByName("nama Barang")
                 txtBox(17).Text = mCall.GetFieldByName("Qty Beli")
                 Label1(18).Caption = mCall.GetFieldByName("Unit")
                 Label1(14).Caption = mCall.GetFieldByName("Unit")
            End With
       
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim ss As String
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            cmdLink(0).Enabled = False
            cmdLink(1).Enabled = False
            DTPicker1.SetFocus
       Case tmbAddNew:
            MEdit = False
            txtBox(0).Enabled = False
            cmdLink(0).Enabled = True
            cmdLink(1).Enabled = True
            DTPicker1.Value = Now
            DTPicker2.Value = Now
            ss = IDGen.GetID("PR")
            MyDDE.GetFieldByName("ReturID") = ss   'MyData.PrepareIndex(tmbTransaksiReturBeli, 5, "1", TglIndex)
            MyDDE.GetFieldByName("Alasan") = "-"
            MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
            MyDDE.GetFieldByName("TypeTruck") = "-"
            MyDDE.GetFieldByName("nopol") = "-"
            DTPicker1.SetFocus
       Case tmbDelete:
            SendDataToServer "update [detail PO] set QtyRetur=0,statusTrans=2 where purchaseID='" & MyDDE.GetFieldByName("purchaseID") & "' and NoItem ='" & txtBox(14).Text & "'"
            SendDataToServer "delete from ReturData where ReturID='" & MyDDE.GetFieldByName("ReturID") & "'"
            SendDataToServer "delete from [detail Retur] where ReturID='" & MyDDE.GetFieldByName("ReturID") & "' and noItem='" & MyDDE.GetFieldByName("noItem") & "'"
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               Update_DetailPO
               Update_Detail_Retur
               cmdLink(0).Enabled = False
               cmdLink(1).Enabled = False
            End If
       Case tmbCancel:
            cmdLink(0).Enabled = False
            cmdLink(1).Enabled = False
       Case tmbPrint:
         Dim aReport As New utility
         aReport.CallReportView "select * from QueryReturBeli where ReturID='" & MyDDE.GetFieldByName("ReturID") & "'", "ReturBeli.rpt", ReportPath, "Surat Retur"
         Set aReport = Nothing
            
End Select
txtBox(0).Enabled = False
txtBox(1).Enabled = False
End Sub

Private Sub Update_Detail_Retur()
   If MEdit Then
      SendDataToServer "Update [Detail Retur]" & _
                                           " ReturID = '" & MyDDE.GetFieldByName("ReturID") & _
                                            "',[Qty BJ] = " & FQty(MyDDE.GetFieldByName("Qty BJ")) & _
                                            "',[Retur Beli] = " & FQty(MyDDE.GetFieldByName("Retur Beli")) & _
                                            " ,kondisi = '" & MyDDE.GetFieldByName("kondisi") & _
                                                "',sak = " & FQty(MyDDE.GetFieldByName("sak")) & _
                                             ",JenisRL = '" & MyDDE.GetFieldByName("jenisRL") & _
                                             "',alasan = '" & MyDDE.GetFieldByName("alasan") & _
                                             "',noItem = '" & MyDDE.GetFieldByName("noItem") & _
                       "' where TransID='" & MyDDE.GetFieldByName("TransID") & "' and noItem = '" & MyDDE.GetFieldByName("noItem") & "'"
   Else
      SendDataToServer "insert into [Detail Retur] (returID,noItem,[Retur Beli],kondisi,sak,jenisRL,alasan) values (" & _
                                    " '" & MyDDE.GetFieldByName("ReturID") & _
                                    "','" & MyDDE.GetFieldByName("noItem") & _
                                    "', " & FQty(MyDDE.GetFieldByName("Retur Beli")) & _
                                    " ,'" & MyDDE.GetFieldByName("kondisi") & _
                                    "', " & FQty(MyDDE.GetFieldByName("sak")) & _
                                    " ,'" & MyDDE.GetFieldByName("jenisRL") & _
                                    "','" & MyDDE.GetFieldByName("alasan") & "')"
   End If
End Sub

Private Sub Update_DetailPO()
   SendDataToServer "update [detail PO] set QtyRetur=" & txtBox(5).Text & ",statusTrans=5 where purchaseID='" & MyDDE.GetFieldByName("purchaseID") & "' and NoItem ='" & txtBox(14).Text & "'"
End Sub


Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If Not IsNull(MyDDE.GetFieldByName("UOM")) Then
      Label1(18).Caption = MyDDE.GetFieldByName("UOM")
      Label1(14).Caption = MyDDE.GetFieldByName("UOM")
   Else
      Label1(18).Caption = ""
      Label1(14).Caption = ""
   End If
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
       Case tmbDelete:
       Case tmbSave:
           If MyDDE.CheckEmptyControl = False Then
               If (MyDDE.GetFieldByName("Retur Beli") > 0) And (MyDDE.GetFieldByName("Retur Beli") <= Val(txtBox(17).Text)) Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MessageBox "Jml Retur tidak boleh Nol & tidak > JML PO !", "Warning", msgOkOnly, msgCrtical
                  MyDDE.IsChildMemberReady = False
               End If
           Else
               MyDDE.IsChildMemberReady = False
           End If
End Select
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "RB/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function


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
Dim strSQL As String
On Error Resume Next
With MyDDE
    strSQL = " INSERT INTO ReturData" & _
                     " (ReturID,empid,TransID,DateTrans, dateIssued, RefNotes, typeTrans, Nopol, typeTruck,status,jam)" & _
                     " VALUES ('" & MyDDE.GetFieldByName("ReturID") & _
                            "','" & MainMenu.StatusBar1.Panels(1).Text & _
                            "','" & .GetFieldByName("TransID") & _
                            "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & _
                            "','" & Format(Now, "yyyy-MM-dd") & _
                            "','" & .GetFieldByName("Alasan") & "','RB'" & _
                            " ,'" & .GetFieldByName("NoPol") & _
                            "','" & .GetFieldByName("TypeTruck") & "',0" & _
                            " ,'" & Format(DTPicker2.Value, "yyyy-MM-dd hh:mm:ss") & "')"
   .PrepareAppend = strSQL
                     
   .PrepareUpdate = " UPDATE ReturData set empid='" & MainMenu.StatusBar1.Panels(1).Text & _
                                     "', TransID ='" & MyDDE.GetFieldByName("ID") & _
                                     "', DateTrans ='" & Format(DTPicker1.Value, "yyyy-MM-dd") & _
                                     "', DateIssued = '" & Format(Now, "yyyy-MM-dd") & _
                                     "', PurchaseID ='" & .GetFieldByName("PurchaseID") & _
                                     "', RefNotes = '" & .GetFieldByName("refNote") & _
                                     "', typeTruck = '" & .GetFieldByName("typeTruck") & _
                                     "', NoPol ='" & .GetFieldByName("No Pol") & _
                                     "', jam=" & Format(DTPicker2.Value, "yyyy-MM-dd hh:mm:ss") & _
                       " WHERE (ReturID = '" & .GetFieldByName("ReturID") & "')"
                     
    .PrepareDelete = " DELETE FROM  ReturData WHERE (ReturID = N'" & txtBox(0) & "')"
End With
Err.Clear
End Sub


'Private Sub SimpanDetail()
'Dim MyJournal As New clsJournal
'Dim StrPartic As String
'With MyDDE.ChildRecordset
'     If .Recordcount <> 0 Then
'           StrPartic = "Retur Pembelian "
'           .MoveFirst
'           If SendDataToServer("DELETE FROM [Detail Retur] WHERE     (ReturID = N'" & txtBox(0) & "')") = True Then
''              If SendDataToServer("DELETE From [Table Journal] where TransID =N'" & txtBox(0) & "' and TypeTrans=N'BRPB'") = True Then
''                 If MyJournal.CiptaKaryaHeaderJournal("xxx", txtBox(0), "", "", "", lblSupplier(0), "IDR", DTPicker1.Value, mVarPeriode, "BRPB") = True Then
'                    'Hutang Usaha DR
''                    MyJournal.CiptaKaryaDetailJournal "xxx", CariTypeAccount(28), lblSupplier(0), CDbl(LblTotal(2)), 0, "Hutang Usaha Ke " & lblSupplier(0)
''                    StrPartic = StrPartic & " Hutang Usaha Ke " & lblSupplier(0) & ","
''                    'Potongan Pembelian DR
''                    MyJournal.CiptaKaryaDetailJournal "xxx", CariTypeAccount(63), lblSupplier(0), CDbl(LblTotal(3)), 0, "Potongan Pembelian Pada " & lblSupplier(0)
''                    StrPartic = StrPartic & " Pot. Pemb. Ke " & lblSupplier(0) & ","
'                    Do
'                      If .EOF = True Then Exit Do
'                         SendDataToServer " INSERT INTO [Detail Retur]" & _
'                                          " (ReturID, NoItem, [QTY BJ],[Retur Beli],  Price, VAT,Hpp)" & _
'                                          " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "'," & .Fields("QTYPO") & ", " & .Fields("Retur Beli") & ",  " & .Fields("Price") & ", " & .Fields("Vat") & "," & HppProce(.Fields("NoItem")) & ")"
'                         SendDataToServer (" DELETE FROM  [Inventory Tabel] WHERE (RefTrans = N'" & txtBox(0) & "') and (Noitem=N'" & .Fields("NoItem") & "')")
'                         SendARItem .Fields("NoItem"), CCur(.Fields("Retur Beli")), CDbl(.Fields("Price")), txtBox(0), DTPicker1.value, HppProce(.Fields("NoItem")), "RB"
'                         SendDataToServer ("Update [Detail PO] Set QTYRetur =" & .Fields("Retur Beli") & " where Purchaseid=N'" & lblRN(0) & "' and NoItem=N'" & .Fields("NoItem") & "'")
''                         'Retur Pembelian CR
''                         MyJournal.CiptaKaryaDetailJournal "xxx", CariTypeAccount(58), .Fields("NoItem"), 0, .Fields(4) * .Fields(5), "Retur Pembelian " & .Fields("ItemName")
''                         StrPartic = StrPartic & " Ret. Pemb. " & .Fields("NoItem") & ","
''                         'Retur Pembelian CR
''                         MyJournal.CiptaKaryaDetailJournal "xxx", CariTypeAccount(58), .Fields("NoItem"), .Fields(4) * .Fields(5), 0, "Retur Pembelian " & .Fields("ItemName")
''                         StrPartic = StrPartic & " Ret. Pemb. " & .Fields("NoItem") & ","
''                         'Persediaan
''                         MyJournal.CiptaKaryaDetailJournal "xxx", CariAkunItem(.Fields("NoItem")), .Fields("NoItem"), 0, .Fields(4) * .Fields(5), "Persediaan " & .Fields("ItemName")
''                         StrPartic = StrPartic & "Persediaan " & .Fields("NoItem") & ","
'                        .MoveNext
'                    Loop
'
'                    'Ppn Masukan CR
  ''                  MyJournal.CiptaKaryaDetailJournal "xxx", CariTypeAccount(42), lblSupplier(0), 0, CDbl(LblTotal(1)), "PPN Masukan " & lblSupplier(0)
''                    StrPartic = StrPartic & "PPN MAsukan " & lblSupplier(0)
''                    MyJournal.CreateRefNotes StrPartic
''                 End If
''              End If
'           End If
'           .MoveLast
'           SendDataToServer (" UPDATE [Voucher Batch]" & _
'                             " Set TotalRetur = 0 WHERE (TransID = N'" & txtBox(1) & "') AND (PurchaseID = N'" & lblRN(0) & "')")
'     End If
'End With
'End Sub

'Private Function CekGridKosong() As Boolean
'Dim RcKsg As New DBQuick
'Dim Avdata As Variant
'Dim I As Integer
'Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
'With RcKsg
'     If .Recordcount <> 0 Then
'        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
'        For I = 0 To UBound(Avdata, 2)
'            If Val(IIf(Not IsNull(Avdata(3, I)), Avdata(3, I), 0)) = 0 Or Val(IIf(Not IsNull(Avdata(4, I)), Avdata(4, I), 0)) = 0 Then
'               MessageBox "Data item untuk QTY Beli atau QTY Retur ada yang berisi NOL", "Peringatan", msgOkOnly
'               CekGridKosong = True
'               MyDDE.CancelTrans = True
'               Exit For
'            End If
'        Next I
'     Else
'        CekGridKosong = True
'     End If
'End With
'Set Avdata = Nothing
'End Function


Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
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

Private Sub TotalTrans()
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Dim mDisc As Integer
'Set Rc.DBRecordset = MyDDE.ChildRecordset.Clone(adLockBatchOptimistic)
'LblTotal(0) = 0
'LblTotal(1) = 0
'LblTotal(2) = 0
'LblTotal(3) = 0
'mDisc = IIf(Not IsNull(MyDDE.GetFieldByName("Discount")), MyDDE.GetFieldByName("Discount"), 0)
'With Rc.DBRecordset
'     If .Recordcount <> 0 Then
        ' 4 = QTY Retur 5 = Harga 6 = PPn
'        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
'        For I = 0 To UBound(Avdata, 2)
'            LblTotal(0) = FormatNumber(LblTotal(0) + (Avdata(4, I) * Avdata(5, I)), 0)
'
'            LblTotal(1) = FormatNumber(LblTotal(1) + (Avdata(5, I) * (Avdata(6, I) / 100)) * (Avdata(4, I)), 0)
'
'        Next I
'        If mDisc <> 0 Then
'           LblTotal(3) = FormatNumber(CDbl(LblTotal(0)) * CDbl(mDisc / 100), 0)
'        Else
'           LblTotal(3) = 0
'        End If
'        If LblTotal(0) = "" Then LblTotal(0) = 0
'        If LblTotal(1) = "" Then LblTotal(1) = 0
'        If LblTotal(3) = "" Then LblTotal(3) = 0
'        LblTotal(2) = FormatNumber((CDbl(LblTotal(0)) - CDbl(LblTotal(3))) + CDbl(LblTotal(1)), 0)
'     End If
'End With
'Set Avdata = Nothing
End Sub

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GLAccount.NoAccount, AccType.ID, GLAccount.AccountName FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
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

Private Function TotalRetur(ByVal NoItem As String) As Boolean
Dim RcRetur As New DBQuick
RcRetur.DBOpen "SELECT      [Detail Retur].[QTY BJ] - SUM([Detail Retur].[Retur Beli]) AS [QTY Retur] FROM         [Detail Retur] INNER JOIN ReturData ON [Detail Retur].ReturID = ReturData.ReturID WHERE     (ReturData.TransID = N'" & txtBox(1) & "') AND ([Detail Retur].NoItem = N'" & NoItem & "') GROUP BY [Detail Retur].[QTY BJ]", CNN, lckLockReadOnly
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
RcRetur.DBOpen "SELECT      [Detail Retur].[QTY BJ] - SUM([Detail Retur].[Retur Beli]) AS [QTY Retur] FROM         [Detail Retur] INNER JOIN ReturData ON [Detail Retur].ReturID = ReturData.ReturID WHERE     (ReturData.TransID = N'" & txtBox(1) & "') AND ([Detail Retur].NoItem = N'" & NoItem & "') GROUP BY [Detail Retur].[QTY BJ]", CNN, lckLockReadOnly
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

