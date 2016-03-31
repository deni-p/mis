VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistoryPO 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7875
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11940
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7635
      Left            =   120
      ScaleHeight     =   7605
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   120
      Width           =   11730
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   400
         Left            =   10320
         TabIndex        =   6
         Top             =   7080
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   135
         ScaleHeight     =   6825
         ScaleWidth      =   11415
         TabIndex        =   1
         Top             =   120
         Width           =   11445
         Begin MSComctlLib.ListView ListHeader 
            Height          =   2535
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No PO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tanggal"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Customer"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Alamat"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Batas Bayar"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Loco"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Currency"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Kurs"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Discount"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView listDetail 
            Height          =   3375
            Left            =   120
            TabIndex        =   3
            Top             =   3240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   5953
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No Barang"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nama Barang"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Harga"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "PPN"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Tgl Kirim"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Data Detail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Daftar Purchase Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmHistoryPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dsSO  As New DBQuick
Dim dsDetailSO As New DBQuick
Private xID As String
Private xTgl As String

Public Property Let SONo(vData As String)
   xID = vData
End Property

Public Property Let RangeTanggal(vData As String)
   xTgl = vData
End Property


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim lsv As ListItem
   strSQL = " SELECT * FROM PurchaseOrderValid "
   If (Trim(xID) <> "") Or (Trim(xTgl) <> "") Then
      strSQL = strSQL & " where "
      If Trim(xID) <> "" Then
         strSQL = strSQL & " purchaseID ='" & Trim(xID) & "'"
      End If
      If Trim(xTgl) <> "" Then
         If Trim(xID) <> "" Then strSQL = strSQL & " and "
         strSQL = strSQL & " DatePurchase " & Trim(xTgl)
      End If
   End If
   
   dsSO.DBOpen strSQL, CNN, lckLockBatch, lckLockSync
   
   ListHeader.ListItems.Clear
   If dsSO.Recordcount > 0 Then
      dsSO.MoveTopRecord
      For X = 1 To dsSO.Recordcount
        Set lsv = ListHeader.ListItems.Add(X, "SO" & X, dsSO.Fields("PurchaseID"))
        ListHeader.ListItems(X).SubItems(1) = Format(dsSO.Fields("DatePurchase"), "dd MMM yyyy")
        ListHeader.ListItems(X).SubItems(2) = dsSO.Fields("CompanyName")
        ListHeader.ListItems(X).SubItems(3) = Trim(dsSO.Fields("Address")) & " " & dsSO.Fields("city")
        ListHeader.ListItems(X).SubItems(4) = dsSO.Fields("termPayment")
        ListHeader.ListItems(X).SubItems(5) = dsSO.Fields("TypeFreight")
        ListHeader.ListItems(X).SubItems(6) = dsSO.Fields("Currency Name")
        ListHeader.ListItems(X).SubItems(7) = dsSO.Fields("Kurs")
        ListHeader.ListItems(X).SubItems(8) = dsSO.Fields("Discount")
        dsSO.MoveNextRecord
      Next
   Else
      MsgBox strSQL
   End If
      
End Sub

Private Sub ListHeader_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim lsv As ListItem
   strSQL = " SELECT  [Detail PO].NoItem, Inventory.ItemName, [Detail PO].ItemSupplierID, [Detail PO].QTYPO, [Detail PO].POPrice, [Detail PO].VAT, [Detail PO].ScheduleDate, [Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2) " & _
                " + ([Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2))   * ROUND([Detail PO].VAT / 100, 2) AS FldTotal, [Detail PO].POPrice AS TMP, [Detail PO].PurchaseID, [Detail PO].QTYTemp, [Detail PO].StatusTrans, inventory.UOM FROM         [Detail PO] INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem INNER JOIN  [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE     ([Detail PO].PurchaseID = N'" & Trim(Item.Text) & "') ORDER BY [Detail PO].NoItem"
   dsDetailSO.DBOpen strSQL, CNN
   
   listDetail.ListItems.Clear
   If dsDetailSO.Recordcount > 0 Then
      dsDetailSO.MoveTopRecord
      For X = 1 To dsDetailSO.Recordcount
         Set lsv = listDetail.ListItems.Add(X, "D" & X, dsDetailSO.Fields("NoItem"))
         listDetail.ListItems(X).SubItems(1) = dsDetailSO.Fields("itemName")
         listDetail.ListItems(X).SubItems(2) = dsDetailSO.Fields("UOM")
         listDetail.ListItems(X).SubItems(3) = dsDetailSO.Fields("QtyPO")
         listDetail.ListItems(X).SubItems(4) = dsDetailSO.Fields("POPrice")
         listDetail.ListItems(X).SubItems(5) = dsDetailSO.Fields("VAT")
         listDetail.ListItems(X).SubItems(6) = dsDetailSO.Fields("fldTotal")
         listDetail.ListItems(X).SubItems(7) = Format(dsDetailSO.Fields("ScheduleDate"), "dd MMM yyyy")
         dsDetailSO.MoveNextRecord
      Next
   End If
End Sub

