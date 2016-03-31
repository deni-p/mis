VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmValidasi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validasi Tutup Buku"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12975
   Tag             =   "Period Closing"
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   12945
      TabIndex        =   16
      Top             =   7320
      Width           =   12975
      Begin VB.CommandButton CmdOk 
         Caption         =   "Proses Validasi"
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   120
         Width           =   1620
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Transfer Journal"
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   1710
         TabIndex        =   11
         Top             =   120
         Width           =   1620
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Simpan"
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   4950
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Keluar"
         Height          =   360
         Index           =   3
         Left            =   3330
         TabIndex        =   12
         Top             =   120
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      Height          =   7320
      Left            =   0
      ScaleHeight     =   7320
      ScaleWidth      =   12975
      TabIndex        =   14
      Top             =   0
      Width           =   12975
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0FFFF&
         Height          =   570
         Left            =   3345
         ScaleHeight     =   510
         ScaleWidth      =   9465
         TabIndex        =   15
         Top             =   90
         Width           =   9525
         Begin VB.ComboBox cboPeriode 
            Height          =   315
            Index           =   0
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   105
            Width           =   2235
         End
         Begin VB.ComboBox cboPeriode 
            Height          =   315
            Index           =   1
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   120
            Width           =   2625
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periode"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   165
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   195
            Index           =   1
            Left            =   3255
            TabIndex        =   4
            Top             =   165
            Width           =   570
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6495
         Left            =   3345
         TabIndex        =   6
         Top             =   720
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   11456
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   15380335
         TabCaption(0)   =   "Detail Validasi"
         TabPicture(0)   =   "frmValidasi.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ListView4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ListView1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Detail Closing"
         TabPicture(1)   =   "frmValidasi.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ListView3"
         Tab(1).Control(1)=   "ListView2"
         Tab(1).ControlCount=   2
         Begin MSComctlLib.ListView ListView1 
            Height          =   3720
            Left            =   60
            TabIndex        =   7
            Top             =   345
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   6562
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Validasi"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tanggal Transaksi"
               Object.Width           =   2928
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bukti Transaksi"
               Object.Width           =   2928
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Referensi"
               Object.Width           =   2928
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Customer"
               Object.Width           =   2928
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Total"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Discount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "PPN"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Grand Total"
               Object.Width           =   3175
            EndProperty
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   3720
            Left            =   -74940
            TabIndex        =   8
            Top             =   345
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   6562
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Closing"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "No Journal"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bukti Transaksi"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tanggal Transaksi"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Keterangan"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Tipe Trans"
               Object.Width           =   2417
            EndProperty
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2370
            Left            =   -74940
            TabIndex        =   9
            Top             =   4065
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   4180
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Kode Perkiraan"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Keterangan"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dok Ref"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "DEBET"
               Object.Width           =   3087
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "KREDIT"
               Object.Width           =   3087
            EndProperty
         End
         Begin MSComctlLib.ListView ListView4 
            Height          =   2370
            Left            =   60
            TabIndex        =   17
            Top             =   4065
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   4180
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No.Bukti"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Kode"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Nama Barang"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   3087
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Harga"
               Object.Width           =   3087
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6765
         Left            =   60
         TabIndex        =   1
         Top             =   405
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   11933
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Distribution"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   0
         Top             =   105
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmValidasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rc As New DBQuick
Private StrTanggalPeriode, mVarIndex, mVarNoUrutJournal, AkumAccount, mVarAccDep, mVarDepre As String
Private PrsValid As Boolean
Private TotalAkum As Variant
Private WithEvents RcList As DBQuick
Attribute RcList.VB_VarHelpID = -1
Private RcDetail As New DBQuick
Private mVarJournal As Boolean
Private mVarJrl As String

Private Sub cboPeriode_Click(Index As Integer)
Select Case Index
       Case 0:
            Opendata TreeView1.SelectedItem.Key, cboPeriode(Index).ListIndex + 1
            TanggalPeriode
'            If SSTab1.Tab = 1 Then OpenList
       Case 1:
            FilterTanggalPeriode
End Select
End Sub

Private Sub cmdOk_Click(Index As Integer)
Dim I As Integer
Select Case Index
       Case 0:
            PrsValid = True
            cmdOk(0).Enabled = False
            cmdOk(1).Enabled = True
            cmdOk(2).Enabled = True
       Case 1:
            PrsValid = False
            cmdOk(0).Enabled = True
            cmdOk(1).Enabled = False
            cmdOk(2).Enabled = False
            I = MessageBox("Anda yakin untuk melakukan proses Closing.", "Closing", msgYesNo, msgQuestion)
            If I = 1 Then
               If AccountLink <> "xxx" Then
'                  PrepareJournalFixAssets
                  Closing cboPeriode(0).ListIndex + 1
                  MessageBox "Proses closing telah selesai.", "Closing", msgOkOnly, msgInfo
                  'If PeriodeBerjalan = False Then FrmSetingPeriode.SetFocus
'                  Unload Me
                Else
                  MessageBox "Seting kode akun tampungan rugi laba belum tersedia.", "Kode Akun", msgOkOnly, msgExclamation
                End If
            End If
       Case 2:
            PrsValid = False
            cmdOk(0).Enabled = False
            cmdOk(1).Enabled = True
            cmdOk(2).Enabled = True
       Case 3:
            Unload Me
End Select
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'CmdOk(0).Left = Picture1.Left + 20
'CmdOk(0).Top = Picture1.Height + 150
'CmdOk(1).Left = CmdOk(0).Left + CmdOk(0).Width
'CmdOk(1).Top = CmdOk(0).Top
'CmdOk(2).Left = CmdOk(1).Left + CmdOk(1).Width
'CmdOk(2).Top = CmdOk(1).Top
'cmdOk(3).Left = cmdOk(2).Left + cmdOk(2).Width
cmdOk(3).Top = cmdOk(2).Top
SSTab1.Tab = 0
Set RcList = New DBQuick
Set RcDetail = New DBQuick
OpenMenu
CreatePeriode
Label2.FontBold = True
SSTab1.Tab = 0
TreeView1.Nodes(2).Selected = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Rc = Nothing
End Sub

Private Sub Form_Resize()
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmValidasi = Nothing
End Sub

Private Sub OpenMenu()
TreeView1.Indentation = 300
With TreeView1.Nodes.Add(, , "Pembelian", "PEMBELIAN")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Pembelian", tvwChild, "Order pembelian", "Order Pembelian"
TreeView1.Nodes.Add "Pembelian", tvwChild, "Retur pembelian", "Retur Pembelian"
TreeView1.Nodes.Add "Pembelian", tvwChild, "Petty", "Petty Cash"

With TreeView1.Nodes.Add(, , "Penjualan", "PENJUALAN")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Penjualan", tvwChild, "Order penjualan", "Order Penjualan"
TreeView1.Nodes.Add "Penjualan", tvwChild, "Retur penjualan", "Retur Penjualan"

With TreeView1.Nodes.Add(, , "Akunting", "KEUANGAN")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Akunting", tvwChild, "Pelunasan Hutang", "Hutang"
TreeView1.Nodes.Add "Akunting", tvwChild, "Pelunasan Piutang", "Piutang"

'With TreeView1.Nodes.Add("Akunting", tvwChild, "Transaksi Aktiva Tetap", "Fixed Asset")
'     .Bold = True
'     .Expanded = True
'End With
'TreeView1.Nodes.Add "Transaksi Aktiva Tetap", tvwChild, "Penjualan Aktiva", "Sales"
'TreeView1.Nodes.Add "Transaksi Aktiva Tetap", tvwChild, "Pembelian Aktiva", "Purchase"

TreeView1.Nodes.Item(1).Selected = True
TreeView1.Nodes(3).Selected = True

With TreeView1.Nodes.Add(, , "Prod", "PRODUKSI")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Prod", tvwChild, "BFlushWIP", "Barang Dalam Proses (WIP)"
TreeView1.Nodes.Add "Prod", tvwChild, "BFlushFG", "Barang Jadi (FG)"

With TreeView1.Nodes.Add(, , "Payroll", "PAYROLL")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Payroll", tvwChild, "Direct Labor", "Tenaga Kerja Langsung"
TreeView1.Nodes.Add "Payroll", tvwChild, "Indirect Labor", "Tenaga Kerja Tidak Langsung"


End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If PrsValid = True Then
   Rc.DBRecordset.AbsolutePosition = Item.Index
   If Item.Checked = True Then
      Item.Text = "VALIDASI"
   Else
      Item.Text = "NO"
      SendDataToServer ("Delete from [Table Journal] where Transid=N'" & Item.SubItems(2) & "'")
   End If
   mVarIndex = Item.SubItems(2)
   SimpanValidasi Item.Checked
Else
   Rc.DBRecordset.AbsolutePosition = Item.Index
   If Not IsNull(Rc.DBRecordset.Fields(0)) Then
      Item.Checked = CBool(Rc.DBRecordset.Fields(0))
   Else
      Item.Checked = False
   End If
   MessageBox "Tombol proses validasi belum dipilih.", "Peringatan", msgOkOnly
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
mVarIndex = Item.SubItems(2)
End Sub

Private Sub ListView2_Click()
If RcList.DBRecordset.Recordcount <> 0 Then RcList.DBRecordset.AbsolutePosition = ListView2.SelectedItem.Index
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If PrsValid = True Then
   RcList.DBRecordset.AbsolutePosition = Item.Index
   If Item.Checked = True Then
      Item.Text = "CLOSING"
   Else
      Item.Text = "NO"
   End If
'   mVarIndex = Item.SubItems(1)
    SendDataToServer (" Update [Table Journal] set Status = " & BoolToInt(Item.Checked) & "where JournalID=N'" & Item.SubItems(1) & "'")
    SendDataToServer (" Update  [TransData] Set Closed = " & BoolToInt(Item.Checked) & " where Transid=N'" & Item.SubItems(2) & "'")
Else
    RcList.DBRecordset.AbsolutePosition = Item.Index
    Item.Checked = CBool(RcList.DBRecordset.Fields(0))
    MessageBox "Tombol proses validasi belum dipilih.", "Peringatan", msgOkOnly
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub RcList_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Dim I As Integer
Dim Avdata, mDr, mCr As Variant
'If Not RcDetail.DBRecordset Is Nothing Then
ListView3.ListItems.Clear
Set RcDetail.DBRecordset = pRecordset("ChildMd").UnderlyingValue
If Not RcDetail.DBRecordset Is Nothing Then
With RcDetail.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            With ListView3.ListItems.Add(, , Avdata(1, I))
                 .SubItems(1) = Avdata(3, I)
                 .SubItems(2) = Avdata(4, I)
                 .SubItems(3) = FormatNumber(Avdata(5, I), 0)
                 .SubItems(4) = FormatNumber(Avdata(6, I), 0)
            End With
            mDr = mDr + Avdata(5, I)
            mCr = mCr + Avdata(6, I)
        Next I
        With ListView3.ListItems.Add(, , "")
             .SubItems(2) = "Total"
             .SubItems(3) = FormatNumber(mDr, 0)
             .SubItems(4) = FormatNumber(mCr, 0)
        End With
     End If
End With
End If
'End If
Set Avdata = Nothing
Set mDr = Nothing
Set mCr = Nothing
Err.Clear
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'ListView2.ListItems.Clear
'ListView3.ListItems.Clear
'If SSTab1.Tab = 1 Then OpenList
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Opendata Node.Key, cboPeriode(0).ListIndex + 1
Label2 = Node.Text
End Sub


Function headkol(customer As Boolean)
If customer = False Then
   ListView1.ColumnHeaders(5).Text = "Supplier"
Else
   ListView1.ColumnHeaders(5).Text = "Customer"
End If
End Function

Private Sub Opendata(ByVal MainTrans As String, ByVal PeriodeData As Integer)
On Error Resume Next
Select Case UCase(MainTrans)
       Case "ORDER PEMBELIAN":
            headkol False
            'Rc.DBOpen " SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Kode Transaksi], TransData.PurchaseID AS [PO Ref],TransData.PartnerId AS Supplier, SUM([Detail TransData].QTY_IN * [Detail TransData].Price) AS Total, SUM(([Detail TransData].QTY_IN * [Detail TransData].Price) * (TransData.Discount / 100)) AS Discount,SUM(([Detail TransData].QTY_IN * [Detail TransData].Price - ([Detail TransData].QTY_IN * [Detail TransData].Price) * (TransData.Discount / 100)) " & _
                      " * ROUND([Detail TransData].VAT / 100, 2)) AS PPN, SUM([Detail TransData].QTY_IN * [Detail TransData].Price - [Detail TransData].QTY_IN * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2) + ([Detail TransData].QTY_IN * [Detail TransData].Price - [Detail TransData].QTY_IN * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2))* ROUND([Detail TransData].VAT / 100, 2)) AS [Total Pembelian] FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE     (LEFT(TransData.TransID, 2) = N'RN') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans, TransData.TransID, TransData.PurchaseID, TransData.PartnerId, TransData.Validasi ORDER BY TransData.DateTrans, TransData.TransID", CNN, lckLockReadOnly
            
            Rc.DBOpen " SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Kode Transaksi], TransData.PurchaseID AS [PO Ref], dbo.PartnerDB.CompanyName, SUM([Detail TransData].QTY_IN * [Detail TransData].Price) AS Total, SUM(([Detail TransData].QTY_IN * [Detail TransData].Price) * (TransData.Discount / 100)) AS Discount,SUM(([Detail TransData].QTY_IN * [Detail TransData].Price - ([Detail TransData].QTY_IN * [Detail TransData].Price) * (TransData.Discount / 100)) " & _
                      " * ROUND([Detail TransData].VAT / 100, 2)) AS PPN, SUM([Detail TransData].QTY_IN * [Detail TransData].Price - [Detail TransData].QTY_IN * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2) + ([Detail TransData].QTY_IN * [Detail TransData].Price - [Detail TransData].QTY_IN * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2))* ROUND([Detail TransData].VAT / 100, 2)) AS [Total Pembelian] FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID  inner join  dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE    (LEFT(TransData.TransID, 2) = N'RN') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans, TransData.TransID, TransData.PurchaseID, PartnerDB.CompanyName, TransData.Validasi ORDER BY TransData.DateTrans, TransData.TransID", CNN, lckLockReadOnly
            
            
            StrTanggalPeriode = "SELECT TransData.DateTrans AS Tanggal FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE  (LEFT(TransData.TransID, 2) = N'RN') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans ORDER BY TransData.DateTrans"
            OpenList "BPBK"
            
            
'============Tambahan = =====================================
       Case "PETTY":
            headkol False
            
            Rc.DBOpen " SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Kode Transaksi], TransData.PurchaseID AS [PO Ref], dbo.PartnerDB.CompanyName, SUM([Detail TransData].QTY_IN * [Detail TransData].Price) AS Total, SUM(([Detail TransData].QTY_IN * [Detail TransData].Price) * (TransData.Discount / 100)) AS Discount,SUM(([Detail TransData].QTY_IN * [Detail TransData].Price - ([Detail TransData].QTY_IN * [Detail TransData].Price) * (TransData.Discount / 100)) " & _
                      " * ROUND([Detail TransData].VAT / 100, 2)) AS PPN, SUM([Detail TransData].QTY_IN * [Detail TransData].Price - [Detail TransData].QTY_IN * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2) + ([Detail TransData].QTY_IN * [Detail TransData].Price - [Detail TransData].QTY_IN * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2))* ROUND([Detail TransData].VAT / 100, 2)) AS [Total Pembelian] FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID  inner join  dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE    (LEFT(TransData.TransID, 2) = N'PC') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans, TransData.TransID, TransData.PurchaseID, PartnerDB.CompanyName, TransData.Validasi ORDER BY TransData.DateTrans, TransData.TransID", CNN, lckLockReadOnly
            StrTanggalPeriode = "SELECT TransData.DateTrans AS Tanggal FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE  (LEFT(TransData.TransID, 2) = N'PC') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans ORDER BY TransData.DateTrans"
            OpenList "BPPC"
'============================================================
            
       Case "RETUR PEMBELIAN":
            headkol False
            'Rc.DBOpen " SELECT ReturData.Validasi, ReturData.DateTrans AS [Tanggal], ReturData.ReturID AS [Bukti Transaksi], ReturData.TransID AS Referense,  TransData.PartnerId AS [Partner Id], SUM([Detail Retur].[Retur Beli] * [Detail Retur].Price) AS [Sub Total], SUM([Detail Retur].[Retur Beli] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) AS Discount, SUM(([Detail Retur].[Retur Beli] * [Detail Retur].Price - [Detail Retur].[Retur Beli] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) * ROUND([Detail Retur].VAT / 100, 2)) AS PPN, " & _
                      " SUM((([Detail Retur].[Retur Beli] * [Detail Retur].Price - [Detail Retur].[Retur Beli] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) * Round([Detail Retur].VAT / 100, 2) + [Detail Retur].[Retur Beli] * [Detail Retur].Price) - [Detail Retur].[Retur Beli] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) As Total" & _
                      " FROM ReturData INNER JOIN  TransData ON ReturData.TransID = TransData.TransID INNER JOIN  [Detail Retur] ON ReturData.ReturID = [Detail Retur].ReturID WHERE     (ReturData.TypeTrans = N'RB') AND (MONTH(ReturData.DateTrans) = " & PeriodeData & ") GROUP BY ReturData.DateTrans, ReturData.ReturID, ReturData.TransID, TransData.PartnerId, ReturData.Validasi ORDER BY ReturData.DateTrans, ReturData.ReturID", CNN, lckLockReadOnly
            
            Rc.DBOpen " SELECT ReturData.Validasi, ReturData.DateTrans AS [Tanggal], ReturData.ReturID AS [Bukti Transaksi], ReturData.TransID AS Referense,  dbo.PartnerDB.CompanyName , SUM([Detail Retur].[Retur Beli] * [Detail Retur].Price) AS [Sub Total], SUM([Detail Retur].[Retur Beli] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) AS Discount, SUM(([Detail Retur].[Retur Beli] * [Detail Retur].Price - [Detail Retur].[Retur Beli] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) * ROUND([Detail Retur].VAT / 100, 2)) AS PPN, " & _
                      " SUM((([Detail Retur].[Retur Beli] * [Detail Retur].Price - [Detail Retur].[Retur Beli] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) * Round([Detail Retur].VAT / 100, 2) + [Detail Retur].[Retur Beli] * [Detail Retur].Price) - [Detail Retur].[Retur Beli] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) As Total" & _
                      " FROM ReturData INNER JOIN  TransData ON ReturData.TransID = TransData.TransID INNER JOIN  [Detail Retur] ON ReturData.ReturID = [Detail Retur].ReturID  inner join dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE     (ReturData.TypeTrans = N'RB') AND (MONTH(ReturData.DateTrans) = " & PeriodeData & ") GROUP BY ReturData.DateTrans, ReturData.ReturID, ReturData.TransID, PartnerDB.CompanyName, ReturData.Validasi ORDER BY ReturData.DateTrans, ReturData.ReturID", CNN, lckLockReadOnly
            
            StrTanggalPeriode = "SELECT     DateTrans AS Tanggal FROM         ReturData WHERE     (TypeTrans = N'RB')  AND (MONTH(ReturData.DateTrans) = " & PeriodeData & ") GROUP BY DateTrans ORDER BY DateTrans"
            OpenList "BRPB"
            
       Case "ORDER PENJUALAN":
            headkol True
            'Rc.DBOpen " SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Kode Transaksi], TransData.PurchaseID AS [PO Ref],TransData.PartnerId AS Supplier, SUM([Detail TransData].QTY_OUT * [Detail TransData].Price) AS Total, SUM(([Detail TransData].QTY_OUT * [Detail TransData].Price) * (TransData.Discount / 100)) AS Discount,SUM(([Detail TransData].QTY_OUT * [Detail TransData].Price - ([Detail TransData].QTY_OUT * [Detail TransData].Price) * (TransData.Discount / 100)) " & _
                      " * ROUND([Detail TransData].VAT / 100, 2)) AS PPN, SUM([Detail TransData].QTY_OUT * [Detail TransData].Price - [Detail TransData].QTY_OUT * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2) + ([Detail TransData].QTY_OUT * [Detail TransData].Price - [Detail TransData].QTY_OUT * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2))* ROUND([Detail TransData].VAT / 100, 2)) AS [Total Pembelian] FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE     (LEFT(TransData.TransID, 2) = N'AR') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans, TransData.TransID, TransData.PurchaseID, TransData.PartnerId, TransData.Validasi ORDER BY TransData.DateTrans, TransData.TransID", CNN, lckLockReadOnly
            
            Rc.DBOpen " SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Kode Transaksi], TransData.PurchaseID AS [PO Ref],dbo.PartnerDB.CompanyName, SUM([Detail TransData].QTY_OUT * [Detail TransData].Price) AS Total, SUM(([Detail TransData].QTY_OUT * [Detail TransData].Price) * (TransData.Discount / 100)) AS Discount,SUM(([Detail TransData].QTY_OUT * [Detail TransData].Price - ([Detail TransData].QTY_OUT * [Detail TransData].Price) * (TransData.Discount / 100)) " & _
                      " * ROUND([Detail TransData].VAT / 100, 2)) AS PPN, SUM([Detail TransData].QTY_OUT * [Detail TransData].Price - [Detail TransData].QTY_OUT * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2) + ([Detail TransData].QTY_OUT * [Detail TransData].Price - [Detail TransData].QTY_OUT * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2))* ROUND([Detail TransData].VAT / 100, 2)) AS [Total Pembelian] FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID inner join  dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE     (LEFT(TransData.TransID, 2) = N'AR') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans, TransData.TransID, TransData.PurchaseID,  PartnerDB.CompanyName , TransData.Validasi ORDER BY TransData.DateTrans, TransData.TransID", CNN, lckLockReadOnly
            
            StrTanggalPeriode = "SELECT TransData.DateTrans AS Tanggal FROM TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE  (LEFT(TransData.TransID, 2) = N'AR') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ") GROUP BY TransData.DateTrans ORDER BY TransData.DateTrans"
            OpenList "BPJK"
            
       Case "RETUR PENJUALAN":
            headkol True
            'Rc.DBOpen " SELECT ReturData.Validasi, ReturData.DateTrans AS [Tanggal], ReturData.ReturID AS [Bukti Transaksi], TransData.PurchaseID AS Referense,  TransData.PartnerId AS [Partner Id], SUM([Detail Retur].[Retur Jual] * [Detail Retur].Price) AS [Sub Total], SUM([Detail Retur].[Retur Jual] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) AS Discount, SUM(([Detail Retur].[Retur Jual] * [Detail Retur].Price - [Detail Retur].[Retur Jual] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) * ROUND([Detail Retur].VAT / 100, 2)) AS PPN, " & _
                      " SUM((([Detail Retur].[Retur Jual] * [Detail Retur].Price - [Detail Retur].[Retur Jual] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) * Round([Detail Retur].VAT / 100, 2) + [Detail Retur].[Retur Jual] * [Detail Retur].Price) - [Detail Retur].[Retur Jual] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) As Total" & _
                      " FROM ReturData INNER JOIN  TransData ON ReturData.TransID = TransData.TransID INNER JOIN  [Detail Retur] ON ReturData.ReturID = [Detail Retur].ReturID WHERE     (ReturData.TypeTrans = N'RJ') AND (MONTH(ReturData.DateTrans) = " & PeriodeData & ") GROUP BY ReturData.DateTrans, ReturData.ReturID, TransData.PurchaseID, TransData.PartnerId, ReturData.Validasi ORDER BY ReturData.DateTrans, ReturData.ReturID", CNN, lckLockReadOnly
            
            Rc.DBOpen " SELECT ReturData.Validasi, ReturData.DateTrans AS [Tanggal], ReturData.ReturID AS [Bukti Transaksi], TransData.PurchaseID AS Referense,  dbo.PartnerDB.CompanyName, SUM([Detail Retur].[Retur Jual] * [Detail Retur].Price) AS [Sub Total], SUM([Detail Retur].[Retur Jual] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) AS Discount, SUM(([Detail Retur].[Retur Jual] * [Detail Retur].Price - [Detail Retur].[Retur Jual] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2)) * ROUND([Detail Retur].VAT / 100, 2)) AS PPN, " & _
                      " SUM((([Detail Retur].[Retur Jual] * [Detail Retur].Price - [Detail Retur].[Retur Jual] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) * Round([Detail Retur].VAT / 100, 2) + [Detail Retur].[Retur Jual] * [Detail Retur].Price) - [Detail Retur].[Retur Jual] * [Detail Retur].Price * Round(TransData.Discount / 100, 2)) As Total" & _
                      " FROM ReturData INNER JOIN  TransData ON ReturData.TransID = TransData.TransID INNER JOIN  [Detail Retur] ON ReturData.ReturID = [Detail Retur].ReturID  inner join dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE (ReturData.TypeTrans = N'RJ') AND (MONTH(ReturData.DateTrans) = " & PeriodeData & ") GROUP BY ReturData.DateTrans, ReturData.ReturID, TransData.PurchaseID, PartnerDB.CompanyName, ReturData.Validasi ORDER BY ReturData.DateTrans, ReturData.ReturID", CNN, lckLockReadOnly
            
            
            StrTanggalPeriode = "SELECT DateTrans AS Tanggal FROM ReturData WHERE (TypeTrans = N'RJ')  AND (MONTH(ReturData.DateTrans) = " & PeriodeData & ") GROUP BY DateTrans ORDER BY DateTrans"
            OpenList "BRPJ"
            
       Case "PELUNASAN PIUTANG":
            headkol True
            Rc.DBOpen "SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Bukti Transaksi], TransData.PurchaseID AS Referense,  dbo.PartnerDB.CompanyName, [Detail TransData].Debet AS [Sub Total], 0 AS Discount, 0 AS PPN, [Detail TransData].Debet AS Total FROM         TransData INNER JOIN   [Detail TransData] ON TransData.TransID = [Detail TransData].TransID inner join  dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE     (TransData.TypeTrans = N'BR') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ")"
            OpenList "BKM"
            
       Case "PELUNASAN HUTANG":
            headkol False
            Rc.DBOpen "SELECT TransData.Validasi, TransData.DateTrans AS Tanggal, TransData.TransID AS [Bukti Transaksi], TransData.PurchaseID AS Referense,  dbo.PartnerDB.CompanyName, [Detail TransData].Credit AS [Sub Total], 0 AS Discount, 0 AS PPN, [Detail TransData].Credit AS Total FROM         TransData INNER JOIN   [Detail TransData] ON TransData.TransID = [Detail TransData].TransID inner join  dbo.PartnerDB ON dbo.transdata.PartnerId = dbo.PartnerDB.PartnerID WHERE     (TransData.TypeTrans = N'BP') AND (MONTH(TransData.DateTrans) = " & PeriodeData & ")"
            OpenList "BKK"
            
       Case "PEMBELIAN AKTIVA":
            headkol True
            Rc.DBOpen " SELECT [TR Aktiva Tetap].Validasi, [TR Aktiva Tetap].DateTrans AS Tanggal, [TR Aktiva Tetap].[No FA] AS [Bukti Transaksi], '-' AS Referense,[TR Aktiva Tetap].PartnerID AS [Partner ID], [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga AS [Sub Total], [TR Aktiva Tetap].DP AS Discount,  [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS PPN,                       [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga + [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga * ROUND([DTR Aktiva Tetap].PPn  / 100, 2) AS Total" & _
                      " FROM [TR Aktiva Tetap] INNER JOIN  [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].TypeTrans = N'FB') AND (MONTH([TR Aktiva Tetap].DateTrans) = " & PeriodeData & ") ORDER BY [TR Aktiva Tetap].DateTrans, [TR Aktiva Tetap].[No FA]", CNN, lckLockReadOnly
            StrTanggalPeriode = " SELECT DateTrans FROM [TR Aktiva Tetap] WHERE (TypeTrans = N'FB') AND (MONTH(DateTrans) = " & PeriodeData & ") GROUP BY DateTrans"
            OpenList "BKKAT"
            
       Case "PENJUALAN AKTIVA":
            headkol False
            Rc.DBOpen " SELECT [TR Aktiva Tetap].Validasi, [TR Aktiva Tetap].DateTrans AS Tanggal, [TR Aktiva Tetap].[No FA] AS [Bukti Transaksi], '-' AS Referense, [TR Aktiva Tetap].PartnerID AS [Partner ID], [DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga AS [Sub Total], [TR Aktiva Tetap].DP AS Discount, ([DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga - [TR Aktiva Tetap].DP) * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS PPN, " & _
                      " [DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga - [TR Aktiva Tetap].DP + ([DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga - [TR Aktiva Tetap].DP) * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS Total FROM         [TR Aktiva Tetap] INNER JOIN  [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].TypeTrans = N'FJ') AND (MONTH([TR Aktiva Tetap].DateTrans) = " & PeriodeData & ") ORDER BY [TR Aktiva Tetap].DateTrans, [TR Aktiva Tetap].[No FA]", CNN, lckLockReadOnly
            StrTanggalPeriode = "SELECT     DateTrans FROM         [TR Aktiva Tetap] WHERE (TypeTrans = N'FJ') AND (MONTH(DateTrans) = " & PeriodeData & ") GROUP BY DateTrans"
            OpenList "BKMAT"
      
      Case "BFLUSHWIP":
            headkol False
            Rc.DBOpen "SELECT backflush_header.validasi, backflush_header.dateTrans as Tanggal, " & _
                             " backflush_header.IDTrans as [Bukti Transaksi], backflush_header.OrderID as [Referense], sum(backflush_line.cost) as Total " & _
                      "FROM backflush_header INNER JOIN " & _
                           "backflush_line on backflush_header.IDTrans = backflush_line.IDTrans " & _
                      "WHERE (LEFT([backflush_Header].IDTrans, 3) = N'BFL') AND (MONTH(DateTrans) = " & PeriodeData & ")" & _
                      "GROUP BY backflush_header.IDTrans,backflush_header.validasi, backflush_header.dateTrans , " & _
                               "backflush_line.cost,backflush_header.IDTrans, backflush_header.OrderID", CNN, lckLockBatch
            OpenList "BFWIP"
            
      Case "BFLUSH":
            headkol False
            Rc.DBOpen "SELECT backflush_header.validasi, backflush_header.dateTrans as Tanggal, " & _
                             " backflush_header.IDTrans as [Bukti Transaksi], backflush_header.OrderID as [Referense], sum(backflush_line.cost) as Total " & _
                      "FROM backflush_header INNER JOIN " & _
                           "backflush_line on backflush_header.IDTrans = backflush_line.IDTrans " & _
                      "WHERE (LEFT([backflush_Header].IDTrans, 2) = N'FG') AND (MONTH(DateTrans) = " & PeriodeData & ")" & _
                      "GROUP BY backflush_header.IDTrans,backflush_header.validasi, backflush_header.dateTrans , " & _
                               "backflush_line.cost,backflush_header.IDTrans, backflush_header.OrderID", CNN, lckLockBatch
            
            OpenList "BFFG"
            
End Select
TampilanList Rc.DBRecordset
End Sub

Private Function CekType(ByVal FieldData As Field, ByVal ListValue As Variant) As Variant
If IsDate(ListValue) = True Then
   CekType = Format(IIf(Not IsNull(ListValue), ListValue, Date), ShortDateFormGaris)
ElseIf IsNumeric(ListValue) = True Then
   CekType = FormatNumber(IIf(Not IsNull(ListValue), ListValue, 0), 0)
Else
   CekType = IIf(Not IsNull(ListValue), ListValue, "-")
End If
End Function

Private Sub CreatePeriode()
Dim I As Integer
cboPeriode(0).Clear
For I = 1 To 12
    cboPeriode(0).AddItem Format(DateSerial(Year(Date), I, 1), "MMMM")
    cboPeriode(0).ItemData(cboPeriode(0).NewIndex) = I
Next
cboPeriode(0).ListIndex = mVarPeriode - 1
End Sub

Private Sub TanggalPeriode()
Dim RcTglData As New DBQuick
Dim Avdata As Variant
Dim I As Integer
RcTglData.DBOpen StrTanggalPeriode, CNN, lckLockReadOnly
cboPeriode(1).Clear
cboPeriode(1).AddItem "Semua Tanggal"
With RcTglData.DBRecordset
     If .Recordcount <> 0 Then
         Avdata = .Getrows(.Recordcount, adBookmarkFirst)
         For I = 0 To UBound(Avdata, 2)
             cboPeriode(1).AddItem Format(Avdata(0, I), "dd/mmmm/yyyy")
         Next
     End If
     .Close
End With
cboPeriode(1).ListIndex = 0
Set Avdata = Nothing
Set RcTglData = Nothing
End Sub

Private Sub FilterTanggalPeriode()
If Rc.DBRecordset.State = 1 Then
   With Rc.DBRecordset
        If .Recordcount <> 0 Then
           If cboPeriode(1).Text <> "Semua Tanggal" Then
              .Filter = adFilterNone
              .Filter = "Tanggal ='" & cboPeriode(1).Text & "'"
           Else
              .Filter = adFilterNone
           End If
        End If
        TampilanList Rc.DBRecordset
   End With
End If
End Sub

Private Sub TampilanList(ByVal RecordSetData As Recordset)
On Error Resume Next
Dim I As Long
Dim j As Long
Dim Fld As Field
Dim Ftes As String
Dim Avdata As Variant
ListView1.ListItems.Clear
If RecordSetData.State = 1 Then
    With RecordSetData
         If .Recordcount <> 0 Then
            Avdata = .Getrows(.Recordcount, adBookmarkFirst)
            For I = 0 To UBound(Avdata, 2)
                If Avdata(0, I) = 0 Then Ftes = "NO" Else Ftes = "VALIDASI"
                With ListView1.ListItems.Add(, , Ftes)
                     If Ftes = "VALIDASI" Then .Checked = True
                     j = 0
                     For Each Fld In RecordSetData.Fields
                         .SubItems(j + 1) = CekType(Fld, Avdata(j + 1, I))
                         j = j + 1
                     Next
                End With
            Next
            ListView1.SelectedItem.Index = 1
         End If
    End With
End If
Set Fld = Nothing
Set Avdata = Nothing
End Sub

Private Sub SimpanValidasi(ByVal NilaiValidasi As Boolean)
Select Case TreeView1.SelectedItem.Key

'=========== Edit ==========================================
      'Case "Order pembelian", "Order penjualan":
'===========================================================

       Case "Order pembelian", "Petty", "Order penjualan":
            If SendDataToServer("UPDATE  TransData SET Validasi =" & BoolToInt(NilaiValidasi) & " WHERE     (TransID = N'" & mVarIndex & "')") = True Then
               SendDataToServer (" UPDATE [PO Order] SET StatusSJ = " & BoolToInt(NilaiValidasi) & " WHERE     (PurchaseID = N'" & ListView1.SelectedItem.SubItems(3) & "') ")
               If NilaiValidasi = True Then CreateJournalValidasi mVarIndex
            End If
       Case "Retur pembelian", "Retur penjualan":
            If SendDataToServer("UPDATE    Returdata SET Validasi =" & BoolToInt(NilaiValidasi) & " WHERE     (ReturID = N'" & mVarIndex & "')") = True Then If NilaiValidasi = True Then CreateJournalValidasi mVarIndex
       Case "Pembelian Aktiva", "Penjualan Aktiva":
            If SendDataToServer("UPDATE    [TR Aktiva Tetap] SET Validasi =" & BoolToInt(NilaiValidasi) & " WHERE     ([No FA] = N'" & mVarIndex & "')") = True Then If NilaiValidasi = True Then CreateJournalValidasi mVarIndex
       Case "Pelunasan Hutang", "Pelunasan Piutang":
            If SendDataToServer("UPDATE    [TransData] SET Validasi =" & BoolToInt(NilaiValidasi) & " WHERE     ([TransID] = N'" & mVarIndex & "')") = True Then If NilaiValidasi = True Then CreateJournalValidasi mVarIndex
End Select
End Sub

Private Sub CreateJournalValidasi(ByVal TransID As String)
Dim RcJournal As New DBQuick
Dim Avdata As Variant
Dim mVarTotal As Variant
Dim mVarPPn As Variant
Dim mVarBarang As Variant
Dim mVarDP As Variant
Dim mVarSubtotal As Variant
Dim mVarNilaiBuku As Variant
Dim mTot As Variant
Dim StrRef As String
Dim StrType As String
Dim StrPartic As String
Dim StrKas As String
Dim I As Integer
Dim j As Integer
Dim k As Integer

Select Case TreeView1.SelectedItem.Key
       Case "Order pembelian":
            RcJournal.DBOpen " SELECT TransData.TransID, [PO Order].PurchaseID, [PO Order].PartnerID, TransData.EmpID,TransData.DateTrans, [Detail PO].CurID, [PO Order].Kurs,[Detail PO].NoItem, [Detail PO].QTYPO * [Detail PO].POPrice AS [Sub Total], [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2) AS Discount, " & _
                             " ([Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2))  * ROUND([Detail PO].VAT / 100, 2) AS PPN FROM TransData INNER JOIN  [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN  [Detail PO] ON [PO Order].PurchaseID = [Detail PO].PurchaseID WHERE     (TransData.TransID = N'" & TransID & "') ORDER BY TransData.TransID,[Detail PO].NoItem", CNN, lckLockReadOnly
            StrType = "BPBK"
            j = 0
       
'============================================== Tambahan ==========================
        Case "Petty":
            RcJournal.DBOpen " SELECT TransData.TransID, [PO Order].PurchaseID, [PO Order].PartnerID, TransData.EmpID,TransData.DateTrans, [Detail PO].CurID, [PO Order].Kurs,[Detail PO].NoItem, [Detail PO].QTYPO * [Detail PO].POPrice AS [Sub Total], [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2) AS Discount, " & _
                             " ([Detail PO].QTYPO * [Detail PO].POPrice - [Detail PO].QTYPO * [Detail PO].POPrice * ROUND([PO Order].Discount / 100, 2))  * ROUND([Detail PO].VAT / 100, 2) AS PPN FROM TransData INNER JOIN  [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN  [Detail PO] ON [PO Order].PurchaseID = [Detail PO].PurchaseID WHERE     (TransData.TransID = N'" & TransID & "') ORDER BY TransData.TransID,[Detail PO].NoItem", CNN, lckLockReadOnly
            StrType = "BPPC"
            j = 0
 
'==================================================================================
       
       Case "Retur pembelian":
            RcJournal.DBOpen " SELECT ReturData.ReturID AS TransID, ReturData.TransID AS PurchaseID, TransData.PartnerId AS PartnerId, ReturData.EmpID, ReturData.DateTrans, " & _
                             " TransData.CurrID, TransData.Kurs, [Detail Retur].NoItem, [Detail Retur].[Retur Beli] * [Detail Retur].Price AS [Sub Total], [Detail Retur].[Retur Beli] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2) AS Discount,  ([Detail Retur].[Retur Beli] * [Detail Retur].Price - [Detail Retur].[Retur Beli] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2))   * ROUND([Detail Retur].VAT / 100, 2) AS PPN FROM         ReturData INNER JOIN TransData ON ReturData.TransID = TransData.TransID INNER JOIN [Detail Retur] ON ReturData.ReturID = [Detail Retur].ReturID WHERE     (ReturData.ReturID = N'" & TransID & "') ORDER BY ReturData.DateTrans, ReturData.ReturID", CNN, lckLockReadOnly
            StrPartic = "Retur Pembelian "
            StrType = "BRPB"
            j = 4
            
            
            
       Case "Order penjualan":
            RcJournal.DBOpen " SELECT TransData.TransID, [PO Order].PurchaseID, [PO Order].PartnerID, TransData.EmpID, TransData.DateTrans, [PO Order].CurrID, [PO Order].Kurs, [Detail TransData].NoItem, [Detail TransData].QTY_OUT * [Detail TransData].Price AS [Sub Total], [Detail TransData].QTY_OUT * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2) AS Discount, ([Detail TransData].QTY_OUT * [Detail TransData].Price - [Detail TransData].QTY_OUT * [Detail TransData].Price * ROUND(TransData.Discount / 100, 2))  * ROUND([Detail TransData].VAT / 100, 2) AS PPN" & _
                             " FROM TransData INNER JOIN [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE     (TransData.TransID = N'" & TransID & "') ORDER BY TransData.TransID, [Detail TransData].NoItem", CNN, lckLockReadOnly
            StrType = "BPJK"
            j = 0
       Case "Retur penjualan":
            RcJournal.DBOpen " SELECT ReturData.ReturID AS TransID, TransData.PurchaseID AS PurchaseID, TransData.PartnerId AS PartnerId, ReturData.EmpID, ReturData.DateTrans, " & _
                             " TransData.CurrID, TransData.Kurs, [Detail Retur].NoItem, [Detail Retur].[Retur Jual] * [Detail Retur].Price AS [Sub Total], [Detail Retur].[Retur Jual] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2) AS Discount,  ([Detail Retur].[Retur Jual] * [Detail Retur].Price - [Detail Retur].[Retur Jual] * [Detail Retur].Price * ROUND(TransData.Discount / 100, 2))   * ROUND([Detail Retur].VAT / 100, 2) AS PPN,[Detail Retur].[Retur Jual] , [Detail Retur].Price FROM         ReturData INNER JOIN TransData ON ReturData.TransID = TransData.TransID INNER JOIN [Detail Retur] ON ReturData.ReturID = [Detail Retur].ReturID WHERE     (ReturData.ReturID = N'" & TransID & "') ORDER BY ReturData.DateTrans, ReturData.ReturID", CNN, lckLockReadOnly
            StrType = "BRPJ"
            j = 0


       Case "Pembelian Aktiva":
            RcJournal.DBOpen " SELECT     [TR Aktiva Tetap].[No FA] AS TransID, '-' AS PurchaseID, [TR Aktiva Tetap].PartnerID, '-' AS EmpID, [TR Aktiva Tetap].DateTrans, 'IDR' AS [Mata Uang], 1 AS Kurs, [DTR Aktiva Tetap].[No Aktiva] AS NoItem,  [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga AS [Sub Total], [TR Aktiva Tetap].DP AS Discount, " & _
                             "  [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS PPN,                       [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga + [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS Total, [DTR Aktiva Tetap].[Aktiva Beli], [DTR Aktiva Tetap].Harga,[TR Aktiva Tetap].[Id Group], [TR Aktiva Tetap].BankID FROM [TR Aktiva Tetap] INNER JOIN  [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].[No FA] = N'" & TransID & "')", CNN, lckLockReadOnly
            StrType = "BKKAT"
            j = 0
       Case "Penjualan Aktiva":
            RcJournal.DBOpen " SELECT     [TR Aktiva Tetap].[No FA] AS TransID, '-' AS PurchaseID, [TR Aktiva Tetap].PartnerID, '-' AS EmpID, [TR Aktiva Tetap].DateTrans, 'IDR' AS [Mata Uang], 1 AS Kurs, [DTR Aktiva Tetap].[No Aktiva] AS NoItem,  [DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga AS [Sub Total], [TR Aktiva Tetap].DP AS Discount, " & _
                             "  [DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS PPN, [DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga + [DTR Aktiva Tetap].[Aktiva Jual] * [DTR Aktiva Tetap].Harga * ROUND([DTR Aktiva Tetap].PPn / 100, 2) AS Total, [DTR Aktiva Tetap].[Aktiva Jual], [DTR Aktiva Tetap].Harga,[TR Aktiva Tetap].[Id Group], [TR Aktiva Tetap].BankID,  [DTR Aktiva Tetap].[Doc Reff],[DTR Aktiva Tetap].[Harga Jual] FROM [TR Aktiva Tetap] INNER JOIN  [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].[No FA] = N'" & TransID & "')", CNN, lckLockReadOnly
            StrType = "BKMAT"
            j = 0
       Case "Pelunasan Hutang":
            RcJournal.DBOpen " SELECT     TransData.TransID, TransData.PurchaseID, TransData.PartnerId, '-' AS EmpID, TransData.DateTrans, TransData.CurrID AS [Mata Uang], TransData.Kurs,  TransData.PartnerId AS NoItem, [Detail TransData].Credit AS [Sub Total], 0 AS Discount, [Detail TransData].Credit AS Total,TransData.BankID FROM  TransData INNER JOIN [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE     (TransData.TransID = N'" & TransID & "')", CNN, lckLockReadOnly
            StrType = "BKK"
            j = 0
       Case "Pelunasan Piutang":
            RcJournal.DBOpen " SELECT     TransData.TransID, TransData.PurchaseID, TransData.PartnerId, '-' AS EmpID, TransData.DateTrans, TransData.CurrID AS [Mata Uang],  TransData.Kurs, TransData.PartnerId AS NoItem, [Detail TransData].Debet AS [Sub Total], 0 AS Discount, [Detail TransData].Debet AS Total,  TransData.BankID FROM  TransData INNER JOIN  [Detail TransData] ON TransData.TransID = [Detail TransData].TransID WHERE     (TransData.TransID = N'" & TransID & "')", CNN, lckLockReadOnly
            StrType = "BKM"
            j = 0
End Select
StrPartic = TreeView1.SelectedItem.Key
With RcJournal.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        If CreateHeaderValidasi(Avdata(0, I), Avdata(1, I), Avdata(2, I), "Admin", Avdata(4, I), StrType, Avdata(6, I), Avdata(5, I)) = True Then
            mVarTotal = 0
            mVarPPn = 0
            mVarBarang = 0
            mVarSubtotal = 0
            StrRef = ""
            For I = 0 To UBound(Avdata, 2)
                StrRef = Avdata(2, I)
                j = j + 1
                mVarTotal = mVarTotal + (Avdata(8, I) - Avdata(9, I)) + Avdata(10, I)
                mVarPPn = mVarPPn + Avdata(10, I)
                mVarDP = mVarDP + CCur(Avdata(9, I))
                mVarSubtotal = mVarSubtotal + Avdata(8, I)
                StrPartic = StrPartic & "," & Avdata(7, I)
                

                Select Case TreeView1.SelectedItem.Key
'=================================Tambahan ======================================
                   'Case "Order pembelian", "Order penjualan":
'================================================================================
                    Case "Order pembelian", "Petty", "Order penjualan":
                        mVarBarang = mVarBarang + Avdata(8, I) - Avdata(9, I)
                        DetailValidasi CariAkunItem(Avdata(7, I)), Avdata(7, I), CCur(Avdata(8, I) - Avdata(9, I)), 0, "Persediaan " & Avdata(7, I), j
                   
                   
                   Case "Retur pembelian":
                        mVarBarang = mVarBarang + Avdata(8, I) - Avdata(9, I)
                        DetailValidasi CariAkunItem(Avdata(7, I)), Avdata(7, I), 0, CCur(Avdata(8, I) - Avdata(9, I)), "Persediaan " & Avdata(7, I), j
                   Case "Retur penjualan":
                        DetailValidasi CariAkunItem(Avdata(7, I)), Avdata(7, I), CCur(Avdata(8, I) - Avdata(9, I)), 0, "Persediaan " & Avdata(7, I), j
                        'HPP
                        StrPartic = StrPartic & ",HPP " & Avdata(7, I)
                        j = j + 1
                        mVarBarang = mVarBarang + Avdata(11, I) * HppProce(Avdata(1, I), Avdata(7, I))
                        DetailValidasi CariTypeJournal(23), Avdata(7, I), 0, Avdata(11, I) * HppProce(Avdata(1, I), Avdata(7, I)), "HPP " & Avdata(7, I), j
                   Case "Pembelian Aktiva":
                        StrKas = Avdata(11, I)
                        CariAkumulasi Avdata(7, 0), Avdata(16, 0)
                        CariNoAccountDepre Avdata(16, 0), Avdata(7, 0)
                        StrKas = Avdata(15, I)
                        DetailValidasi Avdata(14, I), Avdata(7, I), CCur(Avdata(8, I)), 0, "Harga Perolehan " & Avdata(7, I), j
                   Case "Penjualan Aktiva":
                        StrKas = Avdata(11, I)
                        'StrKas = Avdata(15, i)
                        CariAkumulasi Avdata(7, 0), Avdata(16, 0)
                        CariNoAccountDepre Avdata(16, 0), Avdata(7, 0)
                        
                        DetailValidasi Avdata(14, I), "xxx", CCur(Avdata(8, I) + Avdata(10, I)), 0, "Harga Perolehan " & Avdata(7, I), j
                        DetailValidasi AkumAccount, "xxx", TotalAkum, 0, "Harga Perolehan " & Avdata(7, I), j
                        mVarNilaiBuku = Avdata(13, I) - TotalAkum
                        mTot = Avdata(8, I) - mVarNilaiBuku
                        If mTot < 0 Then mTot = mTot * (-1)
                        Select Case mVarNilaiBuku
                               Case Is = Avdata(8, I):
                                    'SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & CariTypeAccount("Rugi Penjualan/Penukaran Aktiva") & "', N'" & .Fields("No Aktiva") & "', 0, 0, N'Rugi Penjualan/Penukaran Aktiva" & Left(.Fields("Nama Aktiva"), 200) & "')")
                               Case Is > Avdata(8, I):
                                    DetailValidasi CariTypeJournal(55), Avdata(7, I), TotalAkum, 0, "Harga Perolehan " & Avdata(7, I), j
                               Case Is < Avdata(8, I):
                                    DetailValidasi CariTypeJournal(54), Avdata(7, I), TotalAkum, 0, "Harga Perolehan " & Avdata(7, I), j
                        End Select
                End Select
            Next I
            
            'PPN Masukan
            k = k + (j + 1)
            Select Case TreeView1.SelectedItem.Key
                   Case "Order pembelian":
                        j = j + 1
                        StrPartic = StrPartic & ",PPN masukan"
                        DetailValidasi CariTypeJournal(42), "xxx", CCur(mVarPPn), 0, "PPN masukan", j
'============================================================= Tambahan ===================================
                   'Case "Petty":
                   '     j = j + 1
                   '     StrPartic = StrPartic & ",PPN masukan"
                   '     DetailValidasi CariTypeJournal(42), "xxx", CCur(mVarPPn), 0, "PPN masukan", j
'==========================================================================================================
                   Case "Order penjualan":
                        j = j + 1
                        StrPartic = StrPartic & ",PPN keluaran"
                        DetailValidasi CariTypeJournal(41), "xxx", CCur(mVarPPn), 0, "PPN keluaran", j
                   Case "Retur pembelian":
                        j = j + 1
                        StrPartic = StrPartic & ",PPN masukan"
                        DetailValidasi CariTypeJournal(42), "xxx", 0, CCur(mVarPPn), "PPN masukan", 3
                   Case "Retur penjualan":
                        j = j + 1
                        StrPartic = StrPartic & ",PPN keluaran"
                        DetailValidasi CariTypeJournal(41), "xxx", CCur(mVarPPn), 0, "PPN keluaran", k + 1
                   Case "Pembelian Aktiva":
                        j = j + 1
                        StrPartic = StrPartic & ",PPN masukan"
                        DetailValidasi CariTypeJournal(42), "xxx", CCur(mVarPPn), 0, "PPN masukan", j
                   Case "Penjualan Aktiva":
                        j = j + 1
                        StrPartic = StrPartic & ",PPN keluaran"
                        DetailValidasi CariTypeJournal(41), "xxx", CCur(mVarPPn), 0, "PPN keluaran", j
                        
            End Select
            'Hutang Usaha
            
            Select Case TreeView1.SelectedItem.Key
                   Case "Order pembelian":
                        j = j + 1
                        StrPartic = StrPartic & ",hutang usaha ke " & StrRef
                        DetailValidasi CariTypeJournal(28), StrRef, 0, CCur(mVarTotal), "Hutang Usaha Ke " & StrRef, j
                   
                   '================================================ Tambahan ========================
                   Case "Petty":
                        j = j + 1
                        StrPartic = StrPartic & ",hutang usaha ke " & StrRef
                        DetailValidasi CariTypeJournal(51), StrRef, 0, CCur(mVarTotal), "Hutang Usaha Ke " & StrRef, j
                   '==================================================================================
                   
                   Case "Order penjualan":
                        j = j + 1
                        StrPartic = StrPartic & ",piutang usaha ke " & StrRef
                        DetailValidasi CariTypeJournal(39), StrRef, 0, CCur(mVarTotal), "Piutang usaha ke " & StrRef, j
                   Case "Retur pembelian":
                        j = j + 1
                        StrPartic = StrPartic & ",Hutang usaha ke " & StrRef
                        DetailValidasi CariTypeJournal(28), StrRef, CCur(mVarTotal), 0, "Hutang Usaha Ke " & StrRef, 1
                   Case "Retur penjualan":
                        StrPartic = StrPartic & ",piutang usaha ke " & StrRef
                        DetailValidasi CariTypeJournal(39), StrRef, 0, CCur(mVarTotal), "Piutang Usaha Ke " & StrRef, k + 2
                   Case "Pembelian Aktiva":
                        If mVarDP <> 0 Then
                          If CCur(mVarDP) >= CCur(mVarSubtotal) Then
                             StrPartic = StrPartic & ",Pemb. tunai dr " & StrRef
                             DetailValidasi StrKas, StrRef, 0, CCur(mVarSubtotal + mVarPPn), "Pemb. tunai dr " & StrRef, j + 1
                          Else
                             StrPartic = StrPartic & ",DP Pemb Aktiva " & StrRef
                             DetailValidasi StrKas, StrRef, 0, CCur(mVarDP), "DP Pemb Aktiva " & StrRef, k + 2
                             
                             StrPartic = StrPartic & ",Hutang pemb. aktiva ke " & StrRef
                             DetailValidasi CariTypeJournal(57), StrRef, 0, CCur((mVarSubtotal - mVarDP) + mVarPPn), "Hutang pemb. aktiva ke " & StrRef, k + 2
                          End If
                        Else
                           StrPartic = StrPartic & ",Hutang pemb. aktiva ke " & StrRef
                           DetailValidasi CariTypeJournal(57), StrRef, 0, CCur(mVarTotal), "Hutang pemb. aktiva ke " & StrRef, k + 2
                        End If
                   Case "Penjualan Aktiva":
                        If mVarDP <> 0 Then
                          If CCur(mVarDP) >= CCur(mVarSubtotal) Then
                             StrPartic = StrPartic & ",Pemb. tunai dr " & StrRef
                             DetailValidasi StrKas, StrRef, 0, CCur(mVarSubtotal + mVarPPn), "Pemb. tunai dr " & StrRef, j + 1
                          Else
                             StrPartic = StrPartic & ",DP Pemb Aktiva " & StrRef
                             DetailValidasi StrKas, StrRef, 0, CCur(mVarDP), "DP Pemb Aktiva " & StrRef, k + 2
                             
                             StrPartic = StrPartic & ",Hutang pemb. aktiva ke " & StrRef
                             DetailValidasi CariTypeJournal(57), StrRef, 0, CCur((mVarSubtotal - mVarDP) + mVarPPn), "Hutang pemb. aktiva ke " & StrRef, k + 2
                          End If
                        Else
                           StrPartic = StrPartic & ",Piutang Penj. aktiva ke " & StrRef
                           DetailValidasi CariTypeJournal(56), StrRef, 0, CCur(mVarTotal), "Piutang Penj. aktiva ke " & StrRef, k + 2
                        End If
                   Case "Pelunasan Hutang":
                        j = j + 1
                        StrPartic = StrPartic & ",Pelunasan Hutang usaha ke " & StrRef
                        DetailValidasi CariTypeJournal(28), StrRef, CCur(mVarTotal), 0, "Pelunasan Hutang usaha ke  " & StrRef, 1
                        StrPartic = StrPartic & ",Pembayaran Hutang usaha ke " & StrRef
                        DetailValidasi StrKas, StrRef, 0, CCur(mVarTotal), "Pembayaran Hutang usaha ke  " & StrRef, 2
                   Case "Pelunasan Piutang":
                        j = j + 1
                        StrPartic = StrPartic & ",Penerimaan Piutang usaha Dari " & StrRef
                        DetailValidasi CariTypeJournal(39), StrRef, CCur(mVarTotal), 0, "Penerimaan Piutang usaha Dari " & StrRef, 1
                        StrPartic = StrPartic & ",Pembayaran piutang usaha dari " & StrRef
                        DetailValidasi StrKas, StrRef, 0, CCur(mVarTotal), "Pembayaran piutang usaha dari " & StrRef, 2
                        
            End Select
            'Retur Pembelian
            
            Select Case TreeView1.SelectedItem.Key
                   Case "Retur pembelian":
                        j = j + 1
                        StrPartic = StrPartic & ",Retur pembelian"
                        DetailValidasi CariTypeJournal(58), StrRef, 0, CCur(mVarBarang), "Retur pembelian", 2
                        StrPartic = StrPartic & ",Retur pembelian"
                        DetailValidasi CariTypeJournal(58), StrRef, CCur(mVarBarang), 0, "Retur pembelian", 4
                   Case "Retur penjualan":
                        StrPartic = StrPartic & ",Retur penjualan"
                        DetailValidasi CariTypeJournal(43), StrRef, CCur(mVarBarang), 0, "Retur penjualan", k
                        
            End Select
            SendDataToServer (" UPDATE [Table Journal] SET RefNotes = '" & Left(StrPartic, 249) & "'  WHERE     (JournalID = N'" & mVarNoUrutJournal & "') ")
        End If
     End If
End With
Set Avdata = Nothing
End Sub

Private Function CreateHeaderValidasi(ByVal pTransID As String, _
                                      ByVal pPurchaseID As String, _
                                      ByVal pPartnerId As String, _
                                      ByVal pEmpID As String, _
                                      ByVal pDateTrans As Date, _
                                      ByVal pTypeTrans As String, _
                                      ByVal pKurs As Currency, _
                                      Optional ByVal pCurrency As String = "IDR") As Boolean
Dim I As Integer
I = 0
mVarNoUrutJournal = ""
    Do
Ulang:
      I = I + 1
      mVarNoUrutJournal = NoUrutJournal
      If CekKode(mVarNoUrutJournal) = False Then
         CreateHeaderValidasi = SendDataToServer(" INSERT INTO [Table Journal] " & _
                                                 " (JournalID, TransID, PurchaseID, PartnerID, Currency,Kurs, EmpID, DateTrans,  TypeTrans,  NoUrut, Periode)" & _
                                                 " VALUES (N'" & mVarNoUrutJournal & "', N'" & pTransID & "', N'" & pPurchaseID & "', N'" & pPartnerId & "', N'" & pCurrency & "'," & CDbl(pKurs) & ", N'Admin', CONVERT(DATETIME, '" & Format(pDateTrans, "dd/mm/yy") & "', 3),  N'" & pTypeTrans & "', N'" & mVarNoUrutJournal & "', " & mVarPeriode & ")")
         If CreateHeaderValidasi = False Then
            If I <= 50 Then GoTo Ulang
         Else
            Exit Do
         End If
         
      End If
    Loop
End Function

Private Sub DetailValidasi(ByVal pNoAccount As String, _
                           ByVal pDocReff As String, _
                           ByVal pDebet As Variant, _
                           ByVal pCredit As Variant, _
                           ByVal pKeterangan As String, _
                           ByVal pNo As Integer)
SendDataToServer (" INSERT INTO [Detail Journal]" & _
                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan, [No])" & _
                  " VALUES (N'" & mVarNoUrutJournal & "', N'" & pNoAccount & "', N'" & pDocReff & "', " & pDebet & ", " & pCredit & ", N'" & pKeterangan & "', " & pNo & ")")

End Sub


Private Function NoUrutJournal() As String
Dim RcIdx As New DBQuick
Dim mThn As String
Dim NoJrl As String
Dim mVarNo As Long
NoJrl = "JR" & "-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
RcIdx.DBOpen "SELECT MAX(RIGHT([JournalID], 5)) AS MaxNOm FROM [Table Journal]", CNN, lckLockReadOnly
With RcIdx.DBRecordset
     If .Recordcount <> 0 Then
        If IsNumeric(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) Then
            mVarNo = CDbl(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) + 1
        Else
            mVarNo = Val(IIf(Not IsNull(.Fields(0)), .Fields(0), 0)) + 1
        End If
        'mThn = IIf(Not IsNull(.Fields(1)), .Fields(1), Format(Year(dDateBegin), "0###"))
     Else
        mVarNo = 1
     End If
     NoUrutJournal = NoJrl & KirimNull(5 - Len(Trim(Str(mVarNo)))) & Trim(Str(mVarNo))
End With
End Function

Private Function CekKode(ByVal CariKode As String) As Boolean
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT  JournalID FROM [Table Journal] WHERE (JournalID = N'" & CariKode & "')", CNN, lckLockReadOnly
If RcCek.DBRecordset.Recordcount <> 0 Then CekKode = True
RcCek.CloseDB
Set RcCek = Nothing
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

Private Sub CariAkumulasi(ByVal KodeAktiva As String, ByVal NoFa As String)
Dim RcAkum As New DBQuick
'MsgBox "SELECT     SUM([Detail Journal].Debet) AS Akumulasi, [Detail Journal].NoAccount FROM         [Table Journal] INNER JOIN                       [TR Aktiva Tetap] ON [Table Journal].TransID = [TR Aktiva Tetap].[No FA] INNER JOIN                       [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([TR Aktiva Tetap].Disposal = 0) AND ([Table Journal].TypeTrans = N'AKDEP') AND ([TR Aktiva Tetap].[No FA] = N'" & NoFa & "') AND                        ([Detail Journal].[Doc Reff] = N'" & KodeAktiva & "') GROUP BY [Detail Journal].NoAccount HAVING      (SUM([Detail Journal].Debet) <> 0)"
RcAkum.DBOpen "SELECT     SUM([Detail Journal].Debet) AS Akumulasi, [Detail Journal].NoAccount FROM         [Table Journal] INNER JOIN                       [TR Aktiva Tetap] ON [Table Journal].TransID = [TR Aktiva Tetap].[No FA] INNER JOIN                       [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([TR Aktiva Tetap].Disposal = 0) AND ([Table Journal].TypeTrans = N'AKDEP') AND ([TR Aktiva Tetap].[No FA] = N'" & NoFa & "') AND ([Detail Journal].[Doc Reff] = N'" & KodeAktiva & "') GROUP BY [Detail Journal].NoAccount HAVING      (SUM([Detail Journal].Debet) <> 0)", CNN, lckLockReadOnly
TotalAkum = 0
AkumAccount = ""
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        TotalAkum = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        AkumAccount = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
     Else
        CariNoAccountDepre NoFa, KodeAktiva
        TotalAkum = 0
        AkumAccount = mVarAccDep
     End If
End With
End Sub

Private Sub CariNoAccountDepre(ByVal NoTranBeli As String, ByVal NoFixAssets As String)
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     [TR Aktiva Tetap].AccDep, [TR Aktiva Tetap].DepAktiva FROM         [TR Aktiva Tetap] INNER JOIN                       [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].[No FA] = N'" & NoTranBeli & "') AND ([DTR Aktiva Tetap].[No Aktiva] = N'" & NoFixAssets & "')", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        mVarAccDep = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
        mVarDepre = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
     End If
End With
End Sub

Private Sub OpenList(ByVal FilterData As String)
On Error Resume Next
Dim I As Integer
Dim Avdata As Variant
Dim Ftes  As String
ListView2.ListItems.Clear
RcList.DBOpen "Shape{SELECT Status AS Closing, JournalID AS [No Journal], TransID AS [Bukti Transaksi], " & _
        " DateTrans AS [Tanggal Transaksi], RefNotes, TypeTrans FROM [Table Journal] " & _
        " WHERE (Periode = " & cboPeriode(0).ListIndex + 1 & ") and TypeTrans =N'" & FilterData & "' " & _
        " ORDER BY JournalID} as ParentMenu append ({SELECT     [Detail Journal].JournalID AS [No Journal], " & _
        " [Detail Journal].NoAccount AS [Kode Perkiraan], GLAccount.AccountName AS [Nama Perkiraan], " & _
        " [Detail Journal].Keterangan, [Detail Journal].[Doc Reff], [Detail Journal].Debet, [Detail Journal].Credit " & _
        " FROM [Detail Journal] INNER JOIN GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount " & _
        " ORDER BY [Detail Journal].[No]} as ChildMd relate [No Journal] to [No Journal]) ", CNN, lckLockBatch
'Debug.Print RcList.DBRecordset.Source
With RcList.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            If Avdata(0, I) = 0 Then Ftes = "NO" Else Ftes = "CLOSING"
            With ListView2.ListItems.Add(, , Ftes)
                 If Ftes = "CLOSING" Then .Checked = True
                 .SubItems(1) = Avdata(1, I)
                 .SubItems(2) = Avdata(2, I)
                 .SubItems(3) = Avdata(3, I)
                 .SubItems(4) = Avdata(4, I)
                 .SubItems(5) = Avdata(5, I)
            End With
        Next I
     End If
End With
Set Avdata = Nothing
End Sub

'Tranfer Journal

Private Function AccountLink() As String
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     NoAccount FROM         [Tabel Pembantu] WHERE     ([Seting Relasi] = 1)", CNN, lckLockReadOnly
AccountLink = "xxx"
With Rc
     If .Recordcount <> 0 Then
        AccountLink = IIf(Not IsNull(.Fields(0)), .Fields(0), "xxx")
     End If
End With

End Function

Private Sub PrepareJournalFixAssets()
Dim AccJournal As New DBQuick
Dim mVarData As Variant
Dim mVarI As Integer
Dim Kodeku As String
AccJournal.DBOpen " SELECT [TR Aktiva Tetap].DateTrans AS [Tanggal Bukti], [TR Aktiva Tetap].[No FA] AS [No Bukti], [DTR Aktiva Tetap].[No Aktiva] AS [Kode Aktiva],                       [TR Aktiva Tetap].DepAktiva, [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga / [TR Aktiva Tetap].Umur AS [Perolehan Aktiva],                       [TR Aktiva Tetap].BankID AS [Kode Kas], [TR Aktiva Tetap].AccDep,                       [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga / [TR Aktiva Tetap].Umur AS [Kas Keluar]" & _
                  " FROM [TR Aktiva Tetap] INNER JOIN                       [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].Periode = " & mVarPeriode & ")", CNN, lckLockReadOnly
With AccJournal.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For mVarI = 0 To UBound(mVarData, 2)
            Kodeku = IdxAuto
            If SendDataToServer(" INSERT INTO [Table Journal]" & _
                                " (JournalID,TransID, DateTrans, Periode, TypeTrans, RefNotes,Status) " & _
                                " VALUES     (N'" & Kodeku & "',N'" & mVarData(1, mVarI) & "', CONVERT(DATETIME, '" & Format(mVarData(0, mVarI), "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'AKDEP', N'Akumulasi Penyusutan Aktiva',1)") = True Then
                                
               SendDataToServer " INSERT INTO [Detail Journal]" & _
                                " (JournalID, NoAccount,  Debet, Credit, Keterangan) " & _
                                " VALUES     (N'" & Kodeku & "', N'" & mVarData(6, mVarI) & "'," & mVarData(4, mVarI) & ",0,N'Akumulasi Depresiasi " & mVarData(4, mVarI) & "')"
                                
               SendDataToServer " INSERT INTO [Detail Journal]" & _
                                " (JournalID, [Doc Reff], NoAccount, Debet, Credit, Keterangan) " & _
                                " VALUES     (N'" & Kodeku & "',N'" & mVarData(2, mVarI) & "', N'" & mVarData(3, mVarI) & "',0," & mVarData(4, mVarI) & ",N'Akumulasi Depresiasi " & mVarData(4, mVarI) & "')"
                                
            End If
            mVarJournal = True
        Next mVarI
     End If
End With
End Sub

Private Sub Closing(ByVal PeriodeActive As Integer)
Dim rcSet As New DBQuick
Dim I As Long
Dim mPer As Integer
Dim mBook As Long
Dim mVarData As Variant
Select Case PeriodeActive
       Case 1: mPer = 0
       Case 2: mPer = 1
       Case 3: mPer = 2
       Case 4: mPer = 3
       Case 5: mPer = 4
       Case 6: mPer = 5
       Case 7: mPer = 6
       Case 8: mPer = 7
       Case 9: mPer = 8
       Case 10: mPer = 9
       Case 11: mPer = 10
       Case 12: mPer = 11
End Select
IsiLabaRugi PeriodeActive
'MsgBox "SELECT  [Detail Journal].NoAccount, GLAccount.Period" & mPer & " AS [Saldo Awal], SUM([Detail Journal].Debet) AS Debet, SUM([Detail Journal].Credit) AS Kredit,  ABS(GLAccount.Period" & mPer & " + SUM([Detail Journal].Debet) - SUM([Detail Journal].Credit)) AS Balance FROM         [Detail Journal] INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([Table Journal].Periode = " & PeriodeActive & ") GROUP BY [Detail Journal].NoAccount, GLAccount.Period" & mPer & " ORDER BY [Detail Journal].NoAccoun"
rcSet.DBOpen "SELECT  [Detail Journal].NoAccount, GLAccount.Period" & mPer & " AS [Saldo Awal], SUM([Detail Journal].Debet) AS Debet, SUM([Detail Journal].Credit) AS Kredit,  ABS(GLAccount.Period" & mPer & " + SUM([Detail Journal].Debet) - SUM([Detail Journal].Credit)) AS Balance FROM         [Detail Journal] INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([Table Journal].Periode = " & PeriodeActive & ") GROUP BY [Detail Journal].NoAccount, GLAccount.Period" & mPer & " ORDER BY [Detail Journal].NoAccount", CNN, lckLockReadOnly
With rcSet.DBRecordset
     If .Recordcount <> 0 Then
        CompareAccount
        IsiListDataDetail
     End If
End With
SendDataToServer ("Delete from [Table Journal] where JournalID ='LINK'")
'rcSet.CloseDB
End Sub

Private Sub IsiLabaRugi(ByVal PeriodeData As Integer)
Dim Rc As New DBQuick
Dim mVarJrl As New clsJournal
Rc.DBOpen "SELECT ABS(SUM([Detail Journal].Debet - [Detail Journal].Credit)) AS Debet FROM  [Detail Journal] INNER JOIN GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID WHERE     ([Tabel Pembantu].[Kelompok Perkiraan] = 0) AND ([Table Journal].Periode = " & PeriodeData & ")", CNN, lckLockReadOnly
With Rc
     If .Recordcount <> 0 Then
        If mVarJrl.CiptaKaryaHeaderJournal("LINK", "", "", "", "", "", "IDR", Now(), Trim(Str(PeriodeData)), "LINK") = True Then
           mVarJrl.CiptaKaryaDetailJournal "", AccountLink, "xxx", 0, IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        End If
     Else
     End If
End With
End Sub

Private Sub CompareAccount()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
RcCom.DBOpen "SELECT GLAccount.NoAccount, [Tabel Pembantu].NoAccount AS NoAccountB FROM         GLAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            If IsNull(mVarData(1, I)) Then SendDataToServer ("INSERT INTO [Tabel Pembantu] (NoAccount) VALUES (N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
End Sub


Private Sub IsiListDataDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
Dim mSaldo As Variant
Dim mTotalDR As Variant
Dim mTotalCR As Variant
Dim mCdr As Variant
Dim mCcr As Variant
Dim mPer As Integer
Dim strSQL As String

Select Case mVarPeriode
       Case 1: mPer = 0
       Case 2: mPer = 1
       Case 3: mPer = 2
       Case 4: mPer = 3
       Case 5: mPer = 4
       Case 6: mPer = 5
       Case 7: mPer = 6
       Case 8: mPer = 7
       Case 9: mPer = 8
       Case 10: mPer = 9
       Case 11: mPer = 10
       Case 12: mPer = 11
End Select
'
RcCom.DBOpen "SELECT GLAccount.NoAccount, [Tabel Pembantu].CurrentDR" & mPer & " - [Tabel Pembantu].CurrentCR" & mPer & " AS [Saldo Awal], SUM([Detail Journal].Debet) AS Debet,SUM([Detail Journal].Credit) AS Credit, [Table Journal].Periode, GLAccount.[Group] FROM         GLAccount INNER JOIN [Detail Journal] ON GLAccount.NoAccount = [Detail Journal].NoAccount INNER JOIN [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount GROUP BY GLAccount.NoAccount, [Table Journal].Periode, GLAccount.[Group],[Tabel Pembantu].CurrentDR" & mPer & " , [Tabel Pembantu].CurrentCR" & mPer & " HAVING      ([Table Journal].Periode = " & mVarPeriode & ")", CNN, lckLockReadOnly
'Debug.Print RcCom.PrepareQuery
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            'Awal Variabel
            mCdr = mVarData(2, I)
            mCcr = mVarData(3, I)
            mSaldo = mVarData(1, I)
            If mSaldo < 0 Then mSaldo = mSaldo * (-1)
            If mCdr > mCcr Then 'Saldo      Debet             Credit
               mTotalDR = ((mSaldo + mVarData(2, I)) - mVarData(3, I))
               'If mTotalDR < 0 Then mTotalDR = mTotalDR * (-1)
               mTotalCR = 0
            Else                'Saldo      Credit            Debet
               mTotalCR = ((mSaldo + mVarData(3, I)) - mVarData(2, I))
               'If mTotalCR < 0 Then mTotalCR = mTotalCR * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(IIf(Not IsNull(mTotalCR), mTotalCR, 0)) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            'CLOSING YEAR UPDATE AMOUNT PERIODE 0 FOR BEGINNING NEXT YEAR JANUARY
            If mVarPeriode = 12 Then
                SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR0 = " & CCur(mTotalDR) & ", CurrentCR0 = " & CCur(IIf(Not IsNull(mTotalCR), mTotalCR, 0)) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            End If
            
        Next I
     End If
End With
'RcCom.CloseDB
IsiListData
IsiSubDetail
IsiDetail
IsiGroupDetail
SendDataToServer ("UPDATE SettingPeriod SET Closed = 1 WHERE (Periode=" & mVarPeriode & ") AND Left([GlFile],4)='" & TahunFiskalYear & "'")
strSQL = "UPDATE settingperiod SET Closed = 1, date_closed = CONVERT(DATETIME, '" & Format$(Now, LongDateForm) & "', 3), user_closed = N'" & MainMenu.StatusBar1.Panels(1).Text & "' WHERE (Periode=" & mVarPeriode & ") AND ((Left([GlFile],4)='" & TahunFiskalYear & "')"
End Sub

Private Function IdxAuto() As String
Dim mNo As Double
Dim mStr As String
If mVarJournal = False Then
   IdxAuto = TglIndex
   mVarJrl = IdxAuto
Else
   mNo = Val(Right(mVarJrl, 5))
   mNo = mNo + 1
   mStr = Left(mVarJrl, 10) & KirimNull(5 - Len(Trim(Str(mNo)))) & Trim(Str(mNo))
   IdxAuto = mStr
End If
End Function

Private Sub IsiListData()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
Dim mSaldo As Variant
Dim mTotalDR As Variant
Dim mTotalCR As Variant
Dim mCdr As Variant
Dim mCcr As Variant
RcCom.DBOpen "SELECT  GLAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GLAccount INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GLAccount.[Group] = N'Detail List Account') GROUP BY GLAccount.GroupAccount", CNN, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            If mVarPeriode = 12 Then
                SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR0 = " & CCur(mTotalDR) & ", CurrentCR0 = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            End If
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiSubDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
Dim mSaldo As Variant
Dim mTotalDR As Variant
Dim mTotalCR As Variant
Dim mCdr As Variant
Dim mCcr As Variant

RcCom.DBOpen "SELECT  GLAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GLAccount INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GLAccount.[Group] = N'List Account') GROUP BY GLAccount.GroupAccount", CNN, lckLockReadOnly
'MessageBox RcCom.PrepareQuery
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            If mVarPeriode = 12 Then
                SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR0 = " & CCur(mTotalDR) & ", CurrentCR0 = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            End If
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
Dim mSaldo As Variant
Dim mTotalDR As Variant
Dim mTotalCR As Variant
Dim mCdr As Variant
Dim mCcr As Variant
RcCom.DBOpen "SELECT     GLAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GLAccount INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GLAccount.[Group] = N'Sub Account') GROUP BY GLAccount.GroupAccount", CNN, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            If mVarPeriode = 12 Then
                SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR0 = " & CCur(mTotalDR) & ", CurrentCR0 = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            End If
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiGroupDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
Dim mSaldo As Variant
Dim mTotalDR As Variant
Dim mTotalCR As Variant
Dim mCdr As Variant
Dim mCcr As Variant
RcCom.DBOpen "SELECT     GLAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GLAccount INNER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GLAccount.[Group] = N'Group Account') GROUP BY GLAccount.GroupAccount", CNN, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            If mVarPeriode = 12 Then
                SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR0 = " & CCur(mTotalDR) & ", CurrentCR0 = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
            End If
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Function TglIndex() As String
Dim MyData As New clsTransaksi
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = MyData.PrepareIndex(tmbTransaksiAkumDepre, 5, "", "AD-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-")
End Function
