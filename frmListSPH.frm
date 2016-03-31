VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListSPH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Surat Penawaran Harga"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListSPH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11025
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
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
      Height          =   5730
      Left            =   0
      ScaleHeight     =   5730
      ScaleWidth      =   11025
      TabIndex        =   7
      Top             =   0
      Width           =   11025
      Begin VB.ListBox ListSPH 
         Height          =   5130
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox cb 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   8
         Top             =   1440
         Width           =   200
      End
      Begin VB.CheckBox chk 
         Caption         =   "Tampilkan Semua Permintaan"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk 
         Caption         =   "Pilih Semua"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   5280
         Width           =   1215
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1560
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListSPH.frx":6852
               Key             =   "SPP"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListSPH.frx":D0B4
               Key             =   "SPPH"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListBarang 
         Height          =   4815
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama Barang"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Harga"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Total"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Keterangan"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NO SPH"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Daftar Barang"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10965
      TabIndex        =   6
      Top             =   5730
      Width           =   11025
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   515
         Left            =   10050
         Picture         =   "frmListSPH.frx":13916
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdPO 
         Caption         =   "&Set  PO"
         Height          =   515
         Left            =   50
         Picture         =   "frmListSPH.frx":15410
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Set Ke Order Pembelian"
         Top             =   25
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmListSPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsSupplier As New DBQuick
Dim rsList As New DBQuick
Dim RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim IDSpl As String




Private Sub chk_Click(Index As Integer)
   Dim x As Integer
   Select Case Index
      Case 0:
         If chk(Index).Value = 1 Then
            OpenDetail
         Else
            OpenDetail rsSupplier.DBRecordset.Fields(0)
         End If
      Case 1:
         If chk(Index).Value = 1 Then
            For x = 1 To ListBarang.ListItems.Count
               ListBarang.ListItems(x).Checked = True
            Next
         End If
   End Select
End Sub

Private Sub cmdCetak_Click()
   CallRPTReport "ListSPP.rpt", "Select * From QueryListSPP where Status =0"
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPO_Click()
   If IsCheckedGrid Then
      If MessageBox("Yakin Data Akan Di set Ke PO ? ", "Konfirmasi", msgYesNo) = 1 Then
         'pilih Supplier
          OpenPartner
      End If
   Else
      MessageBox "Tidak ada data yang dicentang !", "Peringatan"
   End If
End Sub


Private Sub OpenPartner()
   
On Error GoTo Hell:
   RcPartner.DBOpen "SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'SUPPLIER') ORDER BY PartnerID", CNN, lckLockReadOnly

   If RcPartner.Recordcount <> 0 Then
      mCall.FromTagActive = "Supplier List"
      mCall.CaptionLink = "Pilih Supplier"
      Set mCall.FormData = RcPartner.DBRecordset
      mCall.LookUp Me
   Else
      MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   End If

Exit Sub
Hell:
    Err.Clear
End Sub


Private Function IsCheckedGrid() As Boolean
   Dim chk As Boolean
   Dim x As Integer
   chk = False
   For x = 1 To ListBarang.ListItems.Count
      If ListBarang.ListItems(x).Checked Then
         chk = True
         Exit For
      End If
   Next
   IsCheckedGrid = chk
End Function



Private Sub Form_Load()
   LoadSphHeader
   OpenDetail rsSupplier.DBRecordset.Fields(0).Value
   
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   
   Set mCall = New frmCaller
End Sub

Private Sub LoadSphHeader()
   rsSupplier.DBOpen "Select * from QuerySphSupplier ", CNN
   ListSPH.Clear
   While Not rsSupplier.DBRecordset.EOF
      ListSPH.AddItem IIf(IsNull(rsSupplier.DBRecordset.Fields("CompanyName")), "Non Supplier", rsSupplier.DBRecordset.Fields("CompanyName")) & _
                       " (" & rsSupplier.DBRecordset.Fields(0) & " )"
      rsSupplier.DBRecordset.MoveNext
   Wend
   
   rsSupplier.DBRecordset.MoveFirst

End Sub

Private Sub OpenDetail(Optional aID As String = "")
   Dim chList As ListItem
   Dim x As Integer
   'status Table SPP_line -> 0:SPP
                            '1:SPPH
                            '2:PO
   If aID = "" Then
      rsList.DBOpen "select * from QueryDetailPurchaseOffer", CNN, lckLockBatch, lckLockSync
   Else
      rsList.DBOpen "select * from QueryDetailPurchaseOffer where SPPHID='" & aID & "'", CNN
   End If
   
   ListBarang.ListItems.Clear
   rsList.DBRecordset.MoveFirst
   For x = 1 To rsList.DBRecordset.Recordcount
         Set chList = ListBarang.ListItems.Add(x, "A" & x, rsList.DBRecordset.Fields("NoItem"))
         ListBarang.ListItems(x).ListSubItems.Add 1, "brg", rsList.DBRecordset.Fields("ItemName")
         ListBarang.ListItems(x).ListSubItems.Add 2, "uom", rsList.DBRecordset.Fields("Uom")
         ListBarang.ListItems(x).ListSubItems.Add 3, "Qty", rsList.DBRecordset.Fields("Qty_SPPH")
         ListBarang.ListItems(x).ListSubItems.Add 4, "hrg", rsList.DBRecordset.Fields("Price")
         ListBarang.ListItems(x).ListSubItems.Add 5, "tot", rsList.DBRecordset.Fields("fTotal")
         ListBarang.ListItems(x).ListSubItems.Add 6, "Dis", rsList.DBRecordset.Fields("Discount")
         ListBarang.ListItems(x).ListSubItems.Add 7, "ket", rsList.DBRecordset.Fields("RefNote")
         rsList.DBRecordset.MoveNext
   Next
   rsList.DBRecordset.MoveFirst
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
IDSpl = ""
End Sub

Private Sub ListBarang_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   If Item.Checked = False Then 'uncheck opsi "pilih semua"
      chk(1).Value = 0
   End If
End Sub

Private Sub ListSPH_Click()
   rsSupplier.DBRecordset.MoveFirst
   rsSupplier.DBRecordset.Move ListSPH.ListIndex
   OpenDetail rsSupplier.DBRecordset.Fields(0)
End Sub

Private Sub mCall_BeforeUnload()
   Dim strSQL As String
   Dim aSPPHID As String
   Dim IDGen As New IDGenerator
   Dim x As Integer
   Dim LastID As Integer
         
   'IDSpl adl ID Supplier yg dihasilkan dr proses procedure OpenPartner
   If IDSpl <> "" Then
      'Buat Header SPPH
      aSPPHID = IDGen.GetID("OF")
      strSQL = "Insert into SPPH_Header (SPPHID,DateTRans,PartnerID,UserReqst) values ('" & aSPPHID & "','" & Format(Now, "yyyy-MM-dd") & "','" & IDSpl & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
      If SendDataToServer(strSQL) Then
         For x = 1 To ListBarang.ListItems.Count
            If ListBarang.ListItems(x).Checked Then
               'Ubah Status SPP_Line = 1
                strSQL = "update SPP_Line set status=1 where SPPID='" & ListBarang.ListItems(x).SubItems(1) & "' and noItem ='" & ListBarang.ListItems(x).Text & "'"
                SendDataToServer (strSQL)
               'Buat SPPH
                strSQL = "insert into SPPH_line (SPPHID,NoItem,Qty_Spph,refNote) values ('" & aSPPHID & _
                                                                                     "','" & ListBarang.ListItems(x).Text & _
                                                                                     "', " & FQty(ListBarang.ListItems(x).SubItems(3)) & _
                                                                                     ",' " & ListBarang.ListItems(x).SubItems(7) & "')"
                SendDataToServer strSQL
             End If
         Next
      'Refresh Form
         LastID = ListSPH.ListIndex
         LoadSphHeader
         rsSupplier.DBRecordset.MoveFirst
         rsSupplier.DBRecordset.Move LastID
         OpenDetail rsSupplier.DBRecordset.Fields(0).Value
         
      Else
         MessageBox "Proses Retrieve Gagal", "Informasi"
      End If
   Else
      MessageBox "Proses Gagal Atau Dibatalkan ", "Informasi"
   End If

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   IDSpl = pRecordset.Fields("Partner ID")
End Sub
