VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListSPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Surat Permintaan Pembelian"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListSPP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11940
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11940
      TabIndex        =   8
      Top             =   7845
      Width           =   11940
      Begin VB.Frame FrTombol 
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
         Height          =   30
         Left            =   -30
         TabIndex        =   14
         Top             =   0
         Width           =   11445
      End
      Begin VB.CommandButton cmdPO 
         Caption         =   "&Set  PO"
         Height          =   555
         Index           =   1
         Left            =   1065
         Picture         =   "frmListSPP.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Set Ke Surat Permintaan Penawaran Harga"
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   555
         Left            =   1785
         Picture         =   "frmListSPP.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   555
         Left            =   2505
         Picture         =   "frmListSPP.frx":138F6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdPO 
         Caption         =   "&Set  SPPH"
         Height          =   555
         Index           =   0
         Left            =   90
         Picture         =   "frmListSPP.frx":153F0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Set Ke Surat Permintaan Penawaran Harga"
         Top             =   60
         Width           =   975
      End
   End
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
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7845
      ScaleWidth      =   11940
      TabIndex        =   9
      Top             =   0
      Width           =   11940
      Begin VB.ListBox ListSupplier 
         Height          =   7080
         Left            =   30
         TabIndex        =   0
         Top             =   300
         Width           =   2655
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Tampilkan Semua Permintaan"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Top             =   7530
         Width           =   2415
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Pilih Semua"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   7530
         Width           =   1215
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8415
         Picture         =   "frmListSPP.frx":1BC42
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   4
         Top             =   7500
         Width           =   255
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
               Picture         =   "frmListSPP.frx":22494
               Key             =   "SPP"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListSPP.frx":28CF6
               Key             =   "SPPH"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListBarang 
         Height          =   7125
         Left            =   2715
         TabIndex        =   1
         Top             =   300
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   12568
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tanggal"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Satuan"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama Barang"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "No SPP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Keperluan"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "No.PO"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tgl Pembelian"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Keterangan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Price"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Issued By"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Top             =   90
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Daftar Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   90
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barang dlm Status Permintaan Penawaran"
         Height          =   255
         Index           =   2
         Left            =   8760
         TabIndex        =   10
         Top             =   7530
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmListSPP"
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
Dim lIdx As Integer

Private Sub chk_Click(Index As Integer)
   Dim x As Integer
   Select Case Index
      Case 0:
         If chk(Index).Value = 1 Then
            OpenDetail
         Else
            OpenDetail rsSupplier.DBRecordset.Fields("PartnerID")
         End If
      Case 1:
         If chk(Index).Value = 1 Then
            For x = 1 To ListBarang.ListItems.Count
               If ListBarang.ListItems(x).SubItems(9) <> "1" Then ListBarang.ListItems(x).Checked = True
            Next
         End If
   End Select
End Sub

Private Sub cmdCetak_Click()
   Dim aReport As New utility
   aReport.CallReportView "Select * From QueryListSPP where Status < 2", "ListSPP.rpt", ReportPath, "Daftar Permintaan Pembelian"
   'CallRPTReport "ListSPP.rpt", "Select * From QueryListSPP where Status =0"
   Set aReport = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub



Private Sub OpenPartner()
   
On Error GoTo Hell:
   RcPartner.DBOpen "SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'SUPPLIER') ORDER BY PartnerID", CNN, lckLockReadOnly

   If RcPartner.Recordcount <> 0 Then
      mCall.FromTagActive = "Supplier List"
      mCall.CaptionLink = "Pilih Supplier"
      Set mCall.FormData = RcPartner.DBRecordset
      mCall.LookUp Me
      mCall.Show vbModal
   Else
      MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
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


Private Sub cmdPO_Click(Index As Integer)
   If IsCheckedGrid Then
      If MessageBox("Yakin data akan di set ke PO ? ", "Konfirmasi", msgYesNo, msgQuestion) = 1 Then
         If chk(0).Value = 1 Then
             '*** Pilih Supplier
             lIdx = Index
             OpenPartner
         Else
            IDSpl = rsSupplier.DBRecordset.Fields("PartnerID")
            SetPurchase Index
         End If
      End If
   Else
      MessageBox "Tidak ada data yang dicentang !", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub

Private Sub Form_Load()
   LoadSupplier
   If rsSupplier.DBRecordset.Recordcount > 0 Then
      OpenDetail rsSupplier.DBRecordset.Fields("PartnerID").Value
   End If
   
   'HiasForm Picture1, Me
   'HiasFormManTell Picture2, Me
   
   Set mCall = New frmCaller
End Sub

Private Sub LoadSupplier()
   rsSupplier.DBOpen "Select * from QuerySPPSupplier ", CNN
   ListSupplier.Clear
   If rsSupplier.DBRecordset.Recordcount > 0 Then
   While Not rsSupplier.DBRecordset.EOF
      ListSupplier.AddItem IIf(IsNull(rsSupplier.DBRecordset.Fields("CompanyName")), "Non Supplier", rsSupplier.DBRecordset.Fields("CompanyName"))
      rsSupplier.DBRecordset.MoveNext
   Wend
   End If
   If rsSupplier.DBRecordset.Recordcount > 0 Then rsSupplier.DBRecordset.MoveFirst
End Sub

Private Sub OpenDetail(Optional aSupplier As String = "")
   Dim chList As ListItem
   Dim x As Integer
   'status Table SPP_line -> 0:SPP
                            '1:SPPH
                            '2:PO
   If aSupplier = "" Then
      rsList.DBOpen "select * from QueryListSPP where (status < 3) and (approved_by is not null) order by SPP_Date desc, SPPID", CNN, lckLockBatch, lckLockSync
   Else
      rsList.DBOpen "select * from QueryListSPP where (PartnerID='" & aSupplier & "') and (status < 3) and (approved_by is not null) order by SPP_Date desc, SPPID", CNN, lckLockBatch, lckLockSync
   End If
   
   ListBarang.ListItems.Clear
   If rsList.DBRecordset.Recordcount > 0 Then
   rsList.DBRecordset.MoveFirst
   For x = 1 To rsList.DBRecordset.Recordcount
      If rsList.DBRecordset.Fields("status") = 1 Then
         Set chList = ListBarang.ListItems.Add(x, "A" & x, rsList.DBRecordset.Fields("NoItem"), , "SPPH")
      Else
         Set chList = ListBarang.ListItems.Add(x, "A" & x, rsList.DBRecordset.Fields("NoItem"))
      End If
         ListBarang.ListItems(x).ListSubItems.Add 1, "tgl", Format(rsList.DBRecordset.Fields("SPP_date"), "dd MMM yyyy")
         ListBarang.ListItems(x).ListSubItems.Add 2, "Qty", rsList.DBRecordset.Fields("Qty_SPP")
         ListBarang.ListItems(x).ListSubItems.Add 3, "uom", rsList.DBRecordset.Fields("uom")
         ListBarang.ListItems(x).ListSubItems.Add 4, "brg", rsList.DBRecordset.Fields("ItemName")
         ListBarang.ListItems(x).ListSubItems.Add 5, "SPP", rsList.DBRecordset.Fields("SPPID")
         ListBarang.ListItems(x).ListSubItems.Add 6, "kp", rsList.DBRecordset.Fields("Keperluan")
         ListBarang.ListItems(x).ListSubItems.Add 7, "po", IIf(IsNull(rsList.DBRecordset.Fields("po")), "", rsList.DBRecordset.Fields("po"))
         ListBarang.ListItems(x).ListSubItems.Add 8, "TglB", IIf(IsNull(rsList.DBRecordset.Fields("receive_date")), "", Format(rsList.DBRecordset.Fields("receive_date"), "dd MMM yyyy"))
         ListBarang.ListItems(x).ListSubItems.Add 9, "sts", rsList.DBRecordset.Fields("Status")
         ListBarang.ListItems(x).ListSubItems.Add 10, "Ket", rsList.DBRecordset.Fields("Note")
         ListBarang.ListItems(x).ListSubItems.Add 11, "pr", rsList.DBRecordset.Fields("priceIn")
         ListBarang.ListItems(x).ListSubItems.Add 12, "Iss", IIf(IsNull(rsList.DBRecordset.Fields("empID")), "", rsList.DBRecordset.Fields("empID"))
      
         rsList.DBRecordset.MoveNext
   Next
   rsList.DBRecordset.MoveFirst
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   IDSpl = ""
End Sub

Private Sub ListBarang_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   If Item.SubItems(9) = "1" Then 'jika item ini dalm status SPPH
      Item.Checked = False
   End If
   
   If Item.Checked = False Then 'uncheck opsi "pilih semua"
      chk(1).Value = 0
   End If
End Sub

Private Sub ListSupplier_Click()
   rsSupplier.DBRecordset.MoveFirst
   rsSupplier.DBRecordset.Move ListSupplier.ListIndex
   OpenDetail rsSupplier.DBRecordset.Fields("PartnerID")
   IDSpl = rsSupplier.DBRecordset.Fields("PartnerID")
End Sub

Private Sub mCall_BeforeUnload()
   SetPurchase lIdx
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   IDSpl = pRecordset.Fields("Partner ID")
End Sub


Private Sub SetPurchase(Index As Integer)
   Dim strSQL As String
   Dim aSPPHID As String
   Dim IDGen As New IDGenerator
   Dim x As Integer
   Dim LastID As Integer
         
'         ListBarang.ListItems(x).ListSubItems.Add 1, "tgl", Format(rsList.DBRecordset.Fields("SPP_date"), "dd MMM yyyy")
'         ListBarang.ListItems(x).ListSubItems.Add 2, "Qty", rsList.DBRecordset.Fields("Qty_SPP")
'         ListBarang.ListItems(x).ListSubItems.Add 3, "uom", rsList.DBRecordset.Fields("uom")
'         ListBarang.ListItems(x).ListSubItems.Add 4, "brg", rsList.DBRecordset.Fields("ItemName")
'         ListBarang.ListItems(x).ListSubItems.Add 5, "SPP", rsList.DBRecordset.Fields("SPPID")
'         ListBarang.ListItems(x).ListSubItems.Add 6, "kp", rsList.DBRecordset.Fields("Keperluan")
'         ListBarang.ListItems(x).ListSubItems.Add 7, "po", ""
'         ListBarang.ListItems(x).ListSubItems.Add 8, "TglB", "" 'rsList.DBRecordset.Fields("tgl_in")
'         ListBarang.ListItems(x).ListSubItems.Add 9, "sts", rsList.DBRecordset.Fields("Status")
'         ListBarang.ListItems(x).ListSubItems.Add 10, "Ket", rsList.DBRecordset.Fields("Note")
'         ListBarang.ListItems(x).ListSubItems.Add 11, "pr", rsList.DBRecordset.Fields("priceIn")
   
   
   '*** IDSpl adl ID Supplier yg dihasilkan dr proses procedure OpenPartner
   If IDSpl <> "" Then
   
      Select Case Index
         Case 0:
            '*** Buat Header SPPH
            aSPPHID = IDGen.GetID("OF")
            strSQL = "Insert into SPPH_Header (SPPHID,DateTRans,PartnerID,UserReqst) values ('" & aSPPHID & "','" & Format(Now, "yyyy-MM-dd") & "','" & IDSpl & "','" & MainMenu.StatusBar1.Panels(1).Text & "')"
            If SendDataToServer(strSQL) Then
               For x = 1 To ListBarang.ListItems.Count
                  If ListBarang.ListItems(x).Checked Then
                     
                     '*** Ubah Status SPP_Line = 1
                      strSQL = "Update SPP_Line set status=1 where SPPID='" & ListBarang.ListItems(x).SubItems(5) & "' and noItem ='" & ListBarang.ListItems(x).Text & "'"
                      SendDataToServer (strSQL)
                      
                     '*** Buat SPPH
                      strSQL = "insert into SPPH_line (SPPHID,NoItem,Qty_Spph,refNote) values ('" & aSPPHID & _
                                                                                           "','" & ListBarang.ListItems(x).Text & _
                                                                                           "', " & FQty(ListBarang.ListItems(x).SubItems(2)) & _
                                                                                           ",' " & ListBarang.ListItems(x).SubItems(10) & "')"
                      SendDataToServer strSQL
                      
                      '*** Update Data Inventory
                      strSQL = "Update Inventory set partnerID = '" & IDSpl & "' where noItem ='" & ListBarang.ListItems(x).Text & "'"
                      SendDataToServer strSQL
                   End If
               Next
            End If
         Case 1:
            '*** Buat Header PO
            aSPPHID = IDGen.GetID("PO")
            strSQL = " INSERT INTO  [PO Order] ( [REquire Date], PurchaseID, EmpID, PartnerID, DatePurchase, TermPayment, Periode, TypeTrans,Account ,Discount ,termMethod, keterangan, type_trans_order) " & _
                        " VALUES ('" & Format(Now, "yyyy-MM-dd") & "',N'" & aSPPHID & "',N'" & MainMenu.StatusBar1.Panels(1).Text & "', N'" & IDSpl & "'," & _
                        " '" & Format(Now, "yyyy-MM-dd") & "', 0, " & Val(Month(Now)) & ", N'PO',N'' , 0 ,'CASH','-',2)"
         
            If SendDataToServer(strSQL) Then
               For x = 1 To ListBarang.ListItems.Count
                  If ListBarang.ListItems(x).Checked Then
                     
                     '*** Ubah Status SPP_Line = 2
                      strSQL = "Update SPP_Line set status=2 where SPPID='" & ListBarang.ListItems(x).SubItems(5) & "' and noItem ='" & ListBarang.ListItems(x).Text & "'"
                      SendDataToServer (strSQL)
                      
                     '*** tambah keterangan SPP
                     strSQL = "update [PO Order] set keterangan = keterangan + '" & ListBarang.ListItems(x).SubItems(5) & " , " & "' where PurchaseID='" & aSPPHID & "'"
                     SendDataToServer strSQL
                     
                     
                     '*** Buat Detail PO
                     strSQL = " INSERT INTO [Detail PO] ( PurchaseID, NoItem,                             QTYPO,                                       ItemSupplierID    , POPrice,                                            ScheduleDate,                                          VAT,   QtyTemp,                                     TCredit,Hpp,                                               sppid,tipe_item,curID,rate)" & _
                               " VALUES (N'" & aSPPHID & "', N'" & ListBarang.ListItems(x).Text & "', " & FQty(ListBarang.ListItems(x).SubItems(2)) & ", N'" & IDSpl & "', " & FQty(ListBarang.ListItems(x).SubItems(11)) & ", convert(Datetime,'" & Format(Now, "dd/mm/yy") & "',3), 0, " & FQty(ListBarang.ListItems(x).SubItems(2)) & ",0," & FQty(ListBarang.ListItems(x).SubItems(11)) & ",'" & ListBarang.ListItems(x).SubItems(5) & "','I','IDR',1)"

                      SendDataToServer strSQL
                      
                      '*** Update Data Inventory
                      strSQL = "Update Inventory set partnerID = '" & IDSpl & "' where noItem ='" & ListBarang.ListItems(x).Text & "'"
                      SendDataToServer strSQL
                      
                      '*** Update SPP_line (Isi No PO)
                      strSQL = "Update Spp_line set PO='" & aSPPHID & "' where SPPID='" & ListBarang.ListItems(x).SubItems(5) & "' and noItem='" & ListBarang.ListItems(x).Text & "'"
                      SendDataToServer strSQL
                   End If
               Next
            End If
      
      End Select
      
      '*** Refresh Form
         LastID = ListSupplier.ListIndex
         LoadSupplier
         If rsSupplier.DBRecordset.Recordcount > 0 Then
            rsSupplier.DBRecordset.MoveFirst
            rsSupplier.DBRecordset.Move LastID
            On Error Resume Next
            OpenDetail rsSupplier.DBRecordset.Fields("PartnerID").Value
         End If
                  
         
   Else
         MessageBox "Proses Retrieve Gagal", "Informasi", msgOkOnly, msgInfo
   End If
End Sub

