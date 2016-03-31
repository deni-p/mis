VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmSFComplete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penerimaan Produk Jadi"
   ClientHeight    =   6105
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "frmTerimaProdukJadi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   10185
      TabIndex        =   5
      Top             =   0
      Width           =   10185
      Begin MSComctlLib.ListView ListFG 
         Height          =   4425
         Left            =   105
         TabIndex        =   4
         Top             =   1005
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   7805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tanggal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Kode Produk"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Produk"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lot"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Keterangan"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "WareHouse"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   1
         Left            =   6735
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "SPP"
         Top             =   150
         Width           =   2940
      End
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   9675
         Picture         =   "frmTerimaProdukJadi.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   158
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "No_penerimaan"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1590
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "SPP"
         Top             =   150
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "tgl_terima"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1590
         TabIndex        =   2
         Tag             =   "SPP"
         Top             =   525
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   65863683
         CurrentDate     =   38272
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5250
         X2              =   6750
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Tujuan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   5280
         TabIndex        =   9
         Top             =   218
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Penerimaan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   600
         Width           =   570
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   150
         X2              =   1650
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   150
         X2              =   1650
         Y1              =   840
         Y2              =   840
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   1005
      BindFormTAG     =   "SPP"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmSFComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDetail                                          As New DBQuick
Attribute RcDetail.VB_VarHelpID = -1
Private RcPartner                                         As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private MEdit                                             As Boolean

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture2, Me
   Set mCall = New frmCaller
   DTPicker1.Value = dDateBegin
   With MyDDE
        .EditModeReplace = False
        Set .BindForm = Me
        Set .ActiveConnection = CNN
        .PrepareQuery = "SELECT * from TerimaProdukJadi_Header"
        .SetPermissions = aksess.MayDo("Penerimaan Barang Jadi")
   End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSFComplete = Nothing
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
   If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
      MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   Else
      If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Then
         If MyDDE.ChildRecordset.Fields(0) = "" Then
            MyDDE.ChildRecordset.CancelBatch adAffectCurrent
            If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
         End If
      End If
   End If
   
End Sub


Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
   Select Case TagForm:
      Case "Produk Dikirim":
         MyDDE.ChildRecordset.Fields("TransID") = MyDDE.GetFieldByName("No Penyerahan")
         MyDDE.ChildRecordset.Fields("noItem") = mCall.GetFieldByName("Kode Barang")
         MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("Nama Barang")
         MyDDE.ChildRecordset.Fields("UOM") = mCall.GetFieldByName("Satuan")
         MyDDE.ChildRecordset.Fields("Quantity input") = mCall.GetFieldByName("Jml Dikirim")
         MyDDE.ChildRecordset.Fields("Qty_received") = mCall.GetFieldByName("jml Dikirim")
      Case "Gudang Tujuan":
         txtBox(1).Text = mCall.GetFieldByName(0)
         
   End Select
End If
End Sub



Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit: MEdit = True
       Case tmbAddNew: MEdit = False
       Case tmbSave:
         If Trim(txtBox(1).Text) <> "" Then
            PrepareQuery
            MyDDE.IsChildMemberReady = True
         End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim IDGen As New IDGenerator
Dim rsBalance As New DBQuick
Dim x As Integer

txtBox(0).Enabled = False
cmdLink(1).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            cmdLink(1).Enabled = True
             
            
       Case tmbAddNew:
            DTPicker1.Value = Now
            MyDDE.GetFieldByName("no_penerimaan") = IDGen.GetID("TP")
            LoadSerahTerima
            DTPicker1.SetFocus
            cmdLink(1).Enabled = True
            
            
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               OpenDetail txtBox(0)
            Else
               MessageBox "Gudang Tujuan belum didefinisikan", "Peringatan", msgOkOnly
               cmdLink(1).Enabled = True
            End If
            
       Case tmbCancel:
             
       Case tmbDetail:
               OpenPartner 0
               
       Case tmbPrint:
               Dim aReport As New utility
               aReport.CallReportView "select * from ReportTerimaProdukJadi where OrderID='" & MyDDE.GetFieldByName("OrderID") & "'", "ReportTerimaProdukJadi.rpt", ReportPath, "Penerimaan Produk Jadi"
               Set aReport = Nothing
      
      Case tmbDelete:
         If ListFG.ListItems.Count <> 0 Then
            For x = 1 To ListFG.ListItems.Count
               With ListFG.ListItems(x)
                  rsBalance.DBOpen "select Qty_out from [inventory tabel] where ln_no='" & .SubItems(4) & "'", CNN
                  If rsBalance.DBRecordset.Recordcount > 0 Then
                     If rsBalance.DBRecordset.Fields(0) = 0 Then
                        '*** Update Stok Barang ***
                        SendDataToServer "delete from [inventory tabel] where ln_no='" & ListFG.ListItems(x).SubItems(4) & "'"
                  
                        '*** Update item FinishGood ***
                        SendDataToServer "update item_finishGood set status=0 where ln_no='" & ListFG.ListItems(x).SubItems(4) & "'"
                  
                        '*** Update FinishGood ***
                        SendDataToServer "update finishgood set received_by='' where ID='" & txtBox(0).Text & "'"
                     End If
                  End If
               End With
            Next
         End If
End Select

'Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Not IsNull(MyDDE.GetFieldByName("ID")) Then OpenDetail MyDDE.GetFieldByName("ID")
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT backflush_line.IDTrans as [No Penyerahan], backflush_line.NoItem as [Kode Barang], Inventory.ItemName as [Nama Barang], Inventory.UOM as Satuan, backflush_line.[Quantity Input] as [Jml Dikirim], " & _
                                      "backflush_line.[Qty Received] " & _
                             "FROM   backflush_line INNER JOIN " & _
                                     " Inventory ON backflush_line.NoItem = Inventory.NoItem INNER JOIN " & _
                                     " [backflush_line Header] ON backflush_line.IDTrans = [backflush_line Header].IDTrans " & _
                             "WHERE  (backflush_line.Status = 0) AND ([backflush_Header].TypeTrans = 'PJ')", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen "select warehouse as Kode,[warehouse name] as Gudang from warehouse", CNN, lckLockReadOnly
            
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Produk Dikirim"
          Case 1:
            mCall.FromTagActive = "Gudang Tujuan"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
   Dim x As Integer
   If ParameterString = "" Then ParameterString = "xxxxxxxx"
   RcDetail.DBOpen "SELECT * FROM TerimaProdukJadi_detail where ID='" & ParameterString & "'", CNN, lckLockBatch
   'Set MyDDE.ChildRecordset = RcDetail.DBRecordset '.Clone(adLockBatchOptimistic)
   ListFG.ListItems.Clear
   x = 1
   With RcDetail.DBRecordset
      If .Recordcount > 0 Then
         .MoveFirst
         While Not .EOF
            ListFG.ListItems.Add x, "A" & x, .Fields("ID")
            ListFG.ListItems(x).SubItems(1) = Format(.Fields("Tanggal"), "dd MMM yyyy")
            ListFG.ListItems(x).SubItems(2) = .Fields("noItem")
            ListFG.ListItems(x).SubItems(3) = .Fields("ItemName")
            ListFG.ListItems(x).SubItems(4) = .Fields("ln_no")
            ListFG.ListItems(x).SubItems(5) = .Fields("kuantitas")
            ListFG.ListItems(x).SubItems(6) = .Fields("satuan")
            ListFG.ListItems(x).SubItems(7) = .Fields("keterangan")
            ListFG.ListItems(x).Checked = True
            .MoveNext
            x = x + 1
         Wend
      End If
   End With
End Sub

Private Sub LoadSerahTerima()
   Dim x As Integer
   RcDetail.DBOpen "SELECT * FROM TerimaProdukJadi_detail where status=0", CNN, lckLockBatch
   'Set MyDDE.ChildRecordset = RcDetail.DBRecordset '.Clone(adLockBatchOptimistic)
   ListFG.ListItems.Clear
   x = 1
   With RcDetail.DBRecordset
      If .Recordcount > 0 Then
         .MoveFirst
         While Not .EOF
            ListFG.ListItems.Add x, "A" & x, .Fields("ID")
            ListFG.ListItems(x).SubItems(1) = Format(.Fields("Tanggal"), "dd MMM yyyy")
            ListFG.ListItems(x).SubItems(2) = .Fields("NoItem")
            ListFG.ListItems(x).SubItems(3) = .Fields("ItemName")
            ListFG.ListItems(x).SubItems(4) = .Fields("ln_no")
            ListFG.ListItems(x).SubItems(5) = .Fields("kuantitas")
            ListFG.ListItems(x).SubItems(6) = .Fields("satuan")
            ListFG.ListItems(x).SubItems(7) = .Fields("keterangan")
            ListFG.ListItems(x).Checked = False
            .MoveNext
            x = x + 1
         Wend
      End If
   End With
End Sub

Private Sub SimpanDetail()
Dim rsBalance As New DBQuick
Dim SisaStock As Double
Dim IsLockFIFO As String
Dim x As Integer
If ListFG.ListItems.Count > 0 Then
   For x = 1 To ListFG.ListItems.Count
      If ListFG.ListItems(x).Checked Then
         With ListFG.ListItems(x)
         
            If MEdit Then
                  
            Else
               '***  update finishGood ***
               SendDataToServer "update backflush_header set status=1 where IDTrans='" & ListFG.ListItems(x).Text & "'"
            
               '** Update item_finishgood **
               SendDataToServer "Update backflush_output set receive_id='" & txtBox(0).Text & "', transfer_date= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "', status =1 where IDTrans='" & .Text & "' and sl_no='" & .SubItems(4) & "'"
               
               '** update stock gudang [inventory tabel] ***
               SendDataToServer "insert into [Inventory tabel] (NoIdx,NoItem,Qty_in,Qty_out,refTrans,DateTrans," & _
                                                              "StockTmp,TypeTRans,lokasiGdg,sl_no) " & _
                                " values (newID(),'" & .SubItems(2) & "'," & FQty(.SubItems(5)) & ",0,'" & txtBox(0).Text & "','" & _
                                         Format(DTPicker1.Value, "yyyy-MM-dd") & "'," & FQty(.SubItems(5)) & ",'FG','" & txtBox(1).Text & _
                                         "','" & .SubItems(4) & "')"
               
               '** Update to MPS **
               UpdateMPS .SubItems(4), Val(.SubItems(5)), DTPicker1.Value
               
               '** UPdate Jurnal **'
               
            End If
         End With
      Else
         If MEdit Then
            rsBalance.DBOpen "select Qty_Out from [inventory tabel] where ln_no='" & ListFG.ListItems(x).SubItems(4) & "'", CNN
            If rsBalance.DBRecordset.Recordcount > 0 Then
               If rsBalance.DBRecordset.Fields(0) = 0 Then
                  '*** Update Stok Barang ***
                  SendDataToServer "delete from [inventory tabel] where sl_no='" & ListFG.ListItems(x).SubItems(4) & "'"
                  
                  '*** Update item FinishGood ***
                  SendDataToServer "update backflush_output set status=0 where sl_no='" & ListFG.ListItems(x).SubItems(4) & "'"
                  
                  '*** Update FinishGood ***
                  'SendDataToServer "update finishgood set received_by='' where ID='" & txtBox(0).Text & "'"
                  
               Else
                  MessageBox "Data LOT : " & ListFG.ListItems(x).SubItems(4) & " Tidak bisa diupdate !, Stok sudah terpakai"
               End If
            Else
               MessageBox "Data LOT : " & ListFG.ListItems(x).SubItems(4) & " Tidak bisa diupdate !"
            End If
         End If
      End If
   Next
End If
rsBalance.CloseDB
End Sub


Private Sub PrepareQuery()

On Error Resume Next
Dim strSQL As String
With MyDDE
    strSQL = "Update backflush_line set status=1 where IDTrans='xxxxxxxxxxxxxxx'"
    .PrepareAppend = strSQL
                     
    strSQL = "Update backflush_line set status=1 where IDTrans='xxxxxxxxxxxxxxx'"
    .PrepareUpdate = strSQL
                     
    .PrepareDelete = "Update backflush_line set status=0, orderID='' WHERE (OrderID = 'xxxxxxxxxxxxxxx')"
End With
Err.Clear
End Sub


Private Sub UpdateMPS(pNoItem As String, pQty As Double, pDate As Date)
   Dim rsCek As New DBQuick
   Dim MPSID As String
   
   rsCek.DBOpen "select * from view_MPS where getdate() between [require Date] and [end date]", CNN
   If rsCek.Recordcount > 0 Then
      MPSID = rsCek.DBRecordset.Fields(0)
   Else
      MPSID = ""
   End If
   
   
   SendDataToServer "update [MPS Detail] set list_value1=" & pQty & " where no_MPS='" & MPSID & "' and fcast_item='ACTUAL' and noItem='" & pNoItem & "' and time_days=" & Format(pDate, "dd")
   
   Set rsCek = Nothing
End Sub




