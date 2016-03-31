VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Guserid As Integer
Private akses As New utility
Public rsAkses As ADODB.Recordset
Public rsAksesMenu As ADODB.Recordset



Public Sub Login(ByVal Username As String, ByVal pass As String)
On Error GoTo LogErr
Dim Rc As New DBQuick
Dim strSQL, txtDenkrip As String

'strSQL = "SELECT user_table_line.* FROM         [Detail User Table] INNER JOIN    [User Table] ON [Detail User Table].[User ID] = [User Table].[User ID] INNER JOIN    [Form Table] ON [Detail User Table].Idx = [Form Table].Idx WHERE     ([User Table].[User Name] = N'" & Username & "') ORDER BY [Form Table].[Form List]"
Guserid = 0
txtDenkrip = EncodeStr64(PROBAencodeString(pass, Key, True), 68) 'membaca data Enkrip password
strSQL = "SELECT  [User Table].[User ID], [User Table].[User Name], User_table_group.[Group Name], User_table_line.ID," & _
                  "User_table_line.[Group ID], User_table_line.[Form ID], User_table_line.Access, dbo.User_table_line.New, User_table_line.Edit," & _
                  "User_table_line.Del,[user table].dept, [user table].Name_Dept " & _
                  "FROM  [User Table] INNER JOIN " & _
                  "User_table_group ON [User Table].[Group User] = User_table_group.id INNER JOIN " & _
                  "User_table_line ON User_table_group.id = User_table_line.[Group ID] " & _
                  "WHERE  ([User Table].[User Name] = '" & Username & "') and ([user table].password='" & txtDenkrip & "')"

Rc.DBOpen strSQL, CNN, lckLockReadOnly
If Rc.DBRecordset.Recordcount > 0 Then
    Guserid = Rc.DBRecordset.Fields("user id")
    CurrentDept = IIf(IsNull(Rc.DBRecordset.Fields("dept")), "", Rc.DBRecordset.Fields("dept"))
    NamaDept = IIf(IsNull(Rc.DBRecordset.Fields("Name_Dept")), "", Rc.DBRecordset.Fields("name_dept"))
    GroupID = IIf(IsNull(Rc.DBRecordset.Fields("group ID")), "", Rc.DBRecordset.Fields("Group ID"))
    'digunakan untuk insert ke user_activity_monitor
'    SendDataToServer (" INSERT INTO User_activity_monitor " & _
'                                         " (Name_user, group_user)" & _
'                                         " VALUES ('" & Rc.DBRecordset.Fields("user name") & "', '" & Rc.DBRecordset.Fields("group name") & "')")
End If
Rc.CloseDB
Exit Sub

LogErr:
    MessageBox Err.Description, "Login Error", msgOkOnly, msgCrtical
End Sub

Public Function GetID() As Integer
   GetID = Guserid
End Function

'Public Function MayDo(FormName As String, userid As Integer) As Integer
Public Function MayDo(FormName As String) As Integer

Dim rs As ADODB.Recordset

strSQL = "SELECT  [User Table].[User Name], dbo.User_table_group.[Group Name], dbo.User_table_line.[Group ID], dbo.User_table_line.[Form ID], " & _
                "User_Table_Form.[Form Name], dbo.User_Table_Form.[Group Form], dbo.User_table_line.Access, dbo.User_table_line.New, " & _
                "User_table_line.Edit , dbo.User_table_line.Del " & _
                "FROM [User Table] INNER JOIN " & _
                "User_table_group ON dbo.[User Table].[Group User] = dbo.User_table_group.id INNER JOIN " & _
                "dbo.User_table_line ON dbo.User_table_group.id = dbo.User_table_line.[Group ID] INNER JOIN " & _
                "dbo.User_Table_Form ON dbo.User_table_line.[Form ID] = dbo.User_Table_Form.ID " & _
                "WHERE (dbo.User_Table_Form.[Form Name] = '" & FormName & "') AND (dbo.[User Table].[User ID] = " & Guserid & ")"


   Set rs = akses.OpenDB(strSQL)
   
   If rs.Recordcount > 0 Then
         If rs.Fields("access") = "True" Then
            If rs.Fields("new") = "False" And rs.Fields("edit") = "True" And rs.Fields("del") = "True" Then
                MayDo = 1 ' useraddnewdenied
            ElseIf rs.Fields("new") = "True" And rs.Fields("edit") = "True" And rs.Fields("del") = "False" Then
                MayDo = 2  'usereditdelete
            ElseIf rs.Fields("new") = "False" And rs.Fields("edit") = "False" And rs.Fields("del") = "True" Then
                MayDo = 3 'usereditaddew
            ElseIf rs.Fields("new") = "False" And rs.Fields("edit") = "True" And rs.Fields("del") = "False" Then
                MayDo = 4  'usereditdelete
            ElseIf rs.Fields("new") = "True" And rs.Fields("edit") = "False" And rs.Fields("del") = "False" Then
                MayDo = 6  'usereditdelete
            ElseIf rs.Fields("new") = "True" And rs.Fields("edit") = "False" And rs.Fields("del") = "True" Then
                MayDo = 0 'useredit
            ElseIf rs.Fields("new") = "True" And rs.Fields("edit") = "True" And rs.Fields("del") = "True" Then
                MayDo = 5
            Else
                MayDo = 3 'usereditaddew
            End If
        Else
            MayDo = 0   'digunakan untuk akses false
        End If
    End If
    
    
      
  '      UserEditDenied             = 0
  '      UserAddnewDenied           = 1
  '      UserDeleteDenied           = 2
  '      UserEditAddnewDenied       = 3
  '      UserAddnewDeleteDenied     = 4
  '      UserOk                     = 5
  '      UserEditDeleteDenied       = 6
      
      
End Function

'Public Sub setMenuAktif(userid As Integer, usergroup As String)
Public Sub setMenuAktif(userid As Integer)
Dim strSQL As String


strSQL = "SELECT  dbo.[User Table].[User ID], dbo.[User Table].[User Name], dbo.User_table_group.[Group Name], dbo.User_table_line.Access, dbo.User_table_line.New, " & _
        "dbo.User_table_line.Edit , dbo.User_table_line.Del, dbo.User_Table_Form.[Form Name],dbo.user_table_form.[Group Form] " & _
        "FROM dbo.[User Table] INNER JOIN " & _
        "dbo.User_table_group ON dbo.[User Table].[Group User] = dbo.User_table_group.id INNER JOIN " & _
        "dbo.User_table_line ON dbo.User_table_group.id = dbo.User_table_line.[Group ID] INNER JOIN " & _
        "dbo.User_Table_Form ON dbo.User_table_line.[Form ID] = dbo.User_Table_Form.ID " & _
        "WHERE (dbo.[User Table].[User ID] =" & userid & ")"
        '"WHERE (dbo.[User Table].[User ID] =" & userid & ") AND (dbo.User_table_group.[Group Name] = '" & usergroup & "') "

'Di gunakan untuk Toolbar Menu
 Set rsAksesMenu = akses.OpenDB(strSQL)
 If rsAksesMenu.Recordcount > 0 Then
 rsAksesMenu.MoveFirst
 Do While rsAksesMenu.EOF = False
      Select Case UCase(rsAksesMenu.Fields("group form"))
             Case "MASTER DATA":
                   MainMenu.mnMaster.Visible = True
                   MainMenu.Toolbar1.Buttons(1).Visible = True
                   MainMenu.Toolbar1.Buttons(2).Visible = True 'garis
             Case "PEMBELIAN":
                   MainMenu.mnPurchase.Visible = True
                   MainMenu.Toolbar1.Buttons(3).Visible = True
                   MainMenu.Toolbar1.Buttons(4).Visible = True 'garis
             Case "PENJUALAN":
                   MainMenu.mnMarketing.Visible = True
                   MainMenu.Toolbar1.Buttons(5).Visible = True
                   MainMenu.Toolbar1.Buttons(6).Visible = True 'garis
             Case "LOGISTIK":
                   MainMenu.mnLogistik.Visible = True
                   MainMenu.Toolbar1.Buttons(7).Visible = True
                   MainMenu.Toolbar1.Buttons(8).Visible = True 'garis
             Case "GUDANG RL":
                   MainMenu.MnGudang.Visible = True
                   MainMenu.Toolbar1.Buttons(9).Visible = True
                   MainMenu.Toolbar1.Buttons(10).Visible = True 'garis
             Case "PRODUCTION":
                   MainMenu.mnInventory.Visible = True
                   MainMenu.Toolbar1.Buttons(11).Visible = True
                   MainMenu.Toolbar1.Buttons(12).Visible = True 'garis
             Case "ACCOUNTING":
                   MainMenu.mnAkun.Visible = True
                   MainMenu.Toolbar1.Buttons(13).Visible = True
                   MainMenu.Toolbar1.Buttons(14).Visible = True 'garis
             Case "QC":
                   MainMenu.mnQuality.Visible = True
                   MainMenu.Toolbar1.Buttons(15).Visible = True
                   MainMenu.Toolbar1.Buttons(16).Visible = True 'garis
             Case "HRD":
                   MainMenu.mnHrd.Visible = True
                   MainMenu.Toolbar1.Buttons(17).Visible = True
                   MainMenu.Toolbar1.Buttons(18).Visible = True 'garis
             Case "MAINTENANCE":
                   MainMenu.mnMaintenance.Visible = True
                   MainMenu.Toolbar1.Buttons(19).Visible = True
                   MainMenu.Toolbar1.Buttons(20).Visible = True 'garis
             Case "REPORT":
                   MainMenu.Toolbar1.Buttons(21).Visible = True
     End Select
     rsAksesMenu.MoveNext
 Loop
 End If

'
'Digunakan untuk akses menu
 Set rsAkses = akses.OpenDB(strSQL)
 If rsAkses.Recordcount > 0 Then
 MainMenu.mnUserArea.Enabled = IIf(UCase(rsAkses.Fields("User Name")) = "SA", True, False)
 rsAkses.MoveFirst
 Do While rsAkses.EOF = False
      Select Case rsAkses.Fields("form name")
            '* Applicaton *
            Case "Report Administration":             MainMenu.mnAdmLap.Enabled = rsAkses.Fields("access")
            Case "Konfigurasi Report":                MainMenu.mnConfReport.Enabled = rsAkses.Fields("access")

            '* Master Data *
               '* Multi Currency *
                  Case "Setup Mata Uang":             MainMenu.mnExcCurrency.Enabled = rsAkses.Fields("access")
                  Case "Nilai Tukar Mata Uang":       MainMenu.mnExMaint.Enabled = rsAkses.Fields("access")

               Case "WareHouse":                      MainMenu.MnGudangLog.Enabled = rsAkses.Fields("access")
               Case "Product Costing":                MainMenu.MnProdCosting.Enabled = rsAkses.Fields("access")
                                                      
                                                      
            
                
            '* Inventory *
               Case "Kelompok":                        MainMenu.mnKelompok.Enabled = rsAkses.Fields("access")
               Case "Kartu Stok":                      MainMenu.mnInvCard.Enabled = rsAkses.Fields("access")
                                                     

               Case "Kartu Stok Manager":              MainMenu.mnInvManager.Enabled = rsAkses.Fields("access")
               Case "Kartu Stok Purchasing":           MainMenu.mnInvPurch.Enabled = rsAkses.Fields("access")
               Case "Kategori":                        MainMenu.mnItemCategories1.Enabled = rsAkses.Fields("access")
                                                     
               Case "Substitusi":                      MainMenu.mnItemReference.Enabled = rsAkses.Fields("access")
                                                     

            Case "Regional":                          MainMenu.mnRegional.Enabled = rsAkses.Fields("access")
            Case "Transporter":                       MainMenu.mnTransport.Enabled = rsAkses.Fields("access")

            '* Customer *'
               Case "Customer Card":                  MainMenu.mnCusMaster.Enabled = rsAkses.Fields("access")
               Case "Customer Warehouse":             MainMenu.mnCusGudang.Enabled = rsAkses.Fields("access")

            Case "Supplier Card":                     MainMenu.mnSup.Enabled = rsAkses.Fields("access")
            Case "Bank Partner":                      MainMenu.mnBankPartner.Enabled = rsAkses.Fields("access")
            Case "Unit Of Measurement":               MainMenu.mnUOM.Enabled = rsAkses.Fields("access")

            Case "Termin Pembayaran":                 MainMenu.mnterminpembayaran.Enabled = rsAkses.Fields("access")
            Case "Tipe Item Transaksi":               MainMenu.mnTipeItem.Enabled = rsAkses.Fields("access")
            Case "Detail Item Jasa":                  MainMenu.mnItemCharge.Enabled = rsAkses.Fields("access")
                
            
                    
            'Case "Shipment":                         MainMenu.mnTipeBayar.Enabled = rsAkses.Fields("access")
            'Case "Employee":                         MainMenu.mnKaryawan.Enabled = rsAkses.Fields("access")

            '* Menu Untuk Pembelian
               '* Order Proses
                  Case "Permintaan Pembelian":        MainMenu.MnRequest.Enabled = rsAkses.Fields("access") 'sdh frmlistspph
                  Case "Permintaan Penawaran Harga":  MainMenu.MnPriceOffer.Enabled = rsAkses.Fields("access") 'sdh frmspph
                  Case "Penawaran Harga":             MainMenu.mnPenawaranHrg.Enabled = rsAkses.Fields("access")
                  Case "Order Pembelian":             MainMenu.mnPurchaseOrder.Enabled = rsAkses.Fields("access")
                  Case "Blanked PO":                  MainMenu.mnPOBlanked.Enabled = rsAkses.Fields("access")
                  Case "Petty Cash":                  MainMenu.mnPOBlanked1.Enabled = rsAkses.Fields("access")
                  Case "Retur Pembelian":             MainMenu.mnPurchaseReturn.Enabled = rsAkses.Fields("access")
                  Case "Planned Order":               MainMenu.MnPlanOrder.Enabled = rsAkses.Fields("access")
                                                      MainMenu.mnProductionPlan.Enabled = MainMenu.MnPlanOrder.Enabled
                  Case "MRP Generation":              MainMenu.mnMRP.Enabled = rsAkses.Fields("access")

               '* Eksekusi *
                  Case "Outstanding Pembelian":       MainMenu.mnoutstand.Enabled = rsAkses.Fields("access")
                  Case "Invoice Pembelian":           MainMenu.MnInvPuchase.Enabled = rsAkses.Fields("access")
                  Case "Evaluasi Supplier":           MainMenu.MnEvaluasi.Enabled = rsAkses.Fields("access")
                  'Case "Billing PPN": MainMenu.MnPPN.Enabled = rsAkses.Fields("access")

               '* Approval *
                  Case "App Permintaan Pembelian":    MainMenu.mnAppSPP.Enabled = rsAkses.Fields("access")
                  Case "App Rencana Pembelian Bulanan":   mnAPPRPBRL = rsAkses.Fields("access")
                  Case "App Evaluasi Supplier":       mnAPPEvaSupp = rsAkses.Fields("access")
                  Case "App Surat Jalan":             mnAPPSuratJalan = rsAkses.Fields("access")
                  Case "App Surat Retur":             mnAPPSuratRetur = rsAkses.Fields("access")
                  Case "App Order Pembelian":         mnBApprovalPO = rsAkses.Fields("access")
                  Case "App Permintaan Penawaran Harga":      mnAPPPermintaanHarga = rsAkses.Fields("access")
                  Case "App Rekap permintaan Pembelian":      mnAPPRKPPermintaanBeli = rsAkses.Fields("access")
                  Case "App Surat Pengembalian Barang":       mnAPPSPB = rsAkses.Fields("access")
                  Case "App Laporan Pembelian Bahan / Bulan": mnAPPLPB = rsAkses.Fields("access")
                  
                  
                  
                
               '* Histroy *
                  Case "History Permintaan Penawaran Harga": MainMenu.MnHSPPH.Enabled = rsAkses.Fields("access")
                  Case "History Permintaan Barang":   MainMenu.MnHSPP.Enabled = rsAkses.Fields("access")
                  Case "History Retur Pembelian":     MainMenu.MnHReturBeli.Enabled = rsAkses.Fields("access")
                  Case "History Planned Order MRP":   MainMenu.mnHMRP.Enabled = rsAkses.Fields("access")
                  Case "History Invoice Pembelian":   MainMenu.MnHInvoiceBeli.Enabled = rsAkses.Fields("access")
                  Case "History Billing PPN":         MainMenu.MnhPPN.Enabled = rsAkses.Fields("access")
                  Case "History Order Pembelian":     MainMenu.mnHPO = rsAkses.Fields("access")

            '* Menu Penjualan
               '* Info *
                  Case "Customer":                    MainMenu.mnSalesCust.Enabled = rsAkses.Fields("access")
                  Case "Sales Team":                  MainMenu.mnSalesTeam.Enabled = rsAkses.Fields("access")

               '* Order Proses *
                  Case "Sales Forcast":               MainMenu.MnSalesFCast.Enabled = rsAkses.Fields("access")
                  Case "Permintaan Barang":           MainMenu.mnSalesCust.Enabled = rsAkses.Fields("access")
'
                                                      
                  
                  Case "Sales Quote":                 MainMenu.mnSalesQuote.Enabled = rsAkses.Fields("access")
                  Case "Order Penjualan":             MainMenu.mnSalesOrder.Enabled = rsAkses.Fields("access")
                  Case "Kontrak Penjualan":           MainMenu.mnContractManagement.Enabled = rsAkses.Fields("access")
                  Case "Outstanding Penjualan":       MainMenu.mnOutstandingMKT.Enabled = rsAkses.Fields("access")
               
               '* Marketing *
                  Case "Contact":                     MainMenu.mnMarketAutoCamp.Enabled = rsAkses.Fields("access")
                  Case "Permintaan Sample":           MainMenu.mnpermintaansample.Enabled = rsAkses.Fields("access")
                  Case "Memo Potongan Harga":         MainMenu.mnmemopotonganharga.Enabled = rsAkses.Fields("access")
                  Case "Customer Feedback":           MainMenu.mncustomerfeedback.Enabled = rsAkses.Fields("access")


               '* Approval
                  Case "App Sales Order":             MainMenu.mnAPPSO.Enabled = rsAkses.Fields("access")
                  Case "App Permintaan Sample":       mnAPPPermintaanSample = rsAkses.Fields("access")
                  Case "App Memo Potongan Harga":     mnAPPMemo = rsAkses.Fields("access")
                  Case "App Customer Feedback":       mnAPPCustomerFeedback = rsAkses.Fields("access")
              
               '* History
                  Case "History Sales Quote":        MainMenu.mnHSalesQuote.Enabled = rsAkses.Fields("access")
                  Case "History Kontrak Penjualan":  MainMenu.mnHkontrakPenjualan.Enabled = rsAkses.Fields("access")
                  Case "History Order Penjualan":    MainMenu.mnHOrderPenjualan.Enabled = rsAkses.Fields("access")
                  Case "History Invoice Penjualan":  MainMenu.mnHInvoicePenjualan.Enabled = rsAkses.Fields("access")


           '* Menu Logistik *
               Case "Permintaan Barang Logistik":              MainMenu.mnPermintaanBrg.Enabled = rsAkses.Fields("access")
'

               Case "Permintaan Pembelian Logistik":           MainMenu.mnPermintaanBeli.Enabled = rsAkses.Fields("access")

               '* Penerimaan *
                  Case "Penerimaan Bahan Penunjang":  MainMenu.mnBPenunjang.Enabled = rsAkses.Fields("access")
                  Case "Penerimaan Barang Jadi":      MainMenu.mnBJadi.Enabled = rsAkses.Fields("access")
                  Case "Penerimaan Retur Customer":  MainMenu.mnReturCust.Enabled = rsAkses.Fields("access")

               '* Pengeluaran *
                  Case "Pengeluaran Barang":          MainMenu.mnBKeluar.Enabled = rsAkses.Fields("access")
                  Case "Pengiriman Export":           MainMenu.mnpengirimanexport.Enabled = rsAkses.Fields("access")
                  Case "Retur Supplier":              MainMenu.mnReturSupp.Enabled = rsAkses.Fields("access")
                  Case "Surat Jalan":                 MainMenu.mnDO.Enabled = rsAkses.Fields("access")

                  Case "Stock Opname":                mnStockOpname = rsAkses.Fields("access")
                  Case "Stock Browser":               mnBInventoryBrowser = rsAkses.Fields("access")
                  Case "Inventory Adjustment":        mnBInventoryAdj = rsAkses.Fields("access")
                  
               '* Approval *
                  Case "App Permintaan Barang":       MainMenu.mnAPSPB.Enabled = rsAkses.Fields("access")
                  Case "App Permintaan Pembelian":    mnBApprovalSPP = rsAkses.Fields("access")

            '* Menu Gudang RL
               Case "Permintaan Barang Gudang":       MainMenu.MnGudangMinta.Enabled = rsAkses.Fields("access")
               Case "Tanda Terima RL":                MainMenu.MnTerimaRL.Enabled = rsAkses.Fields("access")
               Case "Pengiriman RL":                  MainMenu.MnKirimRL.Enabled = rsAkses.Fields("access")
               Case "Lembar Supplier":                MainMenu.mnLembarSupplier.Enabled = rsAkses.Fields("access")
               Case "Konfigurasi RL Batch":           MainMenu.mnKOnfigurasiRLBatch.Enabled = rsAkses.Fields("access")
               
               '* approval
               Case "App Tanda Terima RL":            mnAPPTTRL = rsAkses.Fields("access")
               Case "App Pengiriman RL":              mnAPPKRL = rsAkses.Fields("access")
               Case "App Lembar Supplier":            mnAPPLembSupp = rsAkses.Fields("access")
               
            '* Menu Accounting
               'Setup
               Case "Daftar Perkiraan":               MainMenu.mnPerkiraan.Enabled = rsAkses.Fields("access")
               Case "Periode Transaksi":              MainMenu.mnPeriode.Enabled = rsAkses.Fields("access")
               Case "Konfigurasi Rekening":           MainMenu.mnSetupAccount.Enabled = rsAkses.Fields("access")
               Case "Konfigurasi Biaya":              MainMenu.mncostmethode.Enabled = rsAkses.Fields("access")
               Case "Item Harga Pokok":               MainMenu.mnproductcostingAcc.Enabled = rsAkses.Fields("access")
               
                      
               'Transaction
               Case "Permintaan Barang Acc":            MainMenu.mnpermintaanbarangACC.Enabled = rsAkses.Fields("access")
               Case "Transfer Kas":                     MainMenu.mnTukasKas.Enabled = rsAkses.Fields("access")
               Case "Penerimaan Kas":                   MainMenu.mnBkm.Enabled = rsAkses.Fields("access")
               Case "Pengeluaran Kas":                  MainMenu.mnBkk.Enabled = rsAkses.Fields("access")
               Case "Voucher Pembelian":                MainMenu.Vp.Enabled = rsAkses.Fields("access")
               Case "Voucher Penjualan":                MainMenu.Vpe.Enabled = rsAkses.Fields("access")
               Case "Invoice Penjualan":                MainMenu.MnInvSales.Enabled = rsAkses.Fields("access")
               Case "Invoice Pembelian":                MainMenu.MnInvPuchase.Enabled = rsAkses.Fields("access")
                                                  
               'Memorial
               Case "Memorial Journal":               MainMenu.mnMemoUmum.Enabled = rsAkses.Fields("access")
               Case "Sales / Purchase":               MainMenu.mnMemoJualbeli.Enabled = rsAkses.Fields("access")
                              
               'Revalution
               Case "Sales Series":                   MainMenu.mnRevSales.Enabled = rsAkses.Fields("access")
               Case "Purchase Series":                MainMenu.mnRevPurchase.Enabled = rsAkses.Fields("access")
                             
               'Closing
               Case "Period Closing":                 MainMenu.mnClosing.Enabled = rsAkses.Fields("access")
               'Case "Harga Pokok Penjualan":          MainMenu.mnhargapokokpenjualan.Enabled = rsAkses.Fields("access")

            '* Menu Production
               '* Setup
                  Case "Resource Type":               MainMenu.mnRscType.Enabled = rsAkses.Fields("access")
                  Case "Scheduling Calender":         MainMenu.mnCalendar.Enabled = rsAkses.Fields("access")
                  Case "Work Centers":                MainMenu.mnManWC.Enabled = rsAkses.Fields("access")
                  Case "Routing":                     MainMenu.mnCountPoint.Enabled = rsAkses.Fields("access")
                  Case "Resources":                   MainMenu.mnRsc.Enabled = rsAkses.Fields("access")
                  Case "BOM Methode":                 MainMenu.mnBomMethode.Enabled = rsAkses.Fields("access")
                  Case "Product Costing":             MainMenu.mnJobCosting.Enabled = rsAkses.Fields("access")
                  Case "Bill Of Material":            MainMenu.mnBomBom.Enabled = rsAkses.Fields("access")
                  Case "Engineering Change":          MainMenu.mnECC.Enabled = rsAkses.Fields("access")

               '* Inventory
                  Case "Inventory Card Produksi":       MainMenu.mnMInventory.Enabled = rsAkses.Fields("access")
                  Case "Inventory Reference":           MainMenu.mnItemReference.Enabled = rsAkses.Fields("access")
                  Case "Inventory Categories Produksi": MainMenu.mnItemCategories.Enabled = rsAkses.Fields("access")
                  Case "Master Oursourced":             MainMenu.mnDescrip.Enabled = rsAkses.Fields("access")
                  Case "Outsourced Type":               MainMenu.mnTipeDes.Enabled = rsAkses.Fields("access")
                  Case "Outsourced Reference":          MainMenu.mnDescripRef.Enabled = rsAkses.Fields("access")

               '* Production Planning
                  Case "Manufacturing Order":         MainMenu.mnManOrder.Enabled = rsAkses.Fields("access")
                  Case "Master Production Schedule":  MainMenu.mnSchedule.Enabled = rsAkses.Fields("access")
                  Case "Master Requirement Planning": MainMenu.mnMRPlanning.Enabled = rsAkses.Fields("access")
                  Case "Capacity Planning":           MainMenu.mnCapaPlan.Enabled = rsAkses.Fields("access")
                  Case "Planned Order":               MainMenu.mnProductionPlan.Enabled = rsAkses.Fields("access")

               '* Shop floor
                  Case "Material Requestation":       MainMenu.mnMaterialRequisition.Enabled = rsAkses.Fields("access")
                  Case "Material Issue":              MainMenu.mnMaterialIssue.Enabled = rsAkses.Fields("access")
                  Case "Proses Produksi":             MainMenu.mnProsesProd.Enabled = rsAkses.Fields("access")
                  Case "Proses Mixing & Milling":     MainMenu.mnPrelot.Enabled = rsAkses.Fields("access")
                  Case "Pengiriman Chip":             MainMenu.mnMutasiChip.Enabled = rsAkses.Fields("access")
                  Case "Blending Instruction":        MainMenu.mnLot1.Enabled = rsAkses.Fields("access")
                  Case "BackFlush WIP":               MainMenu.mnBackFlushing.Enabled = rsAkses.Fields("access")
                  Case "Status in Proses":            MainMenu.mnStatusIP.Enabled = rsAkses.Fields("access")
                  Case "Serah Terima Produk Jadi":    MainMenu.mnSTPJadi.Enabled = rsAkses.Fields("access")

               '* Approval
                  Case "App Alkali Treatment":        mnAPPALKALI = rsAkses.Fields("access")
                  Case "App Acid Treatment":          mnAPPACID = rsAkses.Fields("access")
                  Case "App Bleaching Treatment":     mnAPPBLEACHING = rsAkses.Fields("access")
                  Case "App Ekstraksi di Reaktor":    mnAPPEKSREAKTOR = rsAkses.Fields("access")
                  Case "App Ekstraksi di Autoclave":  mnAPPEKSAUTO = rsAkses.Fields("access")
                  Case "App Filter Press":            mnAPPFILTERPRESS = rsAkses.Fields("access")
                  Case "App Gellification":           mnAPPGELL = rsAkses.Fields("access")
                  Case "App Pembungkusan":            mnAPPBUNGKUS = rsAkses.Fields("access")
                  Case "App Concrete Press":          mnAPPCONCRETE = rsAkses.Fields("access")
                  Case "App Hydraulic Press":         mnAPPHYDRAULIC = rsAkses.Fields("access")
                  Case "App Cutter":                  mnAPPCUTTER = rsAkses.Fields("access")
                  Case "App Penjemuran dan Dryer":    mnAPPJEMUR = rsAkses.Fields("access")
                  Case "App Crusher":                 mnAPPCRUSHER = rsAkses.Fields("access")
                  Case "App Mixing & Milling Chip":   mnAPPMIXING = rsAkses.Fields("access")
                  Case "App Blending Instruction":    mnAPPBLENDING = rsAkses.Fields("access")
                  Case "App Serah terima Produk Jadi":  mnAPPSTFG = rsAkses.Fields("access")
                  Case "App Pemakaian Bahan Produksi":  mnAPPPAKAI = rsAkses.Fields("access")
                  
               
               '* History

                  Case "Alkali Treatment":            MainMenu.mnHistAlkali.Enabled = rsAkses.Fields("access")
                  Case "Acid Treatment":              MainMenu.mnHistAcid.Enabled = rsAkses.Fields("access")
                  Case "Bleaching Treatment":         MainMenu.mnHistBleaching.Enabled = rsAkses.Fields("access")
                  Case "Ekstraksi di Autoclave":      MainMenu.mnHistAutoclave.Enabled = rsAkses.Fields("access")
                  Case "Ekstraksi di Reaktor":        MainMenu.mnHistReaktor.Enabled = rsAkses.Fields("access")
                  Case "Filter Press":                MainMenu.mnHistFilter.Enabled = rsAkses.Fields("access")
                  Case "Pembungkusan":                MainMenu.mnHistBungkus.Enabled = rsAkses.Fields("access")
                  Case "Concrete Press":              MainMenu.mnHistConcrete.Enabled = rsAkses.Fields("access")
                  Case "Hydraulic Press":             MainMenu.mnHistHydraulic.Enabled = rsAkses.Fields("access")
                  Case "Penjemuran":                  MainMenu.mnHistJemur.Enabled = rsAkses.Fields("access")
                  Case "Dryer":                       MainMenu.mnHistDryer.Enabled = rsAkses.Fields("access")
                  Case "Crusher":                     MainMenu.mnHistCrusher.Enabled = rsAkses.Fields("access")
                    
    End Select
    rsAkses.MoveNext
 Loop
End If
AktifParentMenu   'Digunakan Apabila Chield Menu Tdk Ada maka Parent Menu Automatis Visible
End Sub


Private Sub AktifParentMenu()
'Di Gunakan Untuk Chek Parent Menu Apabila Sub Menu Tdk Ada Maka Parent Tdk Di Tampilkan
'*Master
'Multi Currency
With MainMenu
    If (.mnExcCurrency.Enabled = False) And (.mnExMaint.Enabled = False) Then
        .mnCurrency.Enabled = False  'Menu Multi Currency menu Parent
    Else
        .mnCurrency.Enabled = True ' Menu Multi Currency Menu Parent
    End If
    
    'Inventory
    If (.mnInvCard.Enabled = False) And (.mnInvManager.Enabled = False) And (.mnInvPurch.Enabled = False) And (.mnKelompok.Enabled = False) And (.mnItemCategories1.Enabled = False) And (.mntemReference1.Enabled = False) Then
        .mnItem.Enabled = False  'Menu Inventory Menu Parent
    Else
        .mnItem.Enabled = True 'Menu Inventory Menu Parent
    End If
    
    'Customer
    If (.mnCusMaster.Enabled = False) And (.mnCusGudang.Enabled = False) Then
        .mncus.Enabled = False ' Menu Customer Menu Parent
    Else
        .mncus.Enabled = True ' Menu Customer Menu Parent
    End If
    
    '*Pembelian
    'Order Proses
    If (.MnRequest.Enabled = False) And (.MnPriceOffer.Enabled = False) And (.mnPenawaranHrg.Enabled = False) And (.mnPurchaseOrder.Enabled = False) And (.mnPOBlanked.Enabled = False) And (.mnPOBlanked1.Enabled = False) And (.mnPurchaseReturn.Enabled = False) And (.MnPlanOrder.Enabled = False) And (.mnMRP.Enabled = False) Then
        .mnorderproses.Enabled = False  'Menu Order Proses Menu Parent
    Else
        .mnorderproses.Enabled = True 'Menu Order Proses Menu Parent
    End If
    
    'Eksekusi
    If (.mnoutstand.Enabled = False) And (.MnInvPuchase.Enabled = False) And (.MnEvaluasi.Enabled = False) Then
        .mneksekusi.Enabled = False
    Else
        .mneksekusi.Enabled = True
    End If
    
    'Approval
    If (.mnAppSPP.Enabled = False) And (mnAPPEvaSupp = False) And (mnAPPSuratJalan = False) And (mnAPPSuratRetur = False) And (mnAPPlembarSupp = False) And (mnBApprovalPO = False) And (mnAPPPermintaanHarga = False) And (mnAPPRKPPermintaanBeli = False) And (mnAPPSPB = False) And (mnAPPLPB = False) Then
       .mnapproval.Enabled = False
    Else
        .mnapproval.Enabled = True
    End If
    
    
        
    'history
    If (.MnHSPP.Enabled = False) And (.MnHReturBeli.Enabled = False) And (.mnHMRP.Enabled = False) And (.MnHInvoiceBeli.Enabled = False) And (.MnhPPN.Enabled = False) And (.mnHPO.Enabled = False) Then
        .MnHistoryPurchase.Enabled = False  'Menu History Menu Parent
    Else
        .MnHistoryPurchase.Enabled = True 'Menu History Menu Parent
    End If
    
    '*Penjualan
    'Info
    If .mnSalesCust.Enabled = False And .mnSalesTeam.Enabled = False Then
       .MnInfo.Enabled = False
    Else
        .MnInfo.Enabled = True
    End If
    
    'Order Proses
    If (.MnSalesFCast.Enabled = False) And (.mnpermintaanbarang.Enabled = False) And (.mnSalesQuote.Enabled = False) And (.mnSalesOrder.Enabled = False) And (.mnContractManagement.Enabled = False) And (.MnInvSales.Enabled = False) Then
        .mnSalesAutomation.Enabled = False
    Else
        .mnSalesAutomation.Enabled = True
    End If
    
    'marketing
    If (.mnMarketAutoCamp.Enabled = False) And (.mnpermintaansample.Enabled = False) And (.mnmemopotonganharga.Enabled = False) And (.mncustomerfeedback.Enabled = False) Then
        .mnMarketAuto.Enabled = False
    Else
        .mnMarketAuto.Enabled = True
    End If
    
    If (.mnAPPSO.Enabled = False) And (mnAPPPermintaanSample = False) And (mnAPPMemo = False) And (mnAPPCustomerFeedback = False) Then
        .mnapprovalMKT.Enabled = False
    Else
        .mnapprovalMKT.Enabled = True
    End If
    
      
    
    '*Logistik
    'penerimaan
    If (.mnBPenunjang.Enabled = False) And (.mnBJadi.Enabled = False) And (.mnReturCust.Enabled = False) Then
       .mnxx1.Enabled = False
    Else
        .mnxx1.Enabled = True
    End If
    
    'Pengeluaran
    If (.mnBKeluar.Enabled = False) And (.mnpengirimanexport.Enabled = False) And (.mnReturSupp.Enabled = False) Then
       .mnXX2.Enabled = False
    Else
        .mnXX2.Enabled = True
    End If
    
    
    'Approval
    If .mnAPSPB.Enabled = False Then
       .mnApproval1.Enabled = False
    Else
        .mnApproval1.Enabled = True
    End If
    
    
                 
    'Gudang RL
    
    If (.MnGudangMinta.Enabled = False) And (.MnTerimaRL.Enabled = False) And (.MnKirimRL.Enabled = False) And (.mnLembarSupplier.Enabled = False) Then
        .MnGudang.Enabled = False
    Else
        .MnGudang.Enabled = True
    End If
    
    
    '*Produksi
    'Setup
    If (.mnRscType.Enabled = False) And (.mnCalendar.Enabled = False) And (.mnManWC.Enabled = False) And (.mnCountPoint.Enabled = False) And (.mnRsc.Enabled = False) And (.mnBomMethode.Enabled = False) And (.mnJobCosting.Enabled = False) And (.mnBomBom.Enabled = False) And (.mnECC.Enabled = False) Then
       .mnEntupProduksi.Enabled = False
    Else
        .mnEntupProduksi.Enabled = True
    End If
    
    'inventory
    If (.mnMInventory.Enabled = False) And (.mnItemReference.Enabled = False) And (.mnItemCategories.Enabled = False) And (.mnDescrip.Enabled = False) And (.mnTipeDes.Enabled = False) And (.mnDescripRef.Enabled = False) Then
       .mnMenuPersediaan.Enabled = False
    Else
        .mnMenuPersediaan.Enabled = True
    End If
    
    'Production Planning
    If (.mnSchedule.Enabled = False) And (.mnMRPlanning.Enabled = False) And (.mnCapaPlan.Enabled = False) And (.mnProductionPlan.Enabled = False) Then
        .mnMenuProduksi.Enabled = False
    Else
        .mnMenuProduksi.Enabled = True
    End If
    
              
    'Shop Floor
    If (.mnMaterialRequisition.Enabled = False) And (.mnMaterialIssue.Enabled = False) And (.mnProsesProd.Enabled = False) And (.mnPrelot.Enabled = False) And (.mnMutasiChip.Enabled = False) And (.mnLot1.Enabled = False) And (.mnBackFlushing.Enabled = False) And (.mnStatusIP.Enabled = False) And (.mnSTPJadi.Enabled = False) Then
        .mnShopFloor.Enabled = False
    Else
        .mnShopFloor.Enabled = True
    End If
    
    
    'History
    If (.mnHistAlkali.Enabled = False) And (.mnHistAcid.Enabled = False) And (.mnHistBleaching.Enabled = False) And (.mnHistAutoclave.Enabled = False) And (.mnHistReaktor.Enabled = False) And (.mnHistFilter.Enabled = False) And (.mnHistBungkus.Enabled = False) And (.mnHistConcrete.Enabled = False) And (.mnHistHydraulic.Enabled = False) And (.mnHistJemur.Enabled = False) And (.mnHistDryer.Enabled = False) And (.mnHistCrusher.Enabled = False) Then
        .mnAsmHistory.Enabled = False
    Else
        .mnAsmHistory.Enabled = True
    End If
    
       
    '*Accounting
    'Setup
    
    If (.mnPerkiraan.Enabled = False) And (.mnPeriode.Enabled = False) And (.mnSetupAccount.Enabled = False) And (.mncostmethode.Enabled = False) And (.mnproductcostingAcc.Enabled = False) Then
        .mnKOnfig.Enabled = False
    Else
        .mnKOnfig.Enabled = True
    End If
    
    'Transaction
    If (.mnpermintaanbarangACC.Enabled = False) And (.mnBkm.Enabled = False) And (.mnBkk.Enabled = False) And (.Vp.Enabled = False) And (.Vpe.Enabled = False) And (.MnInvSales.Enabled = False) And (.MnInvPuchase.Enabled = False) And (.mnTukasKas.Enabled = False) Then
       .mnKass.Enabled = False
    Else
        .mnKass.Enabled = True
    End If
         
    
    'Memorial
    If (.mnMemoUmum.Enabled = False) And (.mnMemoJualbeli.Enabled = False) Then
        .mnMemorial.Enabled = False
    Else
        .mnMemorial.Enabled = True
    End If
    
    'Revaluation
    If (.mnRevSales.Enabled = False) And (.mnRevPurchase.Enabled = False) Then
        .mnReval.Enabled = False
    Else
        .mnReval.Enabled = True
    End If
    
    'Closing
'    If (.mnClosing.Enabled = False) And (.mnhargapokokpenjualan.Enabled = False) Then
'        .mnClosed.Enabled = False
'    Else
'        .mnClosed.Enabled = True
'    End If
End With
End Sub

Public Sub Del_activity()
Dim RcCek As New DBQuick
    RcCek.DBOpen "select id,name_user from user_activity_monitor", CNN, lckLockReadOnly
    RcCek.MoveLastRecord
    SendDataToServer (" delete from [user_activity_monitor] where id=" & RcCek.DBRecordset.Fields("id") & "")
                                         
End Sub
