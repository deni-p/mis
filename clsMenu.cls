VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Sub CreateMenu(ByVal MenuOrder As String)
If Isregis = False Then Exit Sub
With MainMenu.SemeruTree1
     .MenuTreeView.Nodes.Clear
     Set .MenuTreeView.ImageList = MainMenu.ImageList1
         .BackColorTree = &H6D4016
     'If IsEnabledLogin = False Then Exit Sub
     Select Case UCase(MenuOrder)
            Case "MASTER":
                If MainMenu.mnMaster.Enabled = True Then .NodeAdd , tvwChild, "Master", "Master Data", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnCurrency.Enabled = True Then .NodeAdd "Master", tvwChild, "MasterCurrency", "Multi Currency", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnExcCurrency.Enabled = True Then .NodeAdd "MasterCurrency", tvwChild, "CurrSetup", "Setup Mata Uang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnExMaint.Enabled = True Then .NodeAdd "MasterCurrency", tvwChild, "ExcMaintenance", "Nilai Tukar Mata Uang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnGudangLog.Enabled = True Then .NodeAdd "Master", tvwChild, "MasterGudang", "Warehouse", "biru", , , , , , True, , &HFCF1ED, &H6D4016
               
                
                If MainMenu.mnItem.Enabled = True Then .NodeAdd "Master", tvwChild, "MasterItem", "Inventory", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnInvCard.Enabled = True Then .NodeAdd "MasterItem", tvwChild, "InvCard", "Kartu Stok", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnInvManager.Enabled = True Then .NodeAdd "MasterItem", tvwChild, "InvManager", "Kartu Stok Manager", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnInvPurch.Enabled = True Then .NodeAdd "MasterItem", tvwChild, "InvPurch", "Kartu Stok Purchasing", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnKelompok.Enabled = True Then .NodeAdd "MasterItem", tvwChild, "MasterKelompok", "Kelompok", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnItemCategories1.Enabled = True Then .NodeAdd "MasterItem", tvwChild, "Categories", "Kategori", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnItemReference.Enabled = True Then .NodeAdd "MasterItem", tvwChild, "IR", "Substitusi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                If MainMenu.mnRegional.Enabled = True Then .NodeAdd "Master", tvwChild, "MasterRegional", "Regional", "biru", , , , , , True, , &HFCF1ED, &H6D4016
               
                If MainMenu.mnTransport.Enabled = True Then .NodeAdd "Master", tvwChild, "EntriTransport", "Transporter", "biru", , , , , , True, , &HFCF1ED, &H6D4016
               
                If MainMenu.mncus.Enabled = True Then .NodeAdd "Master", tvwChild, "EntriCustomer", "Customer", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnCusMaster.Enabled = True Then .NodeAdd "EntriCustomer", tvwChild, "CustCard", "Customer Card", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnCusGudang.Enabled = True Then .NodeAdd "EntriCustomer", tvwChild, "EntriGudang", "Customer Warehouse", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSup.Enabled = True Then .NodeAdd "Master", tvwChild, "EntriSupplier", "Supplier Card", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBankPartner.Enabled = True Then .NodeAdd "Master", tvwChild, "EntriBank", "Bank Partner", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnterminpembayaran.Enabled = True Then .NodeAdd "Master", tvwChild, "TermPayment", "Termin Pembayaran", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnTipeItem.Enabled Then .NodeAdd "Master", tvwChild, "TipeItem", "Tipe Item Transaksi", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                'edit yudha
                If MainMenu.mnItemCharge.Enabled Then .NodeAdd "Master", tvwChild, "ItemCharge", "Detail Item Jasa", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnUOM.Enabled = True Then .NodeAdd "Master", tvwChild, "EntryUOM", "Unit Of Measurement", "biru", , , , , , True, , &HFCF1ED, &H6D4016
               
           Case "PURCHASE":
                If MainMenu.mnPurchase.Enabled = True Then .NodeAdd , tvwChild, "PurchTrans", "Pembelian", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
                
                'ORDER
                If MainMenu.mnorderproses.Enabled = True Then .NodeAdd "PurchTrans", tvwChild, "OrderPurch", "Order Proses", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnRequest.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "LogRequest", "Permintaan Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnPriceOffer.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "SPPH", "Permintaan Penawaran Harga", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPenawaranHrg.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "TransaksiSPPH", "Penawaran Harga", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPurchaseOrder.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "TransaksiPO", "Order Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPOBlanked.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "POBlanked", "Blanket PO", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPOBlanked1.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "POBlanked1", "Blanket PO Petty Cash", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPurchaseReturn.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "PurchRetur", "Retur Pembelian ", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnPlanOrder.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "PlanOrder", "Planned Order MRP", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMRP.Enabled = True Then .NodeAdd "OrderPurch", tvwChild, "MRPGEN", "MRP Generation", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'EKSEKUSI
                If MainMenu.mneksekusi.Enabled = True Then .NodeAdd "PurchTrans", tvwChild, "OrderNext", "Eksekusi", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.OutSPH.Enabled = True Then .NodeAdd "OrderNext", tvwChild, "PurchOutSPH", "Outstanding SPH", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnoutstand.Enabled = True Then .NodeAdd "OrderNext", tvwChild, "PurchOutstanding", "Outstanding Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                'If MainMenu.MnInvPuchase.Enabled = True Then .NodeAdd "OrderNext", tvwChild, "TransaksiAP", "Invoice Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnEvaluasi.Enabled = True Then .NodeAdd "OrderNext", tvwChild, "PurchEvaluasi", "Evaluasi Supplier", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                
                'APPROVAL
                If MainMenu.mnapproval.Enabled = True Then .NodeAdd "PurchTrans", tvwChild, "PurchApproval", "Approval", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnAppSPP.Enabled Then .NodeAdd "PurchApproval", tvwChild, "APSPP", "Permintaan Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                            
                If mnAPPEvaSupp = True Then .NodeAdd "PurchApproval", tvwChild, "APPEVASUPP", "Evaluasi Supplier", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPSuratJalan = True Then .NodeAdd "PurchApproval", tvwChild, "APPSURATJALAN", "Surat Jalan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPSuratRetur = True Then .NodeAdd "PurchApproval", tvwChild, "APPSURATRETUR", "Surat Retur", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnBApprovalPO = True Then .NodeAdd "PurchApproval", tvwChild, "APPORDPEMB", "Order Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPRPBRL = True Then .NodeAdd "PurchApproval", tvwChild, "APPRPBRL", "Rencana Pembelian Bulanan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPPermintaanHarga = True Then .NodeAdd "PurchApproval", tvwChild, "APPPERHARGA", "Permintaan Penawaran Harga", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPRKPPermintaanBeli = True Then .NodeAdd "PurchApproval", tvwChild, "APPRKPBELI", "Rekap Permintaan Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPSPB = True Then .NodeAdd "PurchApproval", tvwChild, "APPSRTBARANG", "Surat Pengembalian Barang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPLPB = True Then .NodeAdd "PurchApproval", tvwChild, "APPLAPBAHAN", "Laporan Pembelian Bahan / Bulan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                            
                            
                'HISTORY
                If MainMenu.MnHistoryPurchase.Enabled = True Then .NodeAdd "PurchTrans", tvwChild, "PurchHistory", "History", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnHSPPH.Enabled = True Then .NodeAdd "PurchHistory", tvwChild, "PurchHist1", "Permintaan Penawaran Harga", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnHSPP.Enabled = True Then .NodeAdd "PurchHistory", tvwChild, "PurchHist2", "Permintaan Barang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnHInvoiceBeli.Enabled = True Then .NodeAdd "PurchHistory", tvwChild, "HistPO", "Order Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnHReturBeli.Enabled = True Then .NodeAdd "PurchHistory", tvwChild, "PurchHist4", "Retur Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHMRP.Enabled = True Then .NodeAdd "PurchHistory", tvwChild, "PurchHist5", "Planned Order MRP", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnHInvoiceBeli.Enabled = True Then .NodeAdd "PurchHistory", tvwChild, "PurchHist6", "Invoice Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                

            Case "SALES":
                .NodeAdd , tvwChild, "SalesTrans", "Penjualan", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
                'INFO
                If MainMenu.MnInfo.Enabled = True Then .NodeAdd "SalesTrans", tvwChild, "SalesInfo", "Info", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
'                .NodeAdd "SalesInfo", tvwChild, "SalesContact", "Contact", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSalesCust.Enabled = True Then .NodeAdd "SalesInfo", tvwChild, "CustCard", "Customer", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSalesTeam.Enabled = True Then .NodeAdd "SalesInfo", tvwChild, "SalesTim", "Sales Team", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'ORDER PROSES
                If MainMenu.mnSalesAutomation.Enabled = True Then .NodeAdd "SalesTrans", tvwChild, "SalesOrderProses", "Order Proses", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnSalesFCast.Enabled = True Then .NodeAdd "SalesOrderProses", tvwChild, "SalesCast", "Forecast Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnpermintaanbarang.Enabled = True Then .NodeAdd "SalesOrderProses", tvwChild, "MaterialRequest", "Permintaan Barang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSalesQuote.Enabled = True Then .NodeAdd "SalesOrderProses", tvwChild, "SalesQuote", "Quote Order Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnContractManagement.Enabled = True Then .NodeAdd "SalesOrderProses", tvwChild, "SalesKontrak", "Kontrak Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSalesOrder.Enabled = True Then .NodeAdd "SalesOrderProses", tvwChild, "SalesOrder", "Order Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnOutstandingMKT.Enabled = True Then .NodeAdd "SalesOrderProses", tvwChild, "OutstandingMKT", "OutStanding", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                        
                
                'MARKETING
                If MainMenu.mnMarketAuto.Enabled = True Then .NodeAdd "SalesTrans", tvwChild, "Marketing", "Marketing", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMarketAutoCamp.Enabled = True Then .NodeAdd "Marketing", tvwChild, "MKTContact", "Contact", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnpermintaansample.Enabled = True Then .NodeAdd "Marketing", tvwChild, "MKTSample", "Permintaan Sample", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnmemopotonganharga.Enabled = True Then .NodeAdd "Marketing", tvwChild, "MKTMemo", "Memo Potongan Harga", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mncustomerfeedback.Enabled = True Then .NodeAdd "Marketing", tvwChild, "MKTFeedback", "Customer Feedback", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                
                'APPROVAL
                If MainMenu.mnapprovalMKT.Enabled = True Then .NodeAdd "SalesTrans", tvwChild, "MKTApproval", "Approval", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnAPPSO.Enabled = True Then .NodeAdd "MKTApproval", tvwChild, "APPSalesOrder", "Sales Order", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPPermintaanSample = True Then .NodeAdd "MKTApproval", tvwChild, "APPpermintaansample", "Permintaan Sample", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPMemo = True Then .NodeAdd "MKTApproval", tvwChild, "APPmemo", "Memo Potongan Harga", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPCustomerFeedback = True Then .NodeAdd "MKTApproval", tvwChild, "APPcustomerfeedback", "Customer Feedback", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'HISTORY
                If MainMenu.mnSalesHIstory.Enabled = True Then .NodeAdd "SalesTrans", tvwChild, "MKTHistory", "History", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHSalesQuote.Enabled = True Then .NodeAdd "MKTHistory", tvwChild, "MKTSalesQuote", "Sales Quote", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHkontrakPenjualan.Enabled = True Then .NodeAdd "MKTHistory", tvwChild, "MKTkontrakpenjualan", "Kontrak Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHOrderPenjualan.Enabled = True Then .NodeAdd "MKTHistory", tvwChild, "MKTorderpenjualan", "Order Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHInvoicePenjualan.Enabled = True Then .NodeAdd "MKTHistory", tvwChild, "MKTINVPenjualan", "Invoice Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
            Case "LOGISTIK":
                .NodeAdd , tvwChild, "LogTrans", "Logistik", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPermintaanBrg.Enabled = True Then .NodeAdd "LogTrans", tvwChild, "MaterialRequest", "Permintaan Barang", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPermintaanBeli.Enabled = True Then .NodeAdd "LogTrans", tvwChild, "OrderRequest", "Permintaan Pembelian", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                
                If MainMenu.mnxx1.Enabled = True Then .NodeAdd "LogTrans", tvwChild, "LogInbound", "Penerimaan", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBPenunjang.Enabled = True Then .NodeAdd "LogInbound", tvwChild, "LogRawSupport", "Bahan Penunjang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBJadi.Enabled = True Then .NodeAdd "LogInbound", tvwChild, "LogRFG", "Barang Jadi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnReturCust.Enabled = True Then .NodeAdd "LogInbound", tvwChild, "LogReturCust", "Retur Customer", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                If MainMenu.mnXX2.Enabled = True Then .NodeAdd "LogTrans", tvwChild, "LogOutbound", "Pengeluaran", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBKeluar.Enabled = True Then .NodeAdd "LogOutbound", tvwChild, "LogOut", "Pengeluaran Barang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnpengirimanexport.Enabled = True Then .NodeAdd "LogOutbound", tvwChild, "LogExport", "Pengiriman Export", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnReturSupp.Enabled = True Then .NodeAdd "LogOutbound", tvwChild, "LogReturSupp", "Retur Supplier", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnDO.Enabled = True Then .NodeAdd "LogOutbound", tvwChild, "LogDO", "Surat Jalan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                If mnStockOpname = True Then .NodeAdd "LogTrans", tvwChild, "StockOpname", "Stock Opname", "biru", , , False, , , True, , &HFCF1ED, &H6D4016
                If mnBInventoryBrowser = True Then .NodeAdd "LogTrans", tvwChild, "StockBrowser", "Stock Browser", "biru", , , False, , , True, , &HFCF1ED, &H6D4016
                
                'APPROVAL
                If MainMenu.mnApproval1.Enabled = True Then .NodeAdd "LogTrans", tvwChild, "LogApproval", "Approval", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnAPSPB.Enabled = True Then .NodeAdd "LogApproval", tvwChild, "SPBApproval", "Permintaan Barang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                
                'HISTORY
                .NodeAdd "LogTrans", tvwChild, "LogHistory", "History", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                      
            Case "GUDANGRL":
                'GUDANG RL
                If MainMenu.MnGudang.Enabled = True Then .NodeAdd , tvwChild, "GudangTrans", "Gudang RL", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnGudangMinta.Enabled = True Then .NodeAdd "GudangTrans", tvwChild, "MaterialRequest", "Permintaan Barang", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnTerimaRL.Enabled = True Then .NodeAdd "GudangTrans", tvwChild, "GudangTerima", "Penerimaan RL", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnKirimRL.Enabled = True Then .NodeAdd "GudangTrans", tvwChild, "GudangKirim", "Pengiriman RL", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnLembarSupplier.Enabled = True Then .NodeAdd "GudangTrans", tvwChild, "LembarSupplier", "Lembar Supplier", "biru", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnKOnfigurasiRLBatch.Enabled = True Then .NodeAdd "GudangTrans", tvwChild, "KonfigurasiRLBatch", "Konfigurasi RL Batch", "biru", , , , , , True, , &HFCF1ED, &H6D4016
            
                'approval
                .NodeAdd "GudangTrans", tvwChild, "GdgApproval", "Approval", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If mnAPPTTRL = True Then .NodeAdd "GdgApproval", tvwChild, "APPTTRL", "Tanda Terima RL", "ijo", , , False, , , True, , &HFCF1ED, &H6D4016
                If mnAPPLembSupp = True Then .NodeAdd "GdgApproval", tvwChild, "APPLEMBSUPP", "Lembar Supplier", "ijo", , , False, , , True, , &HFCF1ED, &H6D4016
                If mnAPPKRL = True Then .NodeAdd "GdgApproval", tvwChild, "APPKRL", "Pengiriman RL", "ijo", , , False, , , True, , &HFCF1ED, &H6D4016
                
            Case "PRODUKSI":
                'Setup
                .NodeAdd , tvwChild, "Asembly", "Production", "Master", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnEntupProduksi.Enabled = True Then .NodeAdd "Asembly", tvwChild, "Setup Produksi", "Setup", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                
                If MainMenu.mnRscType.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "Resources", "Resources Type", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnCalendar.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "Calendar", "Scheduling Calendar", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnManWC.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "WC", "Work Centers", "ijo", , , , , , True, , &HFCF1ED, &H6D4016 '
                If MainMenu.mnCountPoint.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "Stage", "Routing", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnRsc.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "ResourcesII", "Resources", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBomMethode.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "BOM Methode", "BOM Metode", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnJobCosting.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "INVADJ1", "Product Costing", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBomBom.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "AsemblyA3", "Bill Of Material", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnECC.Enabled = True Then .NodeAdd "Setup Produksi", tvwChild, "ECC", "Enginnering Change", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'Inventory
                If MainMenu.mnMenuPersediaan.Enabled = True Then .NodeAdd "Asembly", tvwChild, "Persediaan", "Inventory", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMInventory.Enabled = True Then .NodeAdd "Persediaan", tvwChild, "InvProduksi", "Kartu Stok", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnItemReference.Enabled = True Then .NodeAdd "Persediaan", tvwChild, "IR", "Substitusi Stok", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnItemCategories.Enabled = True Then .NodeAdd "Persediaan", tvwChild, "Categories", "Kategori", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnDescrip.Enabled = True Then .NodeAdd "Persediaan", tvwChild, "MD", "Master Outsourced ", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnTipeDes.Enabled = True Then .NodeAdd "Persediaan", tvwChild, "Descriptor", "Tipe Outsourced", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnDescripRef.Enabled = True Then .NodeAdd "Persediaan", tvwChild, "DESCREF", "Substitusi Outsourced", "ijo", , , , , , True, , &HFCF1ED, &H6D4016

   
                'Production Plan
                If MainMenu.mnMenuProduksi.Enabled = True Then .NodeAdd "Asembly", tvwChild, "PLAN", "Production Planning", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSchedule.Enabled = True Then .NodeAdd "PLAN", tvwChild, "MPS", "Master Production Schedule", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnManOrder.Enabled = True Then .NodeAdd "PLAN", tvwChild, "AsemblyA2", "Manufacturing Order", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMRPlanning.Enabled = True Then .NodeAdd "PLAN", tvwChild, "MRP", "Material Requirement Planning", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnCapaPlan.Enabled = True Then .NodeAdd "PLAN", tvwChild, "MutasiGudang2", "Capacity Planning", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnProductionPlan.Enabled = True Then .NodeAdd "PLAN", tvwChild, "PLO", "Planned Order", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'Shop Floor
                If MainMenu.mnShopFloor.Enabled = True Then .NodeAdd "Asembly", tvwChild, "Shop Floor", "Shop Floor", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMaterialRequisition.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "SFREQUEST", "Material Requisition", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMaterialIssue.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "SFISSUE", "Material Issue", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnProsesProd.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "PROD", "Proses Produksi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                'If MainMenu.mnStatusIP.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "StatusIP", "Status In Process", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPrelot.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "PRELOT", "Proses Mixing & Milling", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMutasiChip.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "MutasiChip", "Pengiriman Chip", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnLot1.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "LOT", "Blending Instruction", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBackFlushing.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "SFFLUSH", "BackFlush WIP", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSTPJadi.Enabled = True Then .NodeAdd "Shop Floor", tvwChild, "SerahTerima", "Serah Terima Produk Jadi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'APPROVAL
                .NodeAdd "Asembly", tvwChild, "ProdApproval", "Approval", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If mnAPPALKALI = True Then .NodeAdd "ProdApproval", tvwChild, "APPALKALI", "Alkali Treatment", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPACID = True Then .NodeAdd "ProdApproval", tvwChild, "APPACID", "Acid Treatment", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPBLEACHING = True Then .NodeAdd "ProdApproval", tvwChild, "APPBLEACHING", "Bleaching Treatment", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPEKSREAKTOR = True Then .NodeAdd "ProdApproval", tvwChild, "APPEKSREAKTOR", "Ekstraksi di Reaktor", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPEKSAUTO = True Then .NodeAdd "ProdApproval", tvwChild, "APPEKSAUTO", "Ekstraksi di Autoclave", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPFILTERPRESS = True Then .NodeAdd "ProdApproval", tvwChild, "APPFILTERPRESS", "Filter Press", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPGELL = True Then .NodeAdd "ProdApproval", tvwChild, "APPGELL", "Gellification", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPBUNGKUS = True Then .NodeAdd "ProdApproval", tvwChild, "APPBUNGKUS", "Pembungkusan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPCONCRETE = True Then .NodeAdd "ProdApproval", tvwChild, "APPCONCRETE", "Concrete Press", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPHYDRAULIC = True Then .NodeAdd "ProdApproval", tvwChild, "APPHYDRAULIC", "Hydraulic Press", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPCUTTER = True Then .NodeAdd "ProdApproval", tvwChild, "APPCUTTER", "Cutter", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPJEMUR = True Then .NodeAdd "ProdApproval", tvwChild, "APPJEMUR", "Penjemuran & Dryer", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPCRUSHER = True Then .NodeAdd "ProdApproval", tvwChild, "APPCRUSHER", "Crusher", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPMIXING = True Then .NodeAdd "ProdApproval", tvwChild, "APPMIXING", "Mixing & Milling", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPBLENDING = True Then .NodeAdd "ProdApproval", tvwChild, "APPBLENDING", "Blending Instruction", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPSTFG = True Then .NodeAdd "ProdApproval", tvwChild, "APPSTFG", "Serah Terima Produk Jadi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If mnAPPPAKAI = True Then .NodeAdd "ProdApproval", tvwChild, "APPPAKAI", "Pemakaian Bahan Produksi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
               'History
                If MainMenu.mnAsmHistory.Enabled = True Then .NodeAdd "Asembly", tvwChild, "AsmHistory", "History", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistAlkali.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistAlkali", "Alkali Treatment", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistAcid.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistAcid", "Acid Treatment", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistBleaching.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistBleaching", "Bleaching Treatment", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistAutoclave.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistAutoclave", "Ekstraksi di Autoclave", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistReaktor.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistReaktor", "Ekstraksi di Reaktor", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistFilter.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistFilter", "Filter Press", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistBungkus.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistBungkus", "Pembungkusan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistConcrete.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistConcrete", "Concrete Press", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistHydraulic.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistHydraulic", "Hydraulic Press", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistJemur.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistJemur", "Penjemuran", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistDryer.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistDryer", "Dryer", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnHistCrusher.Enabled = True Then .NodeAdd "AsmHistory", tvwChild, "HistCrusher", "Crusher", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                                                                                    
            Case "AKUNTING":
                .NodeAdd , tvwChild, "Akunting", "Accounting", "Master", , , True, , True, True, , &HFCF1ED, &H6D4016
                'MASTER
                If MainMenu.mnKOnfig.Enabled = True Then .NodeAdd "Akunting", tvwChild, "Master", "Setup", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnSetupAccount.Enabled = True Then .NodeAdd "Master", tvwChild, "SETUPACCOUNT", "Konfigurasi Rekening", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPerkiraan.Enabled = True Then .NodeAdd "Master", tvwChild, "MasterPerkiraan", "Daftar Perkiraan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnPeriode.Enabled = True Then .NodeAdd "Master", tvwChild, "KONFIGPERIODE", "Periode Transaksi", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mncostmethode.Enabled = True Then .NodeAdd "Master", tvwChild, "Cost Methode", "Konfigurasi Biaya", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnproductcostingAcc.Enabled = True Then .NodeAdd "Master", tvwChild, "INVADJ1", "Item Harga Pokok", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'TRANSACTION
                If MainMenu.mnKass.Enabled = True Then .NodeAdd "Akunting", tvwChild, "GROUPBKM", "Transaction", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnpermintaanbarangACC.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "PERMINTAANBARANG", "Permintaan Barang", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnTukasKas.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "PENUKARAN", "Transfer Kas", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBkm.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "BAYARTUNAILAIN", "Penerimaan Kas", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnBkk.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "TUNAIBIAYA", "Pengeluaran Kas", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.Vp.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "VoucherTransaksi", "Voucher Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.Vpe.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "VoucherTransaksi2", "Voucher Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnInvSales.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "SalesInvoice", "Invoice Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.MnInvPuchase.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "TransaksiAP", "Invoice Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'MEMORIAL
                If MainMenu.mnMemorial.Enabled = True Then .NodeAdd "Akunting", tvwChild, "MEMORIAL", "Memorial", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMemoUmum.Enabled = True Then .NodeAdd "MEMORIAL", tvwChild, "doubleentry", "Jurnal Umum", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnMemoJualbeli.Enabled = True Then .NodeAdd "MEMORIAL", tvwChild, "INVMEMO", "Penjualan / Pembelian", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                'REVALUATION
                If MainMenu.mnReval.Enabled = True Then .NodeAdd "Akunting", tvwChild, "GROUPREVAL", "Revaluation Calculation", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnRevSales.Enabled = True Then .NodeAdd "GROUPREVAL", tvwChild, "RevalSales", "Sales Series", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnRevPurchase.Enabled = True Then .NodeAdd "GROUPREVAL", tvwChild, "RevalPurchase", "Purchase Series", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
               
               'CLOSING
                If MainMenu.mnClosed.Enabled = True Then .NodeAdd "Akunting", tvwChild, "VALIDASI", "Closing", "biru", , , True, , , True, , &HFCF1ED, &H6D4016
                If MainMenu.mnClosing.Enabled = True Then .NodeAdd "VALIDASI", tvwChild, "CLOSING", "Period Closing", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
'                If MainMenu.mnhargapokokpenjualan.Enabled = True Then .NodeAdd "VALIDASI", tvwChild, "HPP", "Harga Pokok Penjualan", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
                
                '.NodeAdd "Akunting", tvwChild, "TransaksiLainLain", "Hutang/Piutang", "Bayar", , , True, , , True, , &HFCF1ED, &H6D4016
                '.NodeAdd "Akunting", tvwChild, "GROUPBKM", "Penerimaan Kas", "BKM", , , True, , , True, , &HFCF1ED, &H6D4016
                'If MainMenu.mnTukasKas.Enabled = True Then .NodeAdd "GROUPBKM", tvwChild, "PENUKARAN", "Penukaran Setara Kas", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
'                .NodeAdd "Akunting", tvwChild, "GROUPBKK", "Pengeluaran Kas", "BKK", , , True, , True, True, , &HFCF1ED, &H6D4016
                
               
            Case "KONFIGURASI":
                .NodeAdd , tvwChild, "Setup", "Setup Ledger", "KonfigReport", , , True, , , True, , &HFCF1ED, &H6D4016
                .NodeAdd "Setup", tvwChild, "SETUPACCOUNT", "Setup Report Ledger", "ijo", , , , , , True, , &HFCF1ED, &H6D4016
     End Select
     .MenuTreeView.Refresh
End With
End Sub


