VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.MDIForm MainMenu 
   BackColor       =   &H00FF0000&
   Caption         =   "Menu"
   ClientHeight    =   8715
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MainMenu.frx":6852
   StartUpPosition =   2  'CenterScreen
   Tag             =   "MAIN"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3855
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":4D60C
            Key             =   "Main"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":4E88E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":4FF78
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":51662
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":52D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":54436
            Key             =   "Expenses"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":55B20
            Key             =   "Akunting"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":55EBA
            Key             =   "BKM"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":569AC
            Key             =   "ASSETS"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5749E
            Key             =   "Data Master"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":57F90
            Key             =   "BKK"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":58A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":59574
            Key             =   "Anak Akun"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5990E
            Key             =   "Master1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5A400
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5AEF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5B9E4
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5C4D6
            Key             =   "Transaksi"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5C870
            Key             =   "Bayar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5D362
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5DE54
            Key             =   "Retur"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5E946
            Key             =   "ar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5F438
            Key             =   "ap"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":5FF2A
            Key             =   "Konfig"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":60A1C
            Key             =   "Validasi"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":61A6E
            Key             =   "KonfigReport"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":62560
            Key             =   "Memo"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":62AFA
            Key             =   "Fix Assets"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":62E94
            Key             =   "MASET"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":6322E
            Key             =   "TASET"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":635C8
            Key             =   "PRODUKSIPLAN"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":63962
            Key             =   "WHouse"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":6A1C4
            Key             =   "History"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":70A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":77288
            Key             =   "Master"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":7DAEA
            Key             =   "biru"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":8434C
            Key             =   "ijo"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":8ABAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":91410
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":97C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":9E4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":A4D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":AB598
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1005
      ButtonWidth     =   1799
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Master Data"
            Object.ToolTipText     =   "Master Data"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pembelian"
            Object.ToolTipText     =   "Distribution"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Penjualan"
            ImageIndex      =   43
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logistik"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gudang RL"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Produksi"
            Object.ToolTipText     =   "Production"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Akunting"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quality"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "HRIS"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Maintenance"
            ImageIndex      =   40
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   20
         EndProperty
      EndProperty
   End
   Begin SemeruDC.SemeruTree SemeruTree1 
      Align           =   3  'Align Left
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   13785
      BackColorTree   =   7159830
      BackColorBackground=   -2147483643
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   8385
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "User Name"
            TextSave        =   "User Name"
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Department"
            TextSave        =   "Department"
            Object.ToolTipText     =   "Department"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Server"
            TextSave        =   "Server"
            Object.ToolTipText     =   "Server Name"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5733
            MinWidth        =   5733
            Text            =   "Database"
            TextSave        =   "Database"
            Object.ToolTipText     =   "Database"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5644
            MinWidth        =   5644
            Text            =   "Menu"
            TextSave        =   "Menu"
            Object.ToolTipText     =   "Active Menu"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "Current Date"
            TextSave        =   "Current Date"
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "08:50"
            Object.ToolTipText     =   "Local Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnAPP 
      Caption         =   "Administrasi"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu app2 
         Caption         =   "-"
      End
      Begin VB.Menu mnShowMenu 
         Caption         =   "Show Menu"
      End
      Begin VB.Menu mnHideMenu 
         Caption         =   "Hide Menu"
      End
      Begin VB.Menu App3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLaporan 
         Caption         =   "Seting Laporan Baru/Tambahan"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnAdmLap 
         Caption         =   "Preview Laporan"
      End
      Begin VB.Menu mnConfReport 
         Caption         =   "Manajemen Laporan"
      End
      Begin VB.Menu mnExcel 
         Caption         =   "Import Data"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu app4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSetJournal 
         Caption         =   "Journal Setting"
         Visible         =   0   'False
      End
      Begin VB.Menu mnValidasi 
         Caption         =   "Transaction Validation"
         Visible         =   0   'False
      End
      Begin VB.Menu fft 
         Caption         =   "-"
      End
      Begin VB.Menu mnUserArea 
         Caption         =   "Otorisasi User"
      End
      Begin VB.Menu aaw 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnMaster 
      Caption         =   "Master Data"
      Begin VB.Menu mnCurrency 
         Caption         =   "Multi Currency"
         Begin VB.Menu mnExcCurrency 
            Caption         =   "Currency Setup"
         End
         Begin VB.Menu mnExAccess 
            Caption         =   "Exchange Rate Access"
            Visible         =   0   'False
         End
         Begin VB.Menu mnExAccount 
            Caption         =   "Posting Account Setup"
            Visible         =   0   'False
         End
         Begin VB.Menu mnExMaint 
            Caption         =   "Exchange Maintenance"
         End
      End
      Begin VB.Menu MnGudangLog 
         Caption         =   "WareHouse"
      End
      Begin VB.Menu MnProdCosting 
         Caption         =   "Product Costing"
         Visible         =   0   'False
      End
      Begin VB.Menu mnItem 
         Caption         =   "Inventory"
         Begin VB.Menu mnInvCard 
            Caption         =   "Inventory Card"
         End
         Begin VB.Menu mnInvManager 
            Caption         =   "Inventory Manager"
         End
         Begin VB.Menu mnInvPurch 
            Caption         =   "Inventory Purchasing"
         End
         Begin VB.Menu mnKelompok 
            Caption         =   "Inventory Class"
         End
         Begin VB.Menu mnItemCategories1 
            Caption         =   "Inventory Categories"
         End
         Begin VB.Menu mntemReference1 
            Caption         =   "Inventory Substitutions"
         End
      End
      Begin VB.Menu mnRegional 
         Caption         =   "Regional"
      End
      Begin VB.Menu mnTipeBayar 
         Caption         =   "Payment"
         Visible         =   0   'False
      End
      Begin VB.Menu mnTransport 
         Caption         =   "Transporter"
      End
      Begin VB.Menu mnKaryawan 
         Caption         =   "Employee"
         Visible         =   0   'False
      End
      Begin VB.Menu mncus 
         Caption         =   "Customer"
         Begin VB.Menu mnCusMaster 
            Caption         =   "Customer Card"
         End
         Begin VB.Menu mnCusGudang 
            Caption         =   "Customer Warehouse"
         End
      End
      Begin VB.Menu mnSup 
         Caption         =   "Supplier Card"
      End
      Begin VB.Menu mnBankPartner 
         Caption         =   "Bank Partner"
      End
      Begin VB.Menu mnTransID 
         Caption         =   "Setup Nomor Transaksi"
      End
      Begin VB.Menu mnterminpembayaran 
         Caption         =   "Termin Pembayaran"
      End
      Begin VB.Menu mnTipeItem 
         Caption         =   "Tipe Item Transaksi"
      End
      Begin VB.Menu mnItemCharge 
         Caption         =   "Item Charge"
      End
      Begin VB.Menu mnUOM 
         Caption         =   "Unit Of Measurement (Satuan Barang)"
      End
      Begin VB.Menu mnuProduksi 
         Caption         =   "Produksi"
         Begin VB.Menu mnuMasterAnalisa 
            Caption         =   "Master Analisa"
         End
         Begin VB.Menu mnuMasterProses 
            Caption         =   "Master Proses"
         End
      End
   End
   Begin VB.Menu mnPurchase 
      Caption         =   "Pembelian"
      Begin VB.Menu mnorderproses 
         Caption         =   "Order Proses"
         Begin VB.Menu MnRequest 
            Caption         =   "Permintaan Pembelian"
         End
         Begin VB.Menu MnPriceOffer 
            Caption         =   "Permintaan Penawaran Harga"
         End
         Begin VB.Menu mnPenawaranHrg 
            Caption         =   "Penawaran Harga"
         End
         Begin VB.Menu MnGaris1 
            Caption         =   "-"
         End
         Begin VB.Menu mnPurchaseOrder 
            Caption         =   "Order Pembelian"
         End
         Begin VB.Menu mnPOBlanked 
            Caption         =   "Blanked PO"
         End
         Begin VB.Menu mnPOBlanked1 
            Caption         =   "Blanked PO Patty Cash"
         End
         Begin VB.Menu mnGaris2 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnPurchaseReturn 
            Caption         =   "Retur Pembelian"
         End
         Begin VB.Menu MnPlanOrder 
            Caption         =   "Planned Order MRP"
         End
         Begin VB.Menu mnMRP 
            Caption         =   "MRP Generation"
         End
      End
      Begin VB.Menu mnGaris 
         Caption         =   "-"
      End
      Begin VB.Menu mneksekusi 
         Caption         =   "Eksekusi"
         Begin VB.Menu OutSPH 
            Caption         =   "OutstandingSPH"
         End
         Begin VB.Menu mnoutstand 
            Caption         =   "Outstanding Pembelian"
         End
         Begin VB.Menu MnPPN 
            Caption         =   "Billing PPN"
            Visible         =   0   'False
         End
         Begin VB.Menu MnEvaluasi 
            Caption         =   "Evaluasi Supplier"
         End
      End
      Begin VB.Menu mnGaris4 
         Caption         =   "-"
      End
      Begin VB.Menu mnapproval 
         Caption         =   "Approval"
         Begin VB.Menu mnAppSPP 
            Caption         =   "Permintaan Pembelian"
         End
      End
      Begin VB.Menu mnGaris3 
         Caption         =   "-"
      End
      Begin VB.Menu MnHistoryPurchase 
         Caption         =   "History"
         Begin VB.Menu MnHSPPH 
            Caption         =   "Permintaan Penawaran Harga"
         End
         Begin VB.Menu MnHSPP 
            Caption         =   "Permintaan Barang"
         End
         Begin VB.Menu mnHPO 
            Caption         =   "Order Pembelian"
         End
         Begin VB.Menu MnHReturBeli 
            Caption         =   "Retur Pembelian"
         End
         Begin VB.Menu mnHMRP 
            Caption         =   "Planned Order MRP"
         End
         Begin VB.Menu MnHInvoiceBeli 
            Caption         =   "Invoice Pembelian"
         End
         Begin VB.Menu MnhPPN 
            Caption         =   "Billing PPN"
         End
      End
   End
   Begin VB.Menu mnMarketing 
      Caption         =   "Penjualan && Marketing "
      Begin VB.Menu MnInfo 
         Caption         =   "Info"
         Begin VB.Menu mnSalesContact 
            Caption         =   "Contact"
            Visible         =   0   'False
         End
         Begin VB.Menu mnSalesCust 
            Caption         =   "Customer"
         End
         Begin VB.Menu mnSalesTeam 
            Caption         =   "Sales Team"
         End
      End
      Begin VB.Menu mnSalesAutomation 
         Caption         =   "Order Proses"
         Begin VB.Menu MnSalesFCast 
            Caption         =   "Sales Forecast"
         End
         Begin VB.Menu mnMemoPotHrg 
            Caption         =   "Memo Potongan Harga"
         End
         Begin VB.Menu mnpermintaanbarang 
            Caption         =   "Permintaan Barang"
         End
         Begin VB.Menu mnSalesQuote 
            Caption         =   "Sales Quote"
         End
         Begin VB.Menu mnContractManagement 
            Caption         =   "Kontrak Penjualan"
         End
         Begin VB.Menu mnSalesOrder 
            Caption         =   "Order Penjualan"
         End
         Begin VB.Menu mnOutstandingMKT 
            Caption         =   "Outstanding"
         End
         Begin VB.Menu mnSalesReturn 
            Caption         =   "Return Penjualan"
         End
      End
      Begin VB.Menu mnMarketAuto 
         Caption         =   "Marketing"
         Begin VB.Menu mnMarketAutoCamp 
            Caption         =   "Contact"
         End
         Begin VB.Menu mnMarketCampaign 
            Caption         =   "Campaign"
         End
         Begin VB.Menu mnpermintaansampleSales 
            Caption         =   "Permintaan Sample"
         End
         Begin VB.Menu mnmemopotonganharga 
            Caption         =   "Memo Potongan Harga"
         End
         Begin VB.Menu mncustomerfeedback 
            Caption         =   "Customer Feedback"
         End
      End
      Begin VB.Menu mnapprovalMKT 
         Caption         =   "Approval"
         Begin VB.Menu mnAPPSO 
            Caption         =   "Approval Sales Order"
         End
      End
      Begin VB.Menu mnsalesReport 
         Caption         =   "Report"
         Visible         =   0   'False
         Begin VB.Menu mnpermintaansample 
            Caption         =   "Permintaan Sample"
         End
         Begin VB.Menu mnmemopotongan 
            Caption         =   "Memo Potongan Harga"
         End
         Begin VB.Menu mnCustFeedback 
            Caption         =   "Customer Feedback"
         End
         Begin VB.Menu mnReportSalesOrder 
            Caption         =   "Sales Order"
         End
         Begin VB.Menu mnReportSalesQuote 
            Caption         =   "Sales Quote"
         End
      End
      Begin VB.Menu mnSalesHIstory 
         Caption         =   "History"
         Begin VB.Menu mnHSalesQuote 
            Caption         =   "History Sales Quote"
         End
         Begin VB.Menu mnHkontrakPenjualan 
            Caption         =   "History Kontrak Penjualan"
         End
         Begin VB.Menu mnHOrderPenjualan 
            Caption         =   "History Order Penjualan"
         End
         Begin VB.Menu mnHInvoicePenjualan 
            Caption         =   "History Invoice Penjualan"
         End
      End
   End
   Begin VB.Menu MnGudang 
      Caption         =   "Gudang RL"
      Begin VB.Menu MnGudangMinta 
         Caption         =   "Permintaan Barang"
      End
      Begin VB.Menu MnTerimaRL 
         Caption         =   "Penerimaan Rumput Laut"
      End
      Begin VB.Menu MnKirimRL 
         Caption         =   "Pengiriman RL"
      End
      Begin VB.Menu mnLembarSupplier 
         Caption         =   "Lembar Supplier"
      End
      Begin VB.Menu mnKOnfigurasiRLBatch 
         Caption         =   "Konfigurasi RL Batch"
      End
   End
   Begin VB.Menu mnLogistik 
      Caption         =   "Logistik"
      Begin VB.Menu mnPermintaanBrg 
         Caption         =   "Permintaan Barang"
      End
      Begin VB.Menu mnPermintaanBeli 
         Caption         =   "Permintaan Pembelian"
      End
      Begin VB.Menu mnxx3 
         Caption         =   "-"
      End
      Begin VB.Menu mnxx1 
         Caption         =   "Penerimaan"
         Begin VB.Menu mnBPenunjang 
            Caption         =   "Bahan Penunjang"
         End
         Begin VB.Menu mnBJadi 
            Caption         =   "Barang Jadi"
         End
         Begin VB.Menu mnReturCust 
            Caption         =   "Retur Customer"
         End
      End
      Begin VB.Menu MnTrans1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnXX2 
         Caption         =   "Pengeluaran"
         Begin VB.Menu mnBKeluar 
            Caption         =   "Pengeluaran Barang"
         End
         Begin VB.Menu mnpengirimanexport 
            Caption         =   "Pengiriman Export"
         End
         Begin VB.Menu mnReturSupp 
            Caption         =   "Retur Supplier"
         End
         Begin VB.Menu mnDO 
            Caption         =   "Surat Jalan"
         End
      End
      Begin VB.Menu mnGAris5 
         Caption         =   "-"
      End
      Begin VB.Menu mnApproval1 
         Caption         =   "Approval"
         Begin VB.Menu mnAPSPB 
            Caption         =   "Permintaan Barang"
         End
      End
      Begin VB.Menu mnxx4 
         Caption         =   "-"
      End
      Begin VB.Menu mnHistoryLogistik 
         Caption         =   "History"
      End
   End
   Begin VB.Menu mnInventory 
      Caption         =   "Produksi"
      Begin VB.Menu mnEntupProduksi 
         Caption         =   "Setup"
         Begin VB.Menu mnTypeCost 
            Caption         =   "Cost Methode"
            Visible         =   0   'False
         End
         Begin VB.Menu mnRscType 
            Caption         =   "Resources Type"
         End
         Begin VB.Menu mnCalendar 
            Caption         =   "Scheduling Calendar"
         End
         Begin VB.Menu mnManWC 
            Caption         =   "Work Center"
         End
         Begin VB.Menu mnLot 
            Caption         =   "Lot Sizing"
            Visible         =   0   'False
         End
         Begin VB.Menu mnCountPoint 
            Caption         =   "Routing"
         End
         Begin VB.Menu mnRsc 
            Caption         =   "Resources"
         End
         Begin VB.Menu mnBomMethode 
            Caption         =   "BOM Methode"
         End
         Begin VB.Menu mnJobCosting 
            Caption         =   "Product Costing"
         End
         Begin VB.Menu mnBomBom 
            Caption         =   "Bill Of Material"
         End
         Begin VB.Menu mnECC 
            Caption         =   "Enginering Change"
         End
      End
      Begin VB.Menu mnuKonfigurasi 
         Caption         =   "Konfigurasi"
         Begin VB.Menu mnuKonfigurasiProsedur 
            Caption         =   "Konfigurasi Prosedur"
         End
         Begin VB.Menu mnuKonfigurasiSampleForm 
            Caption         =   "Konfigurasi Form Produksi"
         End
      End
      Begin VB.Menu mnMenuPersediaan 
         Caption         =   "Inventory"
         Begin VB.Menu mnMInventory 
            Caption         =   "Inventory Card"
         End
         Begin VB.Menu mnItemReference 
            Caption         =   "Inventory Reference"
         End
         Begin VB.Menu mnItemCategories 
            Caption         =   "Inventory Categories"
         End
         Begin VB.Menu mnDescrip 
            Caption         =   "Master Outsourced"
         End
         Begin VB.Menu mnTipeDes 
            Caption         =   "Outsourced Type"
         End
         Begin VB.Menu mnDescripRef 
            Caption         =   "Outsourced Referense"
         End
      End
      Begin VB.Menu mnMenuProduksi 
         Caption         =   "Production Planning"
         Begin VB.Menu mnManOrder 
            Caption         =   "Manufacturing Order"
         End
         Begin VB.Menu mnSchedule 
            Caption         =   "Master Production Schedule"
         End
         Begin VB.Menu mnMRPlanning 
            Caption         =   "Master Requirement Planning"
         End
         Begin VB.Menu mnCapaPlan 
            Caption         =   "Capacity Planning"
         End
         Begin VB.Menu mnProductionPlan 
            Caption         =   "Planned Order"
         End
      End
      Begin VB.Menu mnShopFloor 
         Caption         =   "Shop Floor"
         Begin VB.Menu mnMaterialRequisition 
            Caption         =   "Material Requisition"
         End
         Begin VB.Menu mnMaterialIssue 
            Caption         =   "Material Issue"
         End
         Begin VB.Menu mnProsesProd 
            Caption         =   "Proses Produksi"
         End
         Begin VB.Menu mnPrelot 
            Caption         =   "Proses Mixing & Milling"
         End
         Begin VB.Menu mnMutasiChip 
            Caption         =   "Pengiriman Chip"
         End
         Begin VB.Menu mnLot1 
            Caption         =   "Blanding Production"
         End
         Begin VB.Menu mnBackFlushing 
            Caption         =   "Backflush WIP"
         End
         Begin VB.Menu mnStatusIP 
            Caption         =   "Status In Proses"
         End
         Begin VB.Menu mnSTPJadi 
            Caption         =   "Serah Terima Produk Jadi"
         End
      End
      Begin VB.Menu mnAsmHistory 
         Caption         =   "History"
         Begin VB.Menu mnHistAlkali 
            Caption         =   "Alkali Treatment"
         End
         Begin VB.Menu mnHistAcid 
            Caption         =   "Acid Treatment"
         End
         Begin VB.Menu mnHistBleaching 
            Caption         =   "Bleaching Teratment"
         End
         Begin VB.Menu mnHistAutoclave 
            Caption         =   "Ekstraksi di Autoclave"
         End
         Begin VB.Menu mnHistReaktor 
            Caption         =   "Ekstraksi di Reaktor"
         End
         Begin VB.Menu mnHistFilter 
            Caption         =   "Filter Press"
         End
         Begin VB.Menu mnHistBungkus 
            Caption         =   "Pembungkusan"
         End
         Begin VB.Menu mnHistConcrete 
            Caption         =   "Concrete Press"
         End
         Begin VB.Menu mnHistHydraulic 
            Caption         =   "Hydraulic Press"
         End
         Begin VB.Menu mnHistJemur 
            Caption         =   "Penjemuran"
         End
         Begin VB.Menu mnHistDryer 
            Caption         =   "Dryer"
         End
         Begin VB.Menu mnHistCrusher 
            Caption         =   "Crusher"
         End
      End
   End
   Begin VB.Menu mnAkun 
      Caption         =   "Accounting"
      Begin VB.Menu mnMasAkun 
         Caption         =   "Data Master"
         Begin VB.Menu mnKOnfig 
            Caption         =   "Setup"
            Begin VB.Menu mnPerkiraan 
               Caption         =   "Daftar Perkiraan"
            End
            Begin VB.Menu mnPeriode 
               Caption         =   "Periode Transaksi"
            End
            Begin VB.Menu mnSetupAccount 
               Caption         =   "Konfigurasi Rekening"
            End
            Begin VB.Menu mncostmethode 
               Caption         =   "Konfigurasi Biaya"
            End
            Begin VB.Menu mnproductcostingAcc 
               Caption         =   "Item Harga Pokok"
            End
         End
      End
      Begin VB.Menu mnKass 
         Caption         =   "Transaksi"
         Begin VB.Menu mnpermintaanbarangACC 
            Caption         =   "Permintaan Barang"
         End
         Begin VB.Menu mnBkmPiutang 
            Caption         =   "Pelunasan Piutang Karyawan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnBkm 
            Caption         =   "Cash Receipt"
         End
         Begin VB.Menu mnTukasKas 
            Caption         =   "Penukaran Setara Kas Ke Kas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnBkk 
            Caption         =   "Cash Payment"
         End
         Begin VB.Menu Vp 
            Caption         =   "Voucher Pembelian"
         End
         Begin VB.Menu Vpe 
            Caption         =   "Voucher Penjualan"
         End
         Begin VB.Menu mnVoucher 
            Caption         =   "Payables / Receivable"
            Visible         =   0   'False
         End
         Begin VB.Menu mnBkkPiutang 
            Caption         =   "Pengeluaran Piutang Ke Karyawan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnKas 
            Caption         =   "Pengeluaran Kas Harian"
            Visible         =   0   'False
         End
         Begin VB.Menu MnInvSales 
            Caption         =   "Invoice Penjualan"
         End
         Begin VB.Menu MnInvPuchase 
            Caption         =   "Invoice Pembelian"
         End
      End
      Begin VB.Menu mnReval 
         Caption         =   "Revaluation"
         Begin VB.Menu mnRevFinance 
            Caption         =   "Financial Series"
            Visible         =   0   'False
         End
         Begin VB.Menu mnRevSales 
            Caption         =   "Sales Series"
         End
         Begin VB.Menu mnRevPurchase 
            Caption         =   "Purchase Series"
         End
      End
      Begin VB.Menu mnBKas 
         Caption         =   "Pengeluaran Kas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnHutangPiutang 
         Caption         =   "Pelunasan Hutang / Piutang"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnMemorial 
         Caption         =   "Memorial"
         Begin VB.Menu mnMemoUmum 
            Caption         =   "Memorial Jurnal"
         End
         Begin VB.Menu mnMemoJualbeli 
            Caption         =   "Sales / Purchase"
         End
      End
      Begin VB.Menu mnClosed 
         Caption         =   "Period Closing"
         Begin VB.Menu mnClosing 
            Caption         =   "Periode Closing"
         End
      End
   End
   Begin VB.Menu mnQuality 
      Caption         =   "Quality"
   End
   Begin VB.Menu mnMaintenance 
      Caption         =   "Maintenance"
   End
   Begin VB.Menu mnHrd 
      Caption         =   "HRD"
   End
   Begin VB.Menu mnBantu 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnAtur 
         Caption         =   "Atur Window Cascade"
      End
      Begin VB.Menu mnHoris 
         Caption         =   "Atur Window Horizontal"
      End
      Begin VB.Menu mnVertical 
         Caption         =   "Atur Window Vertical"
      End
      Begin VB.Menu mnBantuan0 
         Caption         =   "-"
      End
      Begin VB.Menu mnTTp 
         Caption         =   "Tutup Semua Form"
      End
   End
   Begin VB.Menu mnHelpApp 
      Caption         =   "Help"
      Begin VB.Menu mnHlpApp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnLisence 
         Caption         =   "Lisence Agreement"
      End
      Begin VB.Menu mnUpdatePatch 
         Caption         =   "Update Patch "
      End
   End
   Begin VB.Menu MnNodes 
      Caption         =   "Nodes"
      Visible         =   0   'False
      Begin VB.Menu mnTdep 
         Caption         =   "Tambah Departement"
      End
      Begin VB.Menu Nodea 
         Caption         =   "-"
      End
      Begin VB.Menu mnJabat 
         Caption         =   "Tambah Jabatan"
      End
      Begin VB.Menu Nodeb 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdit 
         Caption         =   "Edit Struktur Organisasi"
      End
      Begin VB.Menu mnHapus 
         Caption         =   "Hapus Struktur Organisasi"
      End
   End
   Begin VB.Menu mnNodePolicy 
      Caption         =   "NodePolicy"
      Visible         =   0   'False
      Begin VB.Menu mnTambahGroupMenu 
         Caption         =   "Tambah Group Menu"
      End
      Begin VB.Menu mnNdPlc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEditGroupMenu 
         Caption         =   "Edit Group Menu"
      End
      Begin VB.Menu mnDeleteGroupMenu 
         Caption         =   "Delete Group Menu"
      End
      Begin VB.Menu mnNdPlc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnTranfer 
         Caption         =   "Tranfer List Form"
      End
      Begin VB.Menu mnSetingOtorisasi 
         Caption         =   "Seting Otorisasi User"
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim myMenu As New clsMenu
Dim mCalV, mCalS As Boolean
Dim mbMoving As Boolean
Private IdleStart As Long
Private ACC As New clsApproval

Private Sub IdleMon1_IdleStateDisengaged(ByVal IdleStopTime As Long)
   IdleStart = IdleStopTime
   'StartFromIdle = False
End Sub

Private Sub IdleMon1_IdleStateEngaged(ByVal IdleStartTime As Long)
   'MainMenu.Caption = IdleStartTime - IdleStart
   If IdleStartTime - IdleStart > Val(IDLELIMIT) Then
      StartFromIdle = True
      frmLogin.Show
      MainMenu.Enabled = False
   End If
End Sub

Private Sub MDIForm_Activate()
'myMenu.CreateMenu "MASTER"
MainMenu.StatusBar1.Panels(6).Text = Format(Date, "dd MMMM yyyy")
MainMenu.StatusBar1.Panels(2).Text = "Departemen : " & NamaDept
'Toolbar1.Buttons(10).Visible = False
Toolbar1.Buttons(4).Visible = True
Toolbar1.Buttons(1).Caption = "Master Data"

End Sub

Private Sub MDIForm_Load()
   On Error Resume Next
   'OpenMenu
   Me.Caption = App.Comments
   'myMenu.CreateMenu "MASTER"
   SemeruTree1.Visible = False
   Err.Clear
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim I As Integer

If StartFromIdle = True Then
    Cancel = True
    MainMenu.Enabled = False
    frmLogin.Show
Else
   I = MessageBox("Anda Yakin Untuk Keluar Aplikasi?", "Keluar Aplikasi", msgYesNo, msgQuestion)
   If I = 1 Then
      If Not CNN Is Nothing Then
         If CNN.State = 1 Then
            CNN.Close
         End If
      End If
      Set CNN = Nothing
      Cancel = False
   Else
      Cancel = True
   End If
End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
   Set MainMenu = Nothing
   Unload frmFindHistory
End Sub

Private Sub mnAdmLap_Click()
If frmReport.Enabled = True Then frmReport.SetFocus
End Sub



Private Sub mnAPPSO_Click()
'Approval Sales Order
   ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
   ACC.Key = "PurchaseID"
   ACC.StrSQLMaster = "Select approved_by as [Approved By],purchaseID, datepurchase as Tanggal,person as [User],Date_approved as [Waktu Approved] from [PO Order] where ((status=0 or status=2) and typetrans='SO')"
   ACC.StrSQLDetail = "select * from QuerySO"
   ACC.MasterTable = "[PO Order]"
   frmApproval.Validasi = ACC
   frmApproval.Caption = "Validasi Sales Order"
   frmApproval.SetFocus
End Sub

Private Sub mnAppSPP_Click()
   ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
   ACC.Key = "SPPID"
   ACC.StrSQLMaster = "Select approved_by as [Approved By],SPPID, SPP_date as Tanggal,Ordered_BY as [User],Date_approved as [Waktu Approved] from SPP_header where status=0"
   ACC.StrSQLDetail = "Select NoItem as Kode, itemName as [Nama Barang], UOM as Satuan,Qty_spp as Jumlah,Keperluan,note as Keterangan from QuerySPP"
   ACC.MasterTable = "SPP_header"
   frmApproval.Validasi = ACC
   frmApproval.Caption = "Validasi Permintaan Pembelian"
   frmApproval.SetFocus
End Sub

Private Sub mnAPSPB_Click()
   ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_Approved"
   ACC.Key = "IDTrans"
   
   Select Case CurrentDept
      Case "PR"
         ACC.StrSQLMaster = "select Approved_By as [Approved By],IDTrans,DateTrans,Note,[Issued By],[Received By],dept as Departemen, Date_approved as [Waktu Approved] from [backflush_Header] where status=0 and typeTrans <> 'MI' and (dept ='PR' or dept = 'MT') order by [issued by],DateTrans"
         'ACC.StrSQLMaster = "select Approved_By as [Approved By],IDTrans,DateTrans,Note,[Issued By],[Received By],dept as Departemen, Date_approved as [Waktu Approved] from [backflush_Header] where status=0 and typeTrans <> 'MI' and dept ='PR'  order by [issued by],DateTrans"
      Case "MG"
         ACC.StrSQLMaster = "select Approved_By as [Approved By],IDTrans,DateTrans,Note,[Issued By],[Received By],dept as Departemen, Date_approved as [Waktu Approved] from [backflush_Header] where status=0 and typeTrans <> 'MI' and dept <>'PR' order by [issued by],DateTrans"
      Case "MT"
         ACC.StrSQLMaster = "select Approved_By as [Approved By],IDTrans,DateTrans,Note,[Issued By],[Received By],dept as Departemen, Date_approved as [Waktu Approved] from [backflush_Header] where status=0 and typeTrans <> 'MI' and (dept = 'MT' or dept='PR') order by [issued by],DateTrans"
   End Select
     
   ACC.StrSQLDetail = "select * from querySPB"
   ACC.MasterTable = "backflush_header"
   frmApproval.Validasi = ACC
   frmApproval.Caption = "Validasi Permintaan Barang"
   frmApproval.SetFocus
End Sub

Private Sub mnAtur_Click()
MainMenu.Arrange vbCascade
End Sub

Private Sub mnBankPartner_Click()
If frmBankPartner.Enabled = True Then frmBankPartner.SetFocus
End Sub


Private Sub mnBJadi_Click()
   FrmSFComplete.SetFocus
End Sub

Private Sub mnBKeluar_Click()
   FrmSJWarehouse.SetFocus
End Sub

Private Sub mnBkk_Click()
If FrmBKK.Enabled = True Then FrmBKK.SetFocus
End Sub

Private Sub mnBkkPiutang_Click()
'If FrmPiutangKaryawan.Enabled = True Then FrmPiutangKaryawan.SetFocus
End Sub

Private Sub mnBkm_Click()
If FrmBKM.Enabled = True Then FrmBKM.SetFocus
End Sub

Private Sub mnBkmPiutang_Click()
'If frmPembayaranPKaryawan.Enabled = True Then frmPembayaranPKaryawan.SetFocus
End Sub

Private Sub mnBomBom_Click()
If frmbom.Enabled = True Then frmbom.SetFocus
End Sub

Private Sub mnBomMethode_Click()
If FrmBOMMethode.Enabled = True Then FrmBOMMethode.SetFocus
End Sub

Private Sub mnBPenunjang_Click()
   frmReceiveNotes.SetFocus
End Sub

Private Sub mnCalendar_Click()
If FrmCalendar.Enabled = True Then FrmCalendar.SetFocus
End Sub


Private Sub mnClosing_Click()
If frmValidasi.Enabled = True Then frmValidasi.SetFocus
End Sub

Private Sub mnConfReport_Click()
   If frmReportConfig.Enabled Then frmReportConfig.SetFocus
End Sub

Private Sub mnContractManagement_Click()
   frmSalesContract.SetFocus
End Sub

Private Sub mnCountPoint_Click()
If FrmRouting.Enabled = True Then FrmRouting.SetFocus
End Sub

Private Sub mnCusMaster_Click()
If frmPartner.Enabled = True Then frmPartner.SetFocus
End Sub

Private Sub mnCustFeedback_Click()
   FrmCustFeedBack.SetFocus
End Sub

Private Sub mncustomerfeedback_Click()
   FrmCustFeedBack.SetFocus
End Sub

Private Sub mnDeleteGroupMenu_Click()
'FrmPolicy.AddNode "delete"
End Sub

Private Sub mnDescrip_Click()
If FrmManDescriptor.Enabled = True Then FrmManDescriptor.SetFocus
End Sub

Private Sub mnDescripRef_Click()
If FrmItemDescriptor.Enabled = True Then FrmItemDescriptor.SetFocus
End Sub

Private Sub mnDn_Click()
If FrmDO.Enabled = True Then FrmDO.SetFocus
End Sub

Private Sub mnDO_Click()
   FrmDO.SetFocus
End Sub

Private Sub mnECC_Click()
If FrmEnginering.Enabled = True Then FrmEnginering.SetFocus
End Sub

Private Sub mnEdit_Click()
Call frmEmployess.mnEdit_Click
End Sub

Private Sub mnEditGroupMenu_Click()
'FrmPolicy.AddNode "EDIT"
End Sub

Private Sub MnEvaluasi_Click()
   FrmEvaluasiSupplier.SetFocus
End Sub

Private Sub mnExAccount_Click()
'If FrmCurrencyAccount.Enabled = True Then FrmCurrencyAccount.SetFocus
End Sub

Private Sub mnExcCurrency_Click()
   If FrmCurrencySetup.Enabled = True Then FrmCurrencySetup.SetFocus
End Sub

Private Sub mnExcel_Click()
If frmImport.Enabled = True Then frmImport.SetFocus
End Sub

Private Sub mnExit_Click()
End
End Sub

Private Sub mnExMaint_Click()
   If FrmCurrencyMaint.Enabled = True Then FrmCurrencyMaint.SetFocus
End Sub

Private Sub mnCusGudang_Click()
If FrmGudangCust.Enabled = True Then FrmGudangCust.SetFocus
End Sub

Private Sub MnGudangLog_Click()
If frmWareHouse.Enabled = True Then frmWareHouse.SetFocus
End Sub

Private Sub MnGudangMinta_Click()
   FrmMatRequest.SetFocus
End Sub

Private Sub mnHapus_Click()
Call frmEmployess.mnHapus_Click
End Sub

Private Sub mnHideMenu_Click()
SemeruTree1.Visible = False
End Sub

Private Sub mnHlpApp_Click()
   ShellExecute 0, "open", "hh.exe", App.Path + "\COSMIC ERP.chm", "", 1
End Sub

Private Sub mnHoris_Click()
MainMenu.Arrange vbTileHorizontal
End Sub



Private Sub mnInvAdj_Click()
If FrmInvAdj.Enabled = True Then FrmInvAdj.SetFocus
End Sub

Private Sub mnInvCard_Click()
If FrmItemData.Enabled = True Then FrmItemData.SetFocus
End Sub

Private Sub mnInvManager_Click()
   FrmItemData.SetMode = "Manager"
   FrmItemData.SetFocus
End Sub

Private Sub MnInvPuchase_Click()
   FrmInvoice.SetFocus
End Sub

Private Sub mnInvPurch_Click()
   FrmItemData.SetMode = "Purchasing"
   FrmItemData.SetFocus
End Sub

Private Sub MnInvSales_Click()
If frmArTrans.Enabled = True Then frmArTrans.SetFocus
End Sub

Private Sub mnItemCategories_Click()
If FrmCategories.Enabled = True Then FrmCategories.SetFocus
End Sub

Private Sub mnItemCategories1_Click()
FrmCategories.SetFocus
End Sub

Private Sub mnItemCharge_Click()
   frmItemCharge.SetFocus
End Sub

Private Sub mnItemReference_Click()
If FrmItemReference.Enabled = True Then FrmItemReference.SetFocus
End Sub

Private Sub mnJabat_Click()
Call frmEmployess.mnJabat_Click
End Sub

Private Sub mnJobCosting_Click()
If frmBOMCosting.Enabled = True Then frmBOMCosting.SetFocus
End Sub


Private Sub mnKaryawan_Click()
frmEmployess.SetFocus
End Sub

Private Sub mnKelompok_Click()
If frmKelompok.Enabled = True Then frmKelompok.SetFocus
End Sub

Private Sub MnKirimRL_Click()
   FrmSJWarehouse.SetFocus
End Sub

'Private Sub mnLain_Click()
''frmItemPrice.SetFocus
'End Sub

Private Sub mnLaporan_Click()
If FrmKonfigurasiAccount.Enabled = True Then FrmKonfigurasiAccount.SetFocus
End Sub

Private Sub mnLembarSupplier_Click()
   frmLembarSupplier.SetFocus
End Sub

Private Sub mnLisence_Click()
If frmAbout.Enabled = True Then frmAbout.SetFocus

End Sub

Private Sub mnLogin_Click()
    IsLogOff = True
    SemeruTree1.Visible = False
    CloseAllForm
    frmLogin.Show vbModal
End Sub

Private Sub mnLot1_Click()
   frmBlending.SetFocus
End Sub

Private Sub mnMaintenance_Click()
  Shell App.Path + "\MMT.EXE", vbNormalFocus
End Sub

Private Sub mnManOrder_Click()
If FrmMOrder.Enabled = True Then FrmMOrder.SetFocus
End Sub

'Private Sub mnManStage_Click()
'If frmRouting.Enabled = True Then frmRouting.SetFocus
'End Sub

Private Sub mnManWC_Click()
If FrmWorkCenter.Enabled = True Then FrmWorkCenter.SetFocus
End Sub

Private Sub mnMarketAutoCamp_Click()
   frmcontact.SetFocus
End Sub

Private Sub mnMaterialIssue_Click()
   FrmSFIssue.SetFocus
End Sub

Private Sub mnMaterialRequisition_Click()
   frmSFRequest.SetFocus
End Sub

Private Sub mnMemoJualbeli_Click()
If frmInvMemo.Enabled = True Then frmInvMemo.SetFocus
End Sub

Private Sub mnMemoPotHrg_Click()
   frmMemoPotongHarga.SetFocus
End Sub

Private Sub mnMemoPotongan_Click()
'frmMemoPotongHarga.SetFocus
End Sub

Private Sub mnmemopotonganharga_Click()
   frmMemoPotongHarga.SetFocus
End Sub

Private Sub mnMemoUmum_Click()
If frmMemorial.Enabled = True Then frmMemorial.SetFocus
End Sub

Private Sub mnMInventory_Click()
If FrmItemData.Enabled = True Then FrmItemData.SetFocus
End Sub

Private Sub mnMrp_Click()
If FrmMRP.Enabled = True Then FrmMRP.SetFocus
End Sub

Private Sub mnMutasi_Click()
If frmMutasiGudang.Enabled = True Then frmMutasiGudang.SetFocus
End Sub

Private Sub mnMRPlanning_Click()
   FrmMRPLogic.SetFocus
End Sub

Private Sub mnMutasiChip_Click()
   frmMutasiChip.SetFocus
End Sub

Private Sub MnOutstand_Click()
   frmOutstandingPO.SetFocus
End Sub

Private Sub mnPenawaranHrg_Click()
   frmSPH.SetFocus
End Sub

Private Sub mnPeriode_Click()
If FrmSetingPeriode.Enabled = True Then FrmSetingPeriode.SetFocus
End Sub

Private Sub mnPerkiraan_Click()
If FrmPerkiraan.Enabled = True Then FrmPerkiraan.SetFocus
End Sub

Private Sub mnpermintaanbarang_Click()
   FrmMatRequest.SetFocus
End Sub

Private Sub mnpermintaanbarangACC_Click()

   FrmMatRequest.SetFocus
End Sub

Private Sub mnPermintaanBeli_Click()
   FrmPRequest.SetFocus
End Sub

Private Sub mnPermintaanBrg_Click()
   FrmMatRequest.SetFocus
End Sub

Private Sub mnPermintaanSample_Click()
   FrmSampleOrder.Show
End Sub

Private Sub mnpermintaansampleSales_Click()
   frmMPermintaanSample.SetFocus
End Sub

Private Sub MnPlanOrder_Click()
   If FrmPlanned.Enabled = True Then FrmPlanned.SetFocus
End Sub


Private Sub mnPOBlanked_Click()
   FrmPOBlanked.SetFocus
End Sub

Private Sub mnPOBlanked1_Click()
   FrmPettyCash.SetFocus
End Sub

Private Sub mnPrelot_Click()
   frmMixingMilling.SetFocus
End Sub

Private Sub MnPriceOffer_Click()
   frmSPPH.SetFocus
End Sub

Private Sub MnProdCosting_Click()
   If frmBOMCosting.Enabled = True Then frmBOMCosting.SetFocus
End Sub

Private Sub mnproductcostingAcc_Click()
   frmBOMCosting.SetFocus
End Sub

Private Sub mnProsesProd_Click()
   If frmProduksi.Enabled = True Then frmProduksi.SetFocus
End Sub

Private Sub mnPurchaseOrder_Click()
   If FrmPurchasing.Enabled = True Then FrmPurchasing.SetFocus
End Sub

Private Sub mnPurchaseReturn_Click()
   FrmReturBeli.SetFocus
End Sub

Private Sub mnQuality_Click()
   Shell App.Path + "\LABSYS.EXE", vbNormalFocus
End Sub

Private Sub mnRegional_Click()
If frmRegional.Enabled = True Then frmRegional.SetFocus
End Sub

Private Sub mnReturBeli_Click()
If FrmReturBeli.Enabled = True Then FrmReturBeli.SetFocus
End Sub

Private Sub MnRequest_Click()
   frmListSPP.SetFocus
End Sub

Private Sub mnReturCust_Click()
   FrmReturJual.SetFocus
End Sub

Private Sub mnReturSupp_Click()
   FrmReturBeli.SetFocus
End Sub

Private Sub MnSalesFCast_Click()
If FrmSalesForecast.Enabled = True Then FrmSalesForecast.SetFocus
End Sub

Private Sub mnSalesQuote_Click()
   frmSalesQuote.SetFocus
End Sub

Private Sub mnSalesReturn_Click()
   'If FrmReturJual.Enabled = True Then FrmReturJual.SetFocus
   FrmCustFeedBack.SetFocus
End Sub

Private Sub mnRn_Click()
If frmReceiveNotes.Enabled = True Then frmReceiveNotes.SetFocus
End Sub

Private Sub mnRsc_Click()
If FrmResource.Enabled = True Then FrmResource.SetFocus
End Sub

Private Sub mnRscType_Click()
If FrmResourceType.Enabled = True Then FrmResourceType.SetFocus
End Sub

Private Sub mnSalesOrder_Click()
   If frmSalesContract.Enabled = True Then frmSalesContract.SetFocus
End Sub

Private Sub mnSalesTeam_Click()
   frmSalesTeam.SetFocus
End Sub

Private Sub mnSchedule_Click()
   frmMPSNew.SetFocus
End Sub

Private Sub mnSetupAccount_Click()
If FrmSetupAccount.Enabled = True Then FrmSetupAccount.SetFocus
End Sub

Private Sub mnShowMenu_Click()
SemeruTree1.Visible = True
End Sub

'Private Sub mnStatusIP_Click()
'   frmStatusInPorses.SetFocus
'End Sub

Private Sub mnSTPJadi_Click()
   FrmKirimProduk.SetFocus
End Sub

Private Sub mnSup_Click()
If frmMasterSup.Enabled = True Then frmMasterSup.SetFocus
End Sub

Private Sub mnTambahGroupMenu_Click()
'FrmPolicy.AddNode "tambah", "Group Menu"
End Sub

Private Sub mnTdep_Click()
Call frmEmployess.mnTdep_Click
End Sub

Private Sub mntemReference1_Click()
   FrmItemDescriptor.SetFocus
End Sub

Private Sub MnTerimaRL_Click()
   frmPenerimaanRL.SetFocus
End Sub

Private Sub mnterminpembayaran_Click()
If frmTermBayar.Enabled = True Then frmTermBayar.SetFocus
End Sub

Private Sub mnTipeBayar_Click()
If frmBebanPembayaran.Enabled = True Then frmBebanPembayaran.SetFocus
End Sub

Private Sub mnTipeDes_Click()
If FrmDescriptor.Enabled = True Then FrmDescriptor.SetFocus
End Sub

Private Sub mnTipeItem_Click()
   frmItemTrans.SetFocus
End Sub

Private Sub mnTransID_Click()
   FrmTransIDSetup.SetFocus
End Sub

Private Sub mnTransport_Click()
If frmTransport.Enabled = True Then frmTransport.SetFocus
End Sub

Private Sub mnTTp_Click()
CloseAllForm
End Sub

'Private Sub mnUpdateHarga_Click()
''frmItemPrice.SetFocus
'End Sub

Private Sub mnTukasKas_Click()
If FrmPenukaranSetaraKas.Enabled = True Then FrmPenukaranSetaraKas.SetFocus
End Sub

Private Sub mnTypeCost_Click()
If FrmCostElement.Enabled = True Then FrmCostElement.SetFocus
End Sub

Private Sub mnuKonfigurasiProsedur_Click()
'FormProsesConfig1.SetFocus
End Sub

Private Sub mnuKonfigurasiSampleForm_Click()
FormFormulaEkstraksi1.SetFocus
End Sub

Private Sub mnuMasterAnalisa_Click()
FormAnalysis.SetFocus
End Sub

Private Sub mnuMasterProses_Click()
FormProses.SetFocus
End Sub

Private Sub mnUOM_Click()
   FrmUOM.SetFocus
End Sub

Private Sub mnupermintaanbarangACC_Click()
 
End Sub

Private Sub mnUpdatePatch_Click()
   frmPatch.SetFocus
End Sub

Private Sub mnUserArea_Click()
If UCase(MainMenu.StatusBar1.Panels(1).Text) = "SA" Or UCase(MainMenu.StatusBar1.Panels(1).Text) = "ADMINISTRATOR" Then
   SemeruTree1.Visible = False
   CloseAllForm
   If FrmPolicy.Enabled = True Then FrmPolicy.SetFocus
Else
   MessageBox "Anda tidak mempunyai hak untuk mengatur akses user.", "Peringatan", msgOkOnly, msgExclamation
End If
End Sub

Private Sub mnValidasi_Click()
'FrmValidasi.SetFocus
End Sub

Private Sub mnVertical_Click()
MainMenu.Arrange vbTileVertical
End Sub

Private Sub mnVoucher_Click()
'If frmVoucher.Enabled = True Then frmVoucher.SetFocus
End Sub

Private Sub OutSPH_Click()
   frmOutstandingSPH.SetFocus
End Sub

Private Sub SemeruTree1_CloseMe()
   If SemeruTree1.Visible Then
      SemeruTree1.Visible = False
   Else
      SemeruTree1.Visible = True
   End If
End Sub

Private Sub SemeruTree1_NodeClick(ByVal Node As MSComctlLib.INode)
'On Error Resume Next
Dim xFind As clsFindHistory


If UCase(Node.Key) <> "MASTERPERKIRAAN" And UCase(Node.Key) <> "SETUPACCOUNT" Then
   If IsConfigReady = False Then
      MessageBox "Master Perkiraan Belum komplet.", "Peringatan", msgOkOnly
      Exit Sub
   End If
End If

Set AINode = Node
Set xFind = New clsFindHistory
Select Case UCase(Node.Key)
   'APPLICATION
   'MASTER
    Case "CURRSETUP": If FrmCurrencySetup.Enabled = True Then FrmCurrencySetup.SetFocus
    Case "EXCMAINTENANCE": If FrmCurrencyMaint.Enabled = True Then FrmCurrencyMaint.SetFocus
    Case "MASTERGUDANG": If FormWH.Enabled = True Then FormWH.SetFocus
    Case "MASTERKELOMPOK": If frmKelompok.Enabled = True Then frmKelompok.SetFocus
    Case "INVCARD": If FrmItemData.Enabled = True Then FrmItemData.SetFocus
    Case "INVMANAGER": mnInvManager_Click
    Case "INVPURCH": mnInvPurch_Click
    Case "MASTERREGIONAL": If frmRegional.Enabled = True Then frmRegional.SetFocus
    Case "ENTRITRANSPORT": If frmTransport.Enabled = True Then frmTransport.SetFocus
    Case "ENTRIKARYAWAN": If frmEmployess.Enabled = True Then frmEmployess.SetFocus
    Case "MASTERFREIGHT": If frmBebanPembayaran.Enabled = True Then frmBebanPembayaran.SetFocus
    Case "ENTRISUPPLIER": If frmMasterSup.Enabled = True Then frmMasterSup.SetFocus
    Case "CUSTCARD": If frmPartner.Enabled = True Then frmPartner.SetFocus
    Case "INVPRODUKSI": If FrmItemData.Enabled = True Then FrmItemData.SetFocus
    Case "ENTRIBANK": If frmBankPartner.Enabled = True Then frmBankPartner.SetFocus
    Case "ENTRIGUDANG": If FrmGudangCust.Enabled = True Then FrmGudangCust.SetFocus
    Case "ENTRYUOM": FrmUOM.SetFocus
    Case "TERMPAYMENT": frmTermBayar.SetFocus
    Case "TIPEITEM": frmItemTrans.SetFocus
    Case "ITEMCHARGE": frmItemCharge.SetFocus
   
    'PURCHASE
    
    Case "TRANSAKSIPO": If FrmPurchasing.Enabled = True Then FrmPurchasing.SetFocus
    Case "PURCHOUTSPH": If frmOutstandingSPH.Enabled = True Then frmOutstandingSPH.SetFocus
    Case "LOGREQUEST": If frmListSPP.Enabled Then frmListSPP.SetFocus
    Case "SPPH": If frmSPPH.Enabled Then frmSPPH.SetFocus
    Case "TRANSAKSISPPH": If frmSPH.Enabled Then frmSPH.SetFocus
    Case "PURCHRETUR": If FrmReturBeli.Enabled Then FrmReturBeli.SetFocus
    Case "PLANORDER": If FrmPlanned.Enabled = True Then FrmPlanned.SetFocus
    Case "POBLANKED": If FrmPOBlanked.Enabled = True Then FrmPOBlanked.SetFocus
    Case "POBLANKED1": If FrmPettyCash.Enabled = True Then FrmPettyCash.SetFocus
    'Purchase Approval
    Case "APSPP": If mnAppSPP.Enabled = True Then mnAppSPP_Click
    
    Case "APPORDPEMB":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "PurchaseID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],PurchaseID, DatePurchase as Tanggal,empID as [User],Date_approved as [Waktu Approved] from [PO Order] where status = 0 and type_trans_order=2 order by datePurchase desc"
        ACC.StrSQLDetail = "Select [detail PO].NoItem as Kode, inventory.internalName as [Nama Barang], UOM as Satuan,QtyPO as Jumlah,POPrice as [Harga Satuan] from [Detail PO] inner join inventory on [detail PO].NOItem = inventory.NoItem"
        ACC.MasterTable = "PO Order"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Order Pembelian"
        frmApproval.SetFocus

   Case "APPRPBRL":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "PurchaseID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],PurchaseID, DatePurchase as Tanggal,empID as [User],Date_approved as [Waktu Approved] from [PO Order] where status = 0 and type_trans_order=3 order by datePurchase desc"
        ACC.StrSQLDetail = "Select [detail PO].NoItem as Kode, inventory.internalName as [Nama Barang], UOM as Satuan,QtyPO as Jumlah,POPrice as [Harga Satuan] from [Detail PO] inner join inventory on [detail PO].NOItem = inventory.NoItem"
        ACC.MasterTable = "PO Order"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Rencana Pembelian Bulanan Rumput Laut"
        frmApproval.SetFocus
        
   Case "APPEVASUPP":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "ID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],ID,CompanyName as Supplier,Date_approved as [Waktu Approved] from [evaluasi_supplier]"
        ACC.StrSQLDetail = "Select [EVALUASI_supplier_detail].NoItem as Kode, inventory.internalName as [Nama Barang], UOM as Satuan,jml_order as [Jml Order], jml_datang as [Jml Datang], jml_selisih as [Jml Selisih], jml_reject as [Jml Reject] from evaluasi_supplier_detail inner join inventory on evaluasi_supplier_detail.NOItem = inventory.NoItem"
        ACC.MasterTable = "evaluasi_supplier"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Evaluasi Supplier"
        frmApproval.SetFocus
      
   Case "APPPERHARGA":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "SPPHID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],SPPHID,DateTrans as Tanggal,CompanyName as Supplier, userReqst as [User] ,Date_approved as [Waktu Approved] from QueryPurchaseOffer where status=0 order by dateTrans desc"
        ACC.StrSQLDetail = "Select spph_line.NoItem as Kode, inventory.internalName as [Nama Barang], UOM as Satuan from spph_line inner join inventory on spph_line.NOItem = inventory.NoItem"
        ACC.MasterTable = "spph_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Permintaan Penawaran Harga"
        frmApproval.SetFocus
   
   Case "APPSURATRETUR":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "TransID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],TransID as ReturID,DateTrans as Tanggal,CompanyName as Supplier, empID as [User] ,Date_approved as [Waktu Approved],transID from transData inner join partnerdb on transData.partnerid = partnerdb.partnerID where status=0 order by dateTrans desc"
        ACC.StrSQLDetail = "Select NoItem as Kode, itemname as [Nama Barang], UOM as Satuan,alasan from queryReturBeli"
        ACC.MasterTable = "transdata"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Retur Pembelian"
        frmApproval.SetFocus
   
    
    'EKSEKUSI
    Case "PURCHOUTSTANDING": If frmOutstandingPO.Enabled = True Then frmOutstandingPO.SetFocus
    Case "TRANSAKSIAP": If FrmInvoice.Enabled Then FrmInvoice.SetFocus
    Case "PURCHEVALUASI": If FrmEvaluasiSupplier.Enabled Then FrmEvaluasiSupplier.SetFocus
    
  
    'SALES
    Case "SALESQUOTE": If frmSalesQuote.Enabled = True Then frmSalesQuote.SetFocus
    Case "SALESORDER": If frmSalesOrder.Enabled = True Then frmSalesOrder.SetFocus
    Case "SALESKONTRAK": If frmSalesContract.Enabled = True Then frmSalesContract.SetFocus
    Case "SALESRETUR": If FrmCustFeedBack.Enabled = True Then FrmCustFeedBack.SetFocus
    Case "SALESCAST": FrmSalesForecast.SetFocus
    Case "SALESINVOICE": If frmArTrans.Enabled = True Then frmArTrans.SetFocus
    Case "MKTFEEDBACK": If FrmCustFeedBack.Enabled Then FrmCustFeedBack.SetFocus
    Case "MKTMEMO": If frmMemoPotongHarga.Enabled Then frmMemoPotongHarga.SetFocus
    Case "MKTCONTACT": If frmcontact.Enabled = True Then frmcontact.SetFocus
    Case "SALESTIM": If frmSalesTeam.Enabled = True Then frmSalesTeam.SetFocus
    Case "OUTSTANDINGMKT": If frmOutstandingMkt.Enabled = True Then frmOutstandingMkt.SetFocus
    'History
    Case "MKTSALESQUOTE": xFind.SetHistory "SalesQuoteValid", frmSalesQuote, "Sales Quote", "Customer"
    Case "MKTKONTRAKPENJUALAN": xFind.SetHistory "SalesContractValid", frmSalesContract, "Sales Contract", "Customer"
    Case "MKTORDERPENJUALAN": xFind.SetHistory "SalesOrderValid", frmSalesOrder, "Sales Order", "Customer"
    Case "MKTINVPENJUALAN": xFind.SetHistory "InvoicePenjualanValid", frmArTrans, "Invoice Penjualan", "Customer"
    
    'APPROVAL
    Case "APPSALESORDER":  If mnAPPSO.Enabled = True Then mnAPPSO_Click
    Case "APPPERMINTAANSAMPLE": If mnAPPPermintaanSample = True Then mnAPPSample
    Case "APPMEMO": If mnAPPMemo = True Then mnAPPMemoPotongan
    Case "APPCUSTOMERFEEDBACK": If mnAPPCustomerFeedback = True Then mnAPPCustFeedBack
    

    
    'LOGISTIK
    Case "LOGTERIMA": If frmPenerimaanRL.Enabled Then frmPenerimaanRL.SetFocus
    Case "LOGRAWSUPPORT": If frmReceiveNotes.Enabled = True Then frmReceiveNotes.SetFocus
    Case "LOGSJ":  If FrmDO.Enabled = True Then FrmDO.SetFocus
    Case "ORDERREQUEST": If FrmPRequest.Enabled Then FrmPRequest.SetFocus
    Case "MATERIALREQUEST":
               If FrmMatRequest.Enabled Then
                  FrmMatRequest.SetFocus
               End If
                  
    Case "LOGOUT":
            If FrmSFIssue.Enabled Then
               FrmSFIssue.Mode = True
               FrmSFIssue.SetFocus
            End If
            
    Case "LOGRETURSUPP": If FrmReturBeli.Enabled Then FrmReturBeli.SetFocus
    Case "LOGRFG": If FrmSFComplete.Enabled Then FrmSFComplete.SetFocus
    Case "LOGRETURCUST": If FrmReturJual.Enabled Then FrmReturJual.SetFocus
    Case "MKTSAMPLE": If frmMPermintaanSample.Enabled Then frmMPermintaanSample.SetFocus
    Case "LOGDO": If FrmDO.Enabled Then FrmDO.SetFocus
    Case "SPBAPPROVAL": If mnAPSPB.Enabled Then mnAPSPB_Click
    Case "STOCKOPNAME": FrmInvAdj.SetFocus
    Case "STOCKBROWSER": frmInventory.SetFocus
    
    'GUDANG
    Case "GUDANGTERIMA": If frmPenerimaanRL.Enabled Then frmPenerimaanRL.SetFocus
    Case "GUDANGKIRIM": If FrmSJWarehouse.Enabled Then FrmSJWarehouse.SetFocus
    Case "LEMBARSUPPLIER": If frmLembarSupplier.Enabled Then frmLembarSupplier.SetFocus
    Case "KONFIGURASIRLBATCH": If frmConfigRL.Enabled Then frmConfigRL.SetFocus
    
    '* PPROVAL GUDANG RL
    Case "APPTTRL":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        'ACC.Key = "TransID"
        ACC.Key = "ID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],ID ,tgl as Tanggal,CompanyName as Supplier, empID as [User] ,Date_approved as [Waktu Approved] from QueryPenerimaanRL where status=0 order by tgl desc"
        ACC.StrSQLDetail = "Select QtyPO as [Qty PO], QtyReceive as [Qty Diterima] from QueryPenerimaanRL "
        ACC.MasterTable = "transdata"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Tanda Terima RL"
        frmApproval.SetFocus

    Case "APPKRL":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "TransID"
        ACC.StrSQLMaster = "SELECT transdata.Approved_by as [Approved By],transdata.TransID, transdata.EmpID as [User], transdata.DateTrans" & _
            ",transdata.[No Pol],transdata.tujuan, transdata.person " & _
            " FROM transdata LEFT OUTER JOIN WareHouse ON transdata.tujuan = WareHouse.WareHouse " & _
            " LEFT OUTER JOIN WareHouse AS WareHouse_1 ON transdata.WareHouse = WareHouse_1.WareHouse " & _
            " WHERE (transdata.TypeTrans = N'SS')"
        ACC.StrSQLDetail = "SELECT * FROM QDetailSJGudang"
        ACC.MasterTable = "transdata"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Pengiriman RL"
        frmApproval.SetFocus
    
    Case "APPLEMBSUPP":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "TransID"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],TransID ,DateTrans as Tanggal,CompanyName as Supplier, empID as [User] ,Date_approved as [Waktu Approved] from view_lembar_supplier where status=0 order by dateTrans desc"
        ACC.StrSQLDetail = "Select Qty_receive,kondisi,cuci,jemur,sortir,napel,packing from view_lembar_supplier"
        ACC.MasterTable = "transdata"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Lembar Supplier"
        frmApproval.SetFocus
    
   'HISTORY PO
    Case "PURCHHIST1": xFind.SetHistory "SPPHValid", frmSPPH, "Permintaan Penawaran Harga", "Supplier"
    Case "PURCHHIST2": xFind.SetHistory "SPPVAlid", FrmPRequest, "Permintaan Pembelian", "Supplier"
    Case "HISTPO":     xFind.SetHistory "PurchaseOrderValid", FrmPurchasing, "Purchase Order", "Supplier"
    Case "PURCHHIST4": xFind.SetHistory "ReturBeliValid", FrmReturBeli, "Retur Pembelian", "Supplier"
    Case "PURCHHIST5": xFind.SetHistory "", FrmMRP, "Planned Order MRP", "Supplier"
    Case "PURCHHIST6": xFind.SetHistory "InvoiceValid", FrmInvoice, "Invoice", "Supplier"
    'Case "PURCHHIST7": xFind.SetHistory "DeliveryValid", Frm, "Billing PPN", ""
      
   'PRODUCTION
      Case "DESCREF": If FrmItemDescriptor.Enabled = True Then FrmItemDescriptor.SetFocus
      Case "ASEMBLYA3": If frmbom.Enabled = True Then frmbom.SetFocus
      Case "ASEMBLYA2": If FrmMOrder.Enabled = True Then FrmMOrder.SetFocus
      Case "RESOURCESII": If FrmResource.Enabled = True Then FrmResource.SetFocus
      Case "WC": If FrmWorkCenter.Enabled = True Then FrmWorkCenter.SetFocus
      Case "MD":  If FrmManDescriptor.Enabled = True Then FrmManDescriptor.SetFocus
      Case "IR": If FrmItemReference.Enabled = True Then FrmItemReference.SetFocus
      Case "CALENDAR": If FrmCalendar.Enabled = True Then FrmCalendar.SetFocus
      Case "CATEGORIES":  If FrmCategories.Enabled = True Then FrmCategories.SetFocus
      Case "COST METHODE": If FrmCostElement.Enabled = True Then FrmCostElement.SetFocus
      Case "INVADJ1": If frmBOMCosting.Enabled = True Then frmBOMCosting.SetFocus
      Case "BOM METHODE": If FrmBOMMethode.Enabled = True Then FrmBOMMethode.SetFocus
      Case "DESCRIPTOR": If FrmDescriptor.Enabled = True Then FrmDescriptor.SetFocus
      Case "STAGE": If FrmRouting.Enabled = True Then FrmRouting.SetFocus
      Case "RESOURCES": If FrmResourceType.Enabled = True Then FrmResourceType.SetFocus
      Case "MUTASIGUDANG": If frmMutasiGudang.Enabled = True Then frmMutasiGudang.SetFocus
      Case "INVADJ": If FrmInvAdj.Enabled = True Then FrmInvAdj.SetFocus
      Case "MRPGEN": If FrmMRP.Enabled = True Then FrmMRP.Show vbModal
      Case "PLO": If FrmPlanned.Enabled = True Then FrmPlanned.SetFocus
      Case "ECC": If FrmEnginering.Enabled = True Then FrmEnginering.SetFocus
      Case "PRELOT": If frmMixingMilling.Enabled = True Then mnPrelot_Click
      Case "MUTASICHIP": If frmMutasiChip.Enabled = True Then mnMutasiChip_Click
      Case "LOT": If frmBlending.Enabled = True Then mnLot1_Click

      Case "MRP": If FrmMRPLogic.Enabled = True Then FrmMRPLogic.SetFocus
      Case "MPS": If frmMPSNew.Enabled = True Then frmMPSNew.SetFocus
      Case "PROD": If frmProduksi.Enabled = True Then frmProduksi.SetFocus
    '  Case "STATUSIP": If frmStatusInPorses.Enabled = True Then frmStatusInPorses.SetFocus
      Case "SERAHTERIMA": If FrmKirimProduk.Enabled = True Then FrmKirimProduk.SetFocus
      Case "SFREQUEST": If frmSFRequest.Enabled = True Then frmSFRequest.SetFocus
      Case "SFFLUSH": If frmSFBackflush.Enabled = True Then frmSFBackflush.SetFocus
      Case "SFISSUE": FrmSFIssue.SetFocus
   'HISTORY'
      Case "HISTCONCRETE":
           ' frmPconcrete.SetMode = 1
            frmPconcrete.SetFocus
      
      Case "HISTHYDRAULIC":
           ' frmpress.SetMode = 1
            frmpress.SetFocus
            
      Case "HISTJEMUR":
            'FRMJEMUR.SetMode = 1
            FrmJemur.SetFocus
      
      Case "HISTDRYER":
           ' frmDryer.SetMode = 1
            frmDryer.SetFocus
            
      Case "HISTCRUSHER":
           ' frmPCrusher.SetMode = 1
            frmPCrusher.SetFocus
            
   
      '* approval
      Case "APPALKALI":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal],[Reaktor],[No_Stock] as [No RL],[rekomNo] as [No Rekomendasi]" & _
                           ",[Berat_rl] as [Berat RL],[tempat_alkali] as [Tempat Alkali],[waktu_mulai] as [Waktu Mulai],[waktu_selesai] as [Waktu Selesai],issued_by as [User] ,Date_approved as [Waktu Approved] from alkali_treatment  order by tanggal desc "
        ACC.StrSQLDetail = "Select * from Alkali_detail"
        ACC.MasterTable = "alkali_treatment"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Alkali Treatment"
        frmApproval.SetFocus

      Case "APPACID":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal],[grup],[tanki],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai] ,ph_akhir as [pH Akhir],issued_by as [User] ,Date_approved as [Waktu Approved] from acid_treatment order by tanggal desc "
        ACC.StrSQLDetail = "Select * from Acid_detail"
        ACC.MasterTable = "acid_treatment"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Acid Treatment"
        frmApproval.SetFocus
      
      Case "APPBLEACHING":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal],[grup],[tanki],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai] ,ph_akhir as [pH Akhir],issued_by as [User] ,Date_approved as [Waktu Approved] from bleaching order by tanggal desc "
        ACC.StrSQLDetail = "Select * from bleaching_detail"
        ACC.MasterTable = "bleaching"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Bleaching Treatment"
        frmApproval.SetFocus
        
     Case "APPEKSREAKTOR":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal],[grup],[tanki],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai] ,ph_akhir as [pH Akhir],issued_by as [User] ,Date_approved as [Waktu Approved] from bleaching order by tanggal desc "
        ACC.StrSQLDetail = "Select * from bleaching_detail"
        ACC.MasterTable = "bleaching"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Bleaching Treatment"
        frmApproval.SetFocus
        
     Case "APPEKSAUTO":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal],[grup],[tanki],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai] ,ph_akhir as [pH Akhir],issued_by as [User] ,Date_approved as [Waktu Approved] from bleaching order by tanggal desc "
        ACC.StrSQLDetail = "Select * from bleaching_detail"
        ACC.MasterTable = "bleaching"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Bleaching Treatment"
        frmApproval.SetFocus
      
     Case "APPFILTERPRESS":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_press"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_press],[tanggal_press] as Tanggal,Keterangan from press_header order by tanggal_press desc"
        ACC.StrSQLDetail = "Select * from press_detail"
        ACC.MasterTable = "press_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Filter Press"
        frmApproval.SetFocus
        
     Case "APPGELL":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstrasi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstrasi],tgl_gell as [tanggal],[grup],[jml_air] as [Jml Air],[total_kempu] as [Total Kempu]" & _
                           " from Gellification order by tgl_gell desc "
        ACC.StrSQLDetail = "Select * from Gellification_Pite "
        ACC.MasterTable = "gellification"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Gellification"
        frmApproval.SetFocus
      
     Case "APPBUNGKUS":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal_bungkus] as Tanggal,[grup],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai] ,hasil_bungkus as [Hasil Bungkus],issued_by as [User] ,Date_approved as [Waktu Approved] from pembungkusan order by tanggal_bungkus desc "
        ACC.StrSQLDetail = "Select * from pembungkusan"
        ACC.MasterTable = "pembungkusan"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Pembungkusan"
        frmApproval.SetFocus
        
     Case "APPCONCRETE":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal_ekstraksi] as Tanggal,[grup],issued_by as [User] ,Date_approved as [Waktu Approved] from concrete_header order by tanggal_ekstraksi desc "
        ACC.StrSQLDetail = "Select * from concrete_detail"
        ACC.MasterTable = "concrete_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Concrete Press"
        frmApproval.SetFocus
      
     Case "APPHYDRAULIC":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_press"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_press],[tanggal_press] as Tanggal,[grup]from press_header order by tanggal_press   desc "
        ACC.StrSQLDetail = "Select * from press_detail"
        ACC.MasterTable = "press_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Hydraulic Press"
        frmApproval.SetFocus
        
     Case "APPCUTTER":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal_cutter] as Tanggal,[grup],[tgl_mulai] as [Waktu Mulai]" & _
                           ",tgl_selesai as [waktu Selesai] from cutter order by tanggal_cutter desc "
        ACC.StrSQLDetail = "Select * from cutter"
        ACC.MasterTable = "cutter"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Cutter"
        frmApproval.SetFocus
      
     Case "APPJEMUR":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_ekstraksi"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[No_ekstraksi],[tanggal],[grup],[tanki],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai] ,ph_akhir as [pH Akhir],issued_by as [User] ,Date_approved as [Waktu Approved] from bleaching order by tanggal desc "
        ACC.StrSQLDetail = "Select * from bleaching_detail"
        ACC.MasterTable = "bleaching"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Bleaching Treatment"
        frmApproval.SetFocus
   
     
     Case "APPCRUSHER":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "no_crusher"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],no_crusher,[tanggal_crusher] as Tanggal,[grup],issued_by as [User] ,Date_approved as [Waktu Approved] from crusher_header order by tanggal_crusher desc "
        ACC.StrSQLDetail = "Select * from crusher_detail"
        ACC.MasterTable = "crusher_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Crusher"
        frmApproval.SetFocus
        
     Case "APPMIXING":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "prelot"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[Prelot],[grup],[tanggal_mulai] as [Waktu Mulai]" & _
                           ",tanggal_selesai as [waktu Selesai], issued_by as [User] ,Date_approved as [Waktu Approved] from mixing_header order by tanggal_mulai desc "
        ACC.StrSQLDetail = "Select * from mixing_detail"
        ACC.MasterTable = "mixing_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Proses Mixing & Milling"
        frmApproval.SetFocus
   
     Case "APPBLENDING":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "lotNo"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],[lotno],[grup],mesh,[tanggal_mulai_blending] as [Waktu Mulai]" & _
                           ",tanggal_selesai_blending as [waktu Selesai] ,issued_by as [User] ,Date_approved as [Waktu Approved] from blending_header order by tanggal_mulai_blending desc "
        ACC.StrSQLDetail = "Select * from blending_detail"
        ACC.MasterTable = "blending_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Blending Instruction"
        frmApproval.SetFocus
        
     Case "APPSTFG":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "IDTrans"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],IDTrans,DateTRans as [tanggal],[issued by] as [User] ,Date_approved as [Waktu Approved] from backflush_header where status = 0 and TypeTrans = 'FG' order by dateTRans desc "
        ACC.StrSQLDetail = "Select * from backflush_line"
        ACC.MasterTable = "backflush_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Serah terima Produk Jadi"
        frmApproval.SetFocus
   
     Case "APPPAKAI":
        ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
        ACC.Key = "IDTrans"
        ACC.StrSQLMaster = "Select approved_by as [Approved By],IDTrans,DateTRans as [tanggal],[issued by] as [User] ,Date_approved as [Waktu Approved] from backflush_header where status = 0 and TypeTrans = 'FG' order by dateTRans desc "
        ACC.StrSQLDetail = "Select * from backflush_line"
        ACC.MasterTable = "backflush_header"
        frmApproval.Validasi = ACC
        frmApproval.Caption = "Approval Serah terima Produk Jadi"
        frmApproval.SetFocus
   
   'ACCOUNTING
      Case "MASTERPERKIRAAN": If FrmPerkiraan.Enabled = True Then FrmPerkiraan.SetFocus
      Case "VOUCHERTRANSAKSI":  If frmVcrBeli.Enabled = True Then frmVcrBeli.SetFocus
      Case "VOUCHERTRANSAKSI2":  If FrmVcrJual.Enabled = True Then FrmVcrJual.SetFocus
     
     ' Case "TUNAIBIAYA": If FrmPengeluaranBiaya.Enabled = True Then FrmPengeluaranBiaya.SetFocus
      Case "PERMINTAANBARANG": If FrmMatRequest.Enabled = True Then FrmMatRequest.SetFocus
      Case "TUNAIBIAYA": If FrmBKK.Enabled = True Then FrmBKK.SetFocus
      Case "BAYARTUNAILAIN": If FrmBKM.Enabled = True Then FrmBKM.SetFocus
      Case "VALIDASIJOURNAL": 'FrmValidasi.SetFocus
      Case "CLOSING": If frmValidasi.Enabled = True Then frmValidasi.SetFocus
      Case "SETUPACCOUNT": If FrmSetupAccount.Enabled = True Then FrmSetupAccount.SetFocus
      Case "KONFIGPERIODE": If FrmSetingPeriode.Enabled = True Then FrmSetingPeriode.SetFocus
      Case "DOUBLEENTRY":  If frmMemorial.Enabled = True Then frmMemorial.SetFocus
      Case "INVMEMO": If frmInvMemo.Enabled = True Then frmInvMemo.SetFocus
      Case "PENUKARAN": If FrmPenukaranSetaraKas.Enabled = True Then FrmPenukaranSetaraKas.SetFocus
      Case "HPP": If frmHPP.Enabled = True Then frmHPP.SetFocus
        
   'WAREHOUSE
      Case "GUDANGTERIMA": If frmPenerimaanRL.Enabled Then frmPenerimaanRL.SetFocus
      
   'FIXED ASSET
'      Case "FAPURCHASE": If FrmPembelianFixAssets.Enabled = True Then FrmPembelianFixAssets.SetFocus
'      Case "FASALES": If FrmPenjualanFixAssets.Enabled = True Then FrmPenjualanFixAssets.SetFocus
'      Case "MASTERAKTIVA": If FrmMasterFixAssets.Enabled = True Then FrmMasterFixAssets.SetFocus
'      Case "FISCAL": FrmSetingPeriode.SetFocus
'      Case "QUARTER": FrmQuarter.SetFocus
'      Case "BOOK": FrmBookSetup.SetFocus
'      Case "CLASS": FrmClassSetup.SetFocus
'      Case "ACCGROUP": FrmAccGroup.SetFocus
'      Case "NUMBERING": FrmAssetsBook.SetFocus
'      Case "RETIREMENT": FrmRetirementMaintenance.SetFocus
'      Case "TRANSFER": FrmTransferMaintenance.SetFocus
      
'       Case "TRANSAKSICLOSESJ": FrmCloseSJ.SetFocus
'       Case "ASEMBLYA": frmItemAsm.SetFocus
'       Case "MASTERAKUNBANK": FrmEntriKas.SetFocus
'       Case "KONFIGB": FrmConfigAccount.SetFocus
'       Case "MUTASIJUAL": If FrmMutasiPenjualan.Enabled = True Then FrmMutasiPenjualan.SetFocus
'       Case "PIUTANGKARYAWAN": If FrmPiutangKaryawan.Enabled = True Then FrmPiutangKaryawan.SetFocus
'       Case "MASTERFGROUP": frmListingFixAssets.SetFocus
'       Case "BAYARPIUTANGKARYAWAN": 'If frmPembayaranPKaryawan.Enabled = True Then frmPembayaranPKaryawan.SetFocus
'       Case "MASTERBIAYA": FrmTabelBiaya.SetFocus
'       Case "MASTERAKUNKAS": FrmSetupKas.SetFocus
'       Case "TUNAILAIN": frmBKK.SetFocus
'       Case "SETINGJOURNAL": FrmConfigAccount.SetFocus
      Case Else:
End Select
Set xFind = Nothing
Err.Clear
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ToolErr
Dim hwnd As Long
'Dim StartDoc As Log

'CloseAllForm
If SemeruTree1.Visible = False Then SemeruTree1.Visible = True
MainMenu.StatusBar1.Panels(5).Text = "Menu : " & Button.Caption

Select Case Button.Index
       Case 1: myMenu.CreateMenu "MASTER"
       Case 3: myMenu.CreateMenu "PURCHASE"
       Case 5: myMenu.CreateMenu "SALES"
       Case 7: myMenu.CreateMenu "LOGISTIK"
       Case 9: myMenu.CreateMenu "GUDANGRL"
       Case 11: myMenu.CreateMenu "PRODUKSI"
       Case 13: myMenu.CreateMenu "AKUNTING"
       Case 15: Shell App.Path + "\LABSYS.EXE", vbNormalFocus
       Case 19: Shell App.Path + "\MMT.EXE", vbNormalFocus
       Case 17: hwnd = apiFindWindow("OPUSAPP", "0")
                ShellExecute hwnd, "open", App.Path & "\HRIS-ASML\payroll.exe", "/d " & aksess.GetID, App.Path & "\HRIS-ASML", SW_SHOWNORMAL
       Case 21: frmReport.SetFocus
End Select
Exit Sub

ToolErr:
    MessageBox Err.Description, "MainMenu - Toolbar1_ButtonClick", msgOkOnly, msgCrtical
End Sub

Private Sub OpenMenu()
'MainMenu.Toolbar1.Buttons(1).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Master Data") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Master Data"), False))
'MainMenu.Toolbar1.Buttons(2).Visible = MainMenu.Toolbar1.Buttons(1).Visible
'MainMenu.Toolbar1.Buttons(3).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Distribution") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Distribution"), False))
'MainMenu.Toolbar1.Buttons(4).Visible = MainMenu.Toolbar1.Buttons(3).Visible
'MainMenu.Toolbar1.Buttons(5).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Produksi") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Production"), False))
'MainMenu.Toolbar1.Buttons(6).Visible = MainMenu.Toolbar1.Buttons(5).Visible
'MainMenu.Toolbar1.Buttons(7).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Akunting") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Accounting"), False))
'MainMenu.Toolbar1.Buttons(8).Visible = MainMenu.Toolbar1.Buttons(7).Visible
'MainMenu.Toolbar1.Buttons(9).Visible = False
'MainMenu.Toolbar1.Buttons(10).Visible = False
'MainMenu.Toolbar1.Buttons(11).Visible = False ' MainMenu.Toolbar1.Buttons(1).Visible + MainMenu.Toolbar1.Buttons(3).Visible + MainMenu.Toolbar1.Buttons(5).Visible + MainMenu.Toolbar1.Buttons(7).Visible
End Sub

Private Function SeekFormByTag(ByVal FormTag As String) As Boolean
Dim I As Integer
Dim frm As Form
On Error GoTo Hell

For Each frm In Forms
    If UCase(frm.Tag) = UCase(FormTag) Then
       SeekFormByTag = True
       frm.ZOrder (0)
    End If
Next
Set frm = Nothing
Hell:
    Err.Clear
End Function


Private Sub Vp_Click()
frmVcrBeli.Show
End Sub

Private Sub Vpe_Click()
FrmVcrJual.Show
End Sub

Private Sub mnAPPSample()
'Approval Permintaan Sample
   ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
   ACC.Key = "nomor"
   ACC.StrSQLMaster = "Select approved_by as [Approved By],Nomor, Tanggal,ordered_by as [User], Date_approved as [Waktu Approved] from [permintaansample]"
   ACC.StrSQLDetail = "select NoItem, InternalName as [Nama Produk], UOM as Satuan, Jumlah, Tanggal_butuh as [Tanggal Di Butuhkan], CompanyName, Keterangan from QueryPermintaanSample"
   ACC.MasterTable = "permintaansample"
   frmApproval.Validasi = ACC
   frmApproval.Caption = "Validasi Permintaan Sample"
   frmApproval.SetFocus
End Sub

Private Sub mnAPPMemoPotongan()
'Approval Memo Potongan
   ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
   ACC.Key = "[Memo ID]"
   ACC.StrSQLMaster = "Select approved_by as [Approved By],[Memo ID], [Date Memo],ordered_by as [User], Date_approved as [Waktu Approved] from [Memo Potongan Harga]"
   ACC.StrSQLDetail = "select * from [querymemopotonganharga]"
   ACC.MasterTable = "[Memo Potongan Harga]"
   frmApproval.Validasi = ACC
   frmApproval.Caption = "Validasi Memo Potongan Harga"
   frmApproval.SetFocus
End Sub


Private Sub mnAPPCustFeedBack()
'Approval Memo Potongan
   ACC.ApprovalField "Approved_by", MainMenu.StatusBar1.Panels(1).Text, "Date_approved"
   ACC.Key = "[Feedback ID]"
   ACC.StrSQLMaster = "Select approved_by as [Approved By],[Feedback ID], [Effective Date],[Date Receipt],ordered_by as [User], Date_approved as [Waktu Approved] from [customer feedback]"
   ACC.StrSQLDetail = "select * from [querycustomerfeedback]"
   ACC.MasterTable = "[customer feedback]"
   frmApproval.Validasi = ACC
   frmApproval.Caption = "Validasi Customer FeedBack"
   frmApproval.SetFocus
End Sub
