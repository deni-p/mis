VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{26FFC84F-9563-4CC0-8A44-D5D7F61389A7}#1.0#0"; "SemeruDC.ocx"
Begin VB.MDIForm MainMenu2 
   BackColor       =   &H8000000C&
   Caption         =   "Manu"
   ClientHeight    =   4620
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10995
   Icon            =   "MainMenu2.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MainMenu2.frx":6852
   StartUpPosition =   2  'CenterScreen
   Tag             =   "MAIN"
   WindowState     =   2  'Maximized
   Begin SemeruDC.SemeruTree SemeruTree1 
      Align           =   3  'Align Left
      Height          =   3720
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   6562
      BackColorTree   =   16744576
      BackColorBackground=   -2147483643
   End
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
         NumListImages   =   33
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":2D2E0
            Key             =   "Main"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":2E562
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":2FC4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":31336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":32A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3410A
            Key             =   "Expenses"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":357F4
            Key             =   "Akunting"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":35B8E
            Key             =   "BKM"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":36680
            Key             =   "ASSETS"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":37172
            Key             =   "Data Master"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":37C64
            Key             =   "BKK"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":38756
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":39248
            Key             =   "Anak Akun"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":395E2
            Key             =   "Master"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3A0D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3ABC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3B6B8
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3C1AA
            Key             =   "Transaksi"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3C544
            Key             =   "Bayar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3D036
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3DB28
            Key             =   "Retur"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3E61A
            Key             =   "ar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3F10C
            Key             =   "ap"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":3FBFE
            Key             =   "Konfig"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":406F0
            Key             =   "Validasi"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":41742
            Key             =   "KonfigReport"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":42234
            Key             =   "Memo"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":427CE
            Key             =   "Fix Assets"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":42B68
            Key             =   "MASET"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":42F02
            Key             =   "TASET"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":4329C
            Key             =   "PRODUKSIPLAN"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":43636
            Key             =   "WHouse"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu2.frx":49E98
            Key             =   "History"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   4290
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Admin"
            TextSave        =   "Admin"
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Server"
            TextSave        =   "Server"
            Object.ToolTipText     =   "Server Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9128
            MinWidth        =   9128
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "22:39"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1005
      ButtonWidth     =   1746
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Master Data"
            Object.ToolTipText     =   "Master Data"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribution"
            Object.ToolTipText     =   "Distribution"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Production"
            Object.ToolTipText     =   "Production"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Accounting"
            Object.ToolTipText     =   "Accounting"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fixed Asset"
            Object.ToolTipText     =   "Fixed Asset"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            Object.ToolTipText     =   "Report"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnApp 
      Caption         =   "Application"
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
         Caption         =   "Report Administration"
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
         Caption         =   "User Authentication"
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
         End
         Begin VB.Menu mnExAccount 
            Caption         =   "Posting Account Setup"
         End
         Begin VB.Menu mnExMaint 
            Caption         =   "Exchange Maintenance"
         End
      End
      Begin VB.Menu mnGudang 
         Caption         =   "Warehouse"
      End
      Begin VB.Menu mnItem 
         Caption         =   "Inventory"
         Begin VB.Menu mnInvCard 
            Caption         =   "Inventory Card"
         End
         Begin VB.Menu mnKelompok 
            Caption         =   "Inventory Class"
         End
      End
      Begin VB.Menu mnRegional 
         Caption         =   "Regional"
      End
      Begin VB.Menu mnTipeBayar 
         Caption         =   "Payment"
      End
      Begin VB.Menu mnTransport 
         Caption         =   "Transporter"
      End
      Begin VB.Menu mnKaryawan 
         Caption         =   "Employee"
      End
      Begin VB.Menu mnCus 
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
   End
   Begin VB.Menu mnTrans 
      Caption         =   "Distribution"
      Begin VB.Menu mnGap 
         Caption         =   "Purchase Processing"
         Begin VB.Menu MnPO 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu MnPlanOrder 
            Caption         =   "Planned Order"
         End
      End
      Begin VB.Menu mnGar 
         Caption         =   "Sales Processing"
         Begin VB.Menu mnSc 
            Caption         =   "Sales Order"
         End
         Begin VB.Menu MnSFCast 
            Caption         =   "Sales Forecast"
         End
         Begin VB.Menu mnAr 
            Caption         =   "Invoicing"
         End
      End
      Begin VB.Menu mnRetBarang 
         Caption         =   "Retur Management"
         Begin VB.Menu mnReturBeli 
            Caption         =   "Purchase Retur"
         End
         Begin VB.Menu mnReturJual 
            Caption         =   "Sales Retur"
         End
      End
      Begin VB.Menu MnWHouse 
         Caption         =   "Warehouse"
         Begin VB.Menu mnRn 
            Caption         =   "Goods Receive Note"
         End
         Begin VB.Menu mnDn 
            Caption         =   "Delivery Order"
         End
         Begin VB.Menu mnCloseSJt 
            Caption         =   "Delivery Order Closing"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MnTrans1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBayar 
         Caption         =   "Pembayaran"
         Visible         =   0   'False
         Begin VB.Menu mnHupiut 
            Caption         =   "Hutang/Piutang"
         End
         Begin VB.Menu mnKas 
            Caption         =   "Pengeluaran Kas Harian"
         End
      End
   End
   Begin VB.Menu mnInventory 
      Caption         =   "Production"
      Begin VB.Menu mnEntupProduksi 
         Caption         =   "Setup"
         Begin VB.Menu mnTypeCost 
            Caption         =   "Cost Methode"
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
         End
         Begin VB.Menu mnCountPoint 
            Caption         =   "Stage/Count Point"
         End
         Begin VB.Menu mnRsc 
            Caption         =   "Resources"
         End
         Begin VB.Menu mnBomMethode 
            Caption         =   "BOM Methode"
         End
         Begin VB.Menu mnJobCosting 
            Caption         =   "BOM Costing"
         End
         Begin VB.Menu mnBomBom 
            Caption         =   "Bill Of Material"
         End
         Begin VB.Menu mnECC 
            Caption         =   "Enginering Change"
         End
      End
      Begin VB.Menu mnMenuPersediaan 
         Caption         =   "Inventory"
         Begin VB.Menu mnMInventory 
            Caption         =   "Master Inventory"
         End
         Begin VB.Menu mnItemCategories 
            Caption         =   "Inventory Categories"
         End
         Begin VB.Menu mnDescrip 
            Caption         =   "Master Outsourced Screen"
         End
         Begin VB.Menu mnTipeDes 
            Caption         =   "Outsourced Type"
         End
         Begin VB.Menu mnItemReference 
            Caption         =   "Inventory Reference"
         End
         Begin VB.Menu mnDescripRef 
            Caption         =   "Outsourced Referense"
         End
         Begin VB.Menu mnMutasi 
            Caption         =   "Stock Transfer"
         End
         Begin VB.Menu mnInvAdj 
            Caption         =   "Adjustment"
         End
      End
      Begin VB.Menu mnMenuProduksi 
         Caption         =   "Production Planning"
         Begin VB.Menu mnSchedule 
            Caption         =   "Master Production Schedule"
         End
         Begin VB.Menu mnManOrder 
            Caption         =   "Manufacturing Order"
         End
         Begin VB.Menu mnCapaPlan 
            Caption         =   "Capacity Planning"
         End
         Begin VB.Menu mnMrp 
            Caption         =   "MRP Generation"
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
         Begin VB.Menu mnBackFlushing 
            Caption         =   "BackFlushing"
         End
      End
   End
   Begin VB.Menu mnAkun 
      Caption         =   "Accounting"
      Begin VB.Menu mnMasAkun 
         Caption         =   "Data Master"
         Begin VB.Menu mnPerkiraan 
            Caption         =   "Master Perkiraan"
         End
         Begin VB.Menu mnKonfig 
            Caption         =   "Setup"
            Begin VB.Menu mnPeriode 
               Caption         =   "Period Setting"
            End
            Begin VB.Menu mnSetupAccount 
               Caption         =   "Configuration Account"
            End
         End
      End
      Begin VB.Menu mnKass 
         Caption         =   "Transaction"
         Begin VB.Menu mnBkmPiutang 
            Caption         =   "Pelunasan Piutang Karyawan"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnBkm 
            Caption         =   "Penerimaan Tunai"
         End
         Begin VB.Menu mnTukasKas 
            Caption         =   "Penukaran Setara Kas Ke Kas"
         End
         Begin VB.Menu mnBkk 
            Caption         =   "Pengeluaran Tunai"
         End
         Begin VB.Menu mnVoucher 
            Caption         =   "Pelunasan Hutang / Piutang"
         End
         Begin VB.Menu mnBkkPiutang 
            Caption         =   "Pengeluaran Piutang Ke Karyawan"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnReval 
         Caption         =   "Revaluation"
         Begin VB.Menu mnRevFinance 
            Caption         =   "Financial Series"
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
            Caption         =   "Memorial Umum"
         End
         Begin VB.Menu mnMemoJualbeli 
            Caption         =   "Memorial Pembelian/Penjualan"
         End
      End
      Begin VB.Menu mnClosed 
         Caption         =   "Period Closing"
         Begin VB.Menu mnClosing 
            Caption         =   "Closing"
         End
      End
   End
   Begin VB.Menu mnMenuAktiva 
      Caption         =   "Fixed Asset"
      Begin VB.Menu MnuFASetup 
         Caption         =   "Setup"
         Begin VB.Menu mnAktiva 
            Caption         =   "Master Asset"
         End
         Begin VB.Menu MnuFAFisCal 
            Caption         =   "Fiscal Calendar"
         End
         Begin VB.Menu MnuFAQuarter 
            Caption         =   "Quarter"
         End
         Begin VB.Menu MnuFABook 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuFAClass 
            Caption         =   "Class"
         End
         Begin VB.Menu MnuFAccGroup 
            Caption         =   "Account Group"
         End
         Begin VB.Menu MnuFAPosting 
            Caption         =   "Posting Account"
         End
         Begin VB.Menu MnuFANumber 
            Caption         =   "Numbering"
         End
         Begin VB.Menu MnuFAStruc 
            Caption         =   "Structure"
         End
         Begin VB.Menu MnuFALoc 
            Caption         =   "Location"
         End
         Begin VB.Menu MnuFAFisikLoc 
            Caption         =   "Physical Location"
         End
         Begin VB.Menu MnuFALease 
            Caption         =   "Lease Company"
         End
         Begin VB.Menu MnuFAInsu 
            Caption         =   "Insurance Class"
         End
      End
      Begin VB.Menu MnuFATrans 
         Caption         =   "Transaction"
         Begin VB.Menu mnBeliAktiva 
            Caption         =   "Purchase"
         End
         Begin VB.Menu mnJualAktiva 
            Caption         =   "Sale"
         End
         Begin VB.Menu MnuFAMaint 
            Caption         =   "Maintenance"
         End
         Begin VB.Menu MnuFARetire 
            Caption         =   "Retirement"
         End
      End
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
Attribute VB_Name = "MainMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim myMenu As New clsMenu
Dim mCalV, mCalS As Boolean
Dim mbMoving As Boolean

Private Sub MDIForm_Activate()
'myMenu.CreateMenu "MASTERORDER"
MainMenu.StatusBar1.Panels(4).Text = Format(Date, "dd MMMM yyyy")
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
'OpenMenu
Me.Caption = App.Comments
myMenu.CreateMenu "MASTERORDER"
SemeruTree1.Visible = False
Err.Clear
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim I As Integer
I = MessageBox("Anda Yakin Untuk Keluar Aplikasi?", "Keluar Aplikasi", msgYesNo)
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
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Set MainMenu = Nothing
End Sub

Private Sub mnAdmLap_Click()
If frmReport.Enabled = True Then frmReport.SetFocus
End Sub

Private Sub mnAktiva_Click()
If FrmMasterFixAssets.Enabled = True Then FrmMasterFixAssets.SetFocus
End Sub

Private Sub mnAr_Click()
If frmArTrans.Enabled = True Then frmArTrans.SetFocus
End Sub

Private Sub mnAtur_Click()
MainMenu.Arrange vbCascade
End Sub

Private Sub mnBeban_Click()
If frmBebanPembayaran.Enabled = True Then frmBebanPembayaran.SetFocus
End Sub

Private Sub mnClose_Click()
CloseAllForm
End Sub

'Private Sub mnBahan_Click()
'frmItemAsm.SetFocus
'End Sub

Private Sub mnBankPartner_Click()
If frmBankPartner.Enabled = True Then frmBankPartner.SetFocus
End Sub

Private Sub mnBeliAktiva_Click()
If FrmPembelianFixAssets.Enabled = True Then FrmPembelianFixAssets.SetFocus
End Sub

Private Sub mnBkk_Click()
If FrmPengeluaranBiaya.Enabled = True Then FrmPengeluaranBiaya.SetFocus
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
If FrmBom.Enabled = True Then FrmBom.SetFocus
End Sub

Private Sub mnBomMethode_Click()
If FrmBOMMethode.Enabled = True Then FrmBOMMethode.SetFocus
End Sub

Private Sub mnCalendar_Click()
If FrmCalendar.Enabled = True Then FrmCalendar.SetFocus
End Sub

Private Sub mnCloseSJt_Click()
'FrmCloseSJ.SetFocus
End Sub

Private Sub mnConnect_Click()
'FrmUserAktive.SetFocus
End Sub

Private Sub mnClosing_Click()
If frmValidasi.Enabled = True Then frmValidasi.SetFocus
End Sub

Private Sub mnCountPoint_Click()
If FrmStage.Enabled = True Then FrmStage.SetFocus
End Sub

Private Sub mnCusMaster_Click()
If frmPartner.Enabled = True Then frmPartner.SetFocus
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

Private Sub mnECC_Click()
If FrmEnginering.Enabled = True Then FrmEnginering.SetFocus
End Sub

Private Sub mnEdit_Click()
Call frmEmployess.mnEdit_Click
End Sub

Private Sub mnEditGroupMenu_Click()
'FrmPolicy.AddNode "EDIT"
End Sub

Private Sub mnExAccount_Click()
If FrmCurrencyAccount.Enabled = True Then FrmCurrencyAccount.SetFocus
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
If FrmCurrencySetup.Enabled = True Then FrmCurrencySetup.SetFocus
End Sub

Private Sub mnCusGudang_Click()
If FrmGudangCust.Enabled = True Then FrmGudangCust.SetFocus
End Sub

Private Sub mnGudang_Click()
If frmWareHouse.Enabled = True Then frmWareHouse.SetFocus
End Sub

Private Sub mnHapus_Click()
Call frmEmployess.mnHapus_Click
End Sub

Private Sub mnHideMenu_Click()
SemeruTree1.Visible = False
End Sub

Private Sub mnHlpApp_Click()
'FrmItemDescriptor.SetFocus
''FrmPlanned.SetFocus
'FrmMPS.SetFocus
'FrmUOM.SetFocus
End Sub

Private Sub mnHoris_Click()
MainMenu.Arrange vbTileHorizontal
End Sub

Private Sub mnHupiut_Click()
If frmVoucher.Enabled = True Then frmVoucher.SetFocus
End Sub

Private Sub mnInvAdj_Click()
If FrmInvAdj.Enabled = True Then FrmInvAdj.SetFocus
End Sub

Private Sub mnInvCard_Click()
If FrmItemData.Enabled = True Then FrmItemData.SetFocus
End Sub

Private Sub mnItemCategories_Click()
If FrmCategories.Enabled = True Then FrmCategories.SetFocus
End Sub

Private Sub mnItemReference_Click()
If FrmItemReference.Enabled = True Then FrmItemReference.SetFocus
End Sub

Private Sub mnJabat_Click()
Call frmEmployess.mnJabat_Click
End Sub

Private Sub mnJobCosting_Click()
If FrmBomCosting.Enabled = True Then FrmBomCosting.SetFocus
End Sub

Private Sub mnJualAktiva_Click()
If FrmPenjualanFixAssets.Enabled = True Then FrmPenjualanFixAssets.SetFocus
End Sub

Private Sub mnKelompok_Click()
If frmKelompok.Enabled = True Then frmKelompok.SetFocus
End Sub

Private Sub mnLain_Click()
'frmItemPrice.SetFocus
End Sub

Private Sub mnLaporan_Click()
If FrmKonfigurasiAccount.Enabled = True Then FrmKonfigurasiAccount.SetFocus
End Sub

Private Sub mnLisence_Click()
If frmAbout.Enabled = True Then frmAbout.SetFocus

End Sub

Private Sub mnLogin_Click()
SemeruTree1.Visible = False
CloseAllForm
frmLogin.Show vbModal
End Sub

Private Sub mnManOrder_Click()
If FrmWorkCenter.Enabled = True Then FrmWorkCenter.SetFocus
End Sub

Private Sub mnManStage_Click()
If FrmStage.Enabled = True Then FrmStage.SetFocus
End Sub

Private Sub mnManWC_Click()
If FrmWCTrans.Enabled = True Then FrmWCTrans.SetFocus
End Sub

Private Sub mnMemoJualbeli_Click()
If frmInvMemo.Enabled = True Then frmInvMemo.SetFocus
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

Private Sub mnPeriode_Click()
If FrmSetingPeriode.Enabled = True Then FrmSetingPeriode.SetFocus
End Sub

Private Sub mnPerkiraan_Click()
If FrmPerkiraan.Enabled = True Then FrmPerkiraan.SetFocus
End Sub

Private Sub MnPlanOrder_Click()
If FrmPlanned.Enabled = True Then FrmPlanned.SetFocus
End Sub

Private Sub mnPo_Click()
If frmSalesContract.Enabled = True Then frmSalesContract.SetFocus
End Sub

Private Sub mnRegional_Click()
If frmRegional.Enabled = True Then frmRegional.SetFocus
End Sub

Private Sub mnReturBeli_Click()
If FrmReturBeli.Enabled = True Then FrmReturBeli.SetFocus
End Sub

Private Sub mnReturJual_Click()
If FrmReturJual.Enabled = True Then FrmReturJual.SetFocus
End Sub

Private Sub mnRn_Click()
If frmReceiveNotes.Enabled = True Then frmReceiveNotes.SetFocus
End Sub

Private Sub mnRsc_Click()
If FrmRsc.Enabled = True Then FrmRsc.SetFocus
End Sub

Private Sub mnRscType_Click()
If FrmResources.Enabled = True Then FrmResources.SetFocus
End Sub

Private Sub mnSc_Click()
If FrmPurchasing.Enabled = True Then FrmPurchasing.SetFocus
End Sub

Private Sub mnSetupAccount_Click()
If FrmSetupAccount.Enabled = True Then FrmSetupAccount.SetFocus
End Sub

Private Sub MnSFCast_Click()
'If frmMasterSup.Enabled = True Then frmMasterSup.SetFocus
End Sub

'Private Sub mnSetJournal_Click()
'FrmConfigAccount.SetFocus
'End Sub

Private Sub mnShowMenu_Click()
SemeruTree1.Visible = True
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

Private Sub mnTipeBayar_Click()
If frmBebanPembayaran.Enabled = True Then frmBebanPembayaran.SetFocus
End Sub

Private Sub mnTipeDes_Click()
If FrmDescriptor.Enabled = True Then FrmDescriptor.SetFocus
End Sub

Private Sub mnTransport_Click()
If frmTransport.Enabled = True Then frmTransport.SetFocus
End Sub

Private Sub mnTTp_Click()
CloseAllForm
End Sub

Private Sub mnUpdateHarga_Click()
'frmItemPrice.SetFocus
End Sub

Private Sub mnTukasKas_Click()
If FrmPenukaranSetaraKas.Enabled = True Then FrmPenukaranSetaraKas.SetFocus
End Sub

Private Sub mnTypeCost_Click()
If FrmCostElement.Enabled = True Then FrmCostElement.SetFocus
End Sub

Private Sub mnUserArea_Click()
If UCase(MainMenu.StatusBar1.Panels(1).Text) = "SA" Or UCase(MainMenu.StatusBar1.Panels(1).Text) = "ADMINISTRATOR" Then
   SemeruTree1.Visible = False
   CloseAllForm
   If FrmPolicy.Enabled = True Then FrmPolicy.SetFocus
Else
   MessageBox "Anda tidak mempunyai hak untuk mengatur akses user.", "Peringatan", msgOkOnly
End If
End Sub

Private Sub mnValidasi_Click()
'FrmValidasi.SetFocus
End Sub

Private Sub mnVertical_Click()
MainMenu.Arrange vbTileVertical
End Sub

Private Sub mnVoucher_Click()
If frmVoucher.Enabled = True Then frmVoucher.SetFocus
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
If UCase(Node.Key) <> "MASTERPERKIRAAN" And UCase(Node.Key) <> "SETUPACCOUNT" Then
   If IsConfigReady = False Then
      MessageBox "Master Perkiraan Belum komplet.", "Peringatan", msgOkOnly
      Exit Sub
   End If
End If
Select Case UCase(Node.Key)
   'APPLICATION
   'MASTER
      Case "CURRSETUP": If FrmCurrencySetup.Enabled = True Then FrmCurrencySetup.SetFocus
      Case "EXCMAINTENANCE": If FrmCurrencyMaint.Enabled = True Then FrmCurrencyMaint.SetFocus
      Case "EXCACCOUNT": If FrmCurrencyAccount.Enabled = True Then FrmCurrencyAccount.SetFocus
      Case "MASTERGUDANG": If frmWareHouse.Enabled = True Then frmWareHouse.SetFocus
      Case "MASTERKELOMPOK": If frmKelompok.Enabled = True Then frmKelompok.SetFocus
      Case "INVCARD": If FrmItemData.Enabled = True Then FrmItemData.SetFocus
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
   
   'DISTRIBUTION
      Case "TRANSAKSISCAST": FrmSalesForecast.SetFocus
      Case "TRANSAKSIPO": If FrmPurchasing.Enabled = True Then FrmPurchasing.SetFocus
      Case "PLANORDER": If FrmPlanned.Enabled = True Then FrmPlanned.SetFocus
      Case "TRANSAKSISC": If frmSalesContract.Enabled = True Then frmSalesContract.SetFocus
      Case "TRANSAKSIRN": If frmReceiveNotes.Enabled = True Then frmReceiveNotes.SetFocus
      Case "TRANSAKSIAR": If frmArTrans.Enabled = True Then frmArTrans.SetFocus
      Case "TRANSAKSIDO":  If FrmDO.Enabled = True Then FrmDO.SetFocus
      Case "RETURBELI": If FrmReturBeli.Enabled = True Then FrmReturBeli.SetFocus
      Case "RETURJUAL": If FrmReturJual.Enabled = True Then FrmReturJual.SetFocus
   
   'PRODUCTION
      Case "DESCREF": If FrmItemDescriptor.Enabled = True Then FrmItemDescriptor.SetFocus
      Case "ASEMBLYA3": If FrmBom.Enabled = True Then FrmBom.SetFocus
      Case "ASEMBLYA2": If FrmWorkCenter.Enabled = True Then FrmWorkCenter.SetFocus
      Case "RESOURCESII": If FrmRsc.Enabled = True Then FrmRsc.SetFocus
      Case "WC": If FrmWCTrans.Enabled = True Then FrmWCTrans.SetFocus
      Case "MD":  If FrmManDescriptor.Enabled = True Then FrmManDescriptor.SetFocus
      Case "IR": If FrmItemReference.Enabled = True Then FrmItemReference.SetFocus
      Case "CALENDAR": If FrmCalendar.Enabled = True Then FrmCalendar.SetFocus
      Case "CATEGORIES":  If FrmCategories.Enabled = True Then FrmCategories.SetFocus
      Case "COST METHODE": If FrmCostElement.Enabled = True Then FrmCostElement.SetFocus
      Case "INVADJ1": If FrmBomCosting.Enabled = True Then FrmBomCosting.SetFocus
      Case "BOM METHODE": If FrmBOMMethode.Enabled = True Then FrmBOMMethode.SetFocus
      Case "DESCRIPTOR": If FrmDescriptor.Enabled = True Then FrmDescriptor.SetFocus
      Case "STAGE": If FrmStage.Enabled = True Then FrmStage.SetFocus
      Case "RESOURCES": If FrmResources.Enabled = True Then FrmResources.SetFocus
      Case "MUTASIGUDANG": If frmMutasiGudang.Enabled = True Then frmMutasiGudang.SetFocus
      Case "INVADJ": If FrmInvAdj.Enabled = True Then FrmInvAdj.SetFocus
      Case "BFINPUT": If FrmBackFlushInput.Enabled = True Then FrmBackFlushInput.SetFocus
      Case "BFOUTPUT": If frmBackFlushOutput.Enabled = True Then frmBackFlushOutput.SetFocus
      Case "MR": If FrmMatRequest.Enabled = True Then FrmMatRequest.SetFocus
      Case "MRPGEN": If FrmMRP.Enabled = True Then FrmMRP.Show vbModal
      Case "PLO": If FrmPlanned.Enabled = True Then FrmPlanned.SetFocus
      Case "ECC": If FrmEnginering.Enabled = True Then FrmEnginering.SetFocus
      Case "MPS": If FrmMPS.Enabled = True Then FrmMPS.SetFocus
   
   'ACCOUNTING
      Case "MASTERPERKIRAAN": If FrmPerkiraan.Enabled = True Then FrmPerkiraan.SetFocus
      Case "VOUCHERTRANSAKSI":  If frmVoucher.Enabled = True Then frmVoucher.SetFocus
      Case "TUNAIBIAYA": If FrmPengeluaranBiaya.Enabled = True Then FrmPengeluaranBiaya.SetFocus
      Case "BAYARTUNAILAIN": If FrmBKM.Enabled = True Then FrmBKM.SetFocus
      Case "VALIDASIJOURNAL": 'FrmValidasi.SetFocus
      Case "CLOSING": If frmValidasi.Enabled = True Then frmValidasi.SetFocus
      Case "SETUPACCOUNT": If FrmSetupAccount.Enabled = True Then FrmSetupAccount.SetFocus
      Case "KONFIGPERIODE": If FrmSetingPeriode.Enabled = True Then FrmSetingPeriode.SetFocus
      Case "DOUBLEENTRY":  If frmMemorial.Enabled = True Then frmMemorial.SetFocus
      Case "INVMEMO": If frmInvMemo.Enabled = True Then frmInvMemo.SetFocus
      Case "PENUKARAN": If FrmPenukaranSetaraKas.Enabled = True Then FrmPenukaranSetaraKas.SetFocus
   
   'FIXED ASSET
      Case "FAPURCHASE": If FrmPembelianFixAssets.Enabled = True Then FrmPembelianFixAssets.SetFocus
      Case "FASALES": If FrmPenjualanFixAssets.Enabled = True Then FrmPenjualanFixAssets.SetFocus
      Case "MASTERAKTIVA": If FrmMasterFixAssets.Enabled = True Then FrmMasterFixAssets.SetFocus
      Case "FISCAL": FrmSetingPeriode.SetFocus
      Case "QUARTER": FrmQuarter.SetFocus
      Case "BOOK": FrmBookSetup.SetFocus
      Case "CLASS": FrmClassSetup.SetFocus
      Case "ACCGROUP": FrmAccGroup.SetFocus
      Case "NUMBERING": FrmAssetsBook.SetFocus
      Case "RETIREMENT": FrmRetirementMaintenance.SetFocus
      Case "TRANSFER": FrmTransferMaintenance.SetFocus
      
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
Err.Clear
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'CloseAllForm
If SemeruTree1.Visible = False Then SemeruTree1.Visible = True
MainMenu.StatusBar1.Panels(3).Text = Button.Caption
Select Case Button.Index
       Case 1: myMenu.CreateMenu "MASTERORDER"
       Case 3: myMenu.CreateMenu "DISTRIBUSI"
       Case 5: myMenu.CreateMenu "PRODUKSI"
       Case 7: myMenu.CreateMenu "AKUNTING"
       Case 9: myMenu.CreateMenu "FIXEDASSET"
       Case 11: frmReport.SetFocus
       Case 13: myMenu.CreateMenu "KONFIGURASI"
End Select
End Sub

Private Sub OpenMenu()
MainMenu.Toolbar1.Buttons(1).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Master Data") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Master Data"), False))
MainMenu.Toolbar1.Buttons(2).Visible = MainMenu.Toolbar1.Buttons(1).Visible
MainMenu.Toolbar1.Buttons(3).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Distribution") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Distribution"), False))
MainMenu.Toolbar1.Buttons(4).Visible = MainMenu.Toolbar1.Buttons(3).Visible
MainMenu.Toolbar1.Buttons(5).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Produksi") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Production"), False))
MainMenu.Toolbar1.Buttons(6).Visible = MainMenu.Toolbar1.Buttons(5).Visible
MainMenu.Toolbar1.Buttons(7).Visible = CBool(IIf((GetSetting(App.EXEName, "Lisence Profile", "Akunting") <> ""), GetSetting(App.EXEName, "Lisence Profile", "Accounting"), False))
MainMenu.Toolbar1.Buttons(8).Visible = MainMenu.Toolbar1.Buttons(7).Visible
MainMenu.Toolbar1.Buttons(9).Visible = False
MainMenu.Toolbar1.Buttons(10).Visible = False
MainMenu.Toolbar1.Buttons(11).Visible = MainMenu.Toolbar1.Buttons(1).Visible + MainMenu.Toolbar1.Buttons(3).Visible + MainMenu.Toolbar1.Buttons(5).Visible + MainMenu.Toolbar1.Buttons(7).Visible
End Sub

Private Function SeekFormByTag(ByVal FormTag As String) As Boolean
Dim I As Integer
Dim Frm As Form
On Error GoTo Hell

For Each Frm In Forms
    If UCase(Frm.Tag) = UCase(FormTag) Then
       SeekFormByTag = True
       Frm.ZOrder (0)
    End If
Next
Set Frm = Nothing
Hell:
    Err.Clear
End Function
