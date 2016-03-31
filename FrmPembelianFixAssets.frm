VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPembelianFixAssets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pembelian Fixed Asset"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPembelianFixAssets.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Tag             =   "Asset Purchase"
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
      Height          =   6300
      Left            =   90
      ScaleHeight     =   6270
      ScaleWidth      =   10425
      TabIndex        =   17
      Top             =   15
      Width           =   10455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5550
         Left            =   120
         ScaleHeight     =   5520
         ScaleWidth      =   10095
         TabIndex        =   18
         Top             =   630
         Width           =   10125
         Begin VB.TextBox txtBox 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "DP"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   7185
            MaxLength       =   12
            TabIndex        =   13
            Tag             =   "ASM"
            Top             =   1425
            Width           =   2715
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "NoCheque"
            Height          =   315
            Index           =   1
            Left            =   7185
            MaxLength       =   15
            TabIndex        =   14
            Tag             =   "ASM"
            Top             =   1755
            Width           =   2715
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Bindings        =   "FrmPembelianFixAssets.frx":6852
            Height          =   1635
            Left            =   105
            TabIndex        =   15
            Top             =   2835
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   2884
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
            HeadLines       =   2
            RowHeight       =   15
            RowDividerStyle =   6
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "No Aktiva"
               Caption         =   "No Aktiva"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Nama Aktiva"
               Caption         =   "Nama Aktiva"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Aktiva Beli"
               Caption         =   "QTY"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Ppn"
               Caption         =   "PPN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Harga"
               Caption         =   "Harga"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   6090
            Picture         =   "FrmPembelianFixAssets.frx":6867
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2445
            Width           =   405
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   6090
            Picture         =   "FrmPembelianFixAssets.frx":6BF1
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2115
            Width           =   405
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   9495
            Picture         =   "FrmPembelianFixAssets.frx":6F7B
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   435
            Width           =   405
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   5310
            Picture         =   "FrmPembelianFixAssets.frx":7305
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   780
            Width           =   405
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Umur"
            Height          =   315
            Index           =   0
            Left            =   2190
            MaxLength       =   15
            TabIndex        =   4
            Tag             =   "ASM"
            Top             =   1440
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Tanggal"
            Height          =   315
            Left            =   2175
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   435
            Width           =   3525
            _ExtentX        =   6218
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
            CustomFormat    =   "dddd dd/MMMM/yyyy"
            Format          =   61145091
            CurrentDate     =   38272
         End
         Begin MSDataListLib.DataCombo cboRakit 
            DataField       =   "ID Group"
            Height          =   330
            Index           =   1
            Left            =   2190
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   1770
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Aktiva Group"
            BoundColumn     =   "ID Group"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   14
            Left            =   6165
            TabIndex        =   38
            Top             =   5115
            Width           =   960
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   7290
            TabIndex        =   37
            Top             =   5100
            Width           =   2625
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   6150
            X2              =   7560
            Y1              =   5340
            Y2              =   5340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   13
            Left            =   6165
            TabIndex        =   36
            Top             =   4845
            Width           =   330
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   7290
            TabIndex        =   35
            Top             =   4830
            Width           =   2625
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   6150
            X2              =   7560
            Y1              =   5070
            Y2              =   5070
         End
         Begin VB.Label lblTotalKas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7185
            TabIndex        =   34
            Top             =   1095
            Width           =   2715
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   5895
            X2              =   7575
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Kas/Bank"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   12
            Left            =   5895
            TabIndex        =   33
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uang Muka"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   11
            Left            =   5895
            TabIndex        =   32
            Top             =   1455
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Cek/BG"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   10
            Left            =   5895
            TabIndex        =   31
            Top             =   1785
            Width           =   885
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   5895
            X2              =   7575
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   5895
            X2              =   7575
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   6150
            X2              =   7560
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   390
            X2              =   2370
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   390
            X2              =   2370
            Y1              =   2415
            Y2              =   2415
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   390
            X2              =   2370
            Y1              =   2085
            Y2              =   2085
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   390
            X2              =   2370
            Y1              =   1740
            Y2              =   1740
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   390
            X2              =   2370
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   390
            X2              =   2370
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   390
            X2              =   2610
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelompok Depresiasi"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   9
            Left            =   420
            TabIndex        =   28
            Top             =   2490
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelompok Akum"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   8
            Left            =   420
            TabIndex        =   29
            Top             =   2160
            Width           =   1320
         End
         Begin VB.Label LblDep 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            DataField       =   "Depresiasi Aktiva"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   2190
            TabIndex        =   8
            Tag             =   "ASM"
            Top             =   2445
            Width           =   3870
         End
         Begin VB.Label LblDep 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            DataField       =   "Kelompok Akumulasi"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   2190
            TabIndex        =   6
            Tag             =   "ASM"
            Top             =   2115
            Width           =   3870
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   7290
            TabIndex        =   27
            Top             =   4560
            Width           =   2625
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   6165
            TabIndex        =   26
            Top             =   4575
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aktiva Group"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   5
            Left            =   420
            TabIndex        =   25
            Top             =   1830
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Umur                                                Bulan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   420
            TabIndex        =   24
            Top             =   1485
            Width           =   3750
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Kode Kas"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   5880
            TabIndex        =   10
            Tag             =   "ASM"
            Top             =   435
            Width           =   3585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kas/Bank"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   5895
            TabIndex        =   23
            Top             =   150
            Width           =   735
         End
         Begin VB.Label lblAlamatBank 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Kas"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5880
            TabIndex        =   12
            Tag             =   "ASM"
            Top             =   765
            Width           =   4020
         End
         Begin VB.Label lblFixAssets 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No Bukti"
            DataField       =   "Nama Perusahaan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   2190
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   1110
            Width           =   3525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Supplier"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   7
            Left            =   420
            TabIndex        =   22
            Top             =   1162
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Supplier"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   6
            Left            =   420
            TabIndex        =   21
            Top             =   832
            Width           =   1125
         End
         Begin VB.Label lblFixAssets 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No Bukti"
            DataField       =   "Kode Supplier"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   2190
            TabIndex        =   2
            Tag             =   "ASM"
            Top             =   780
            Width           =   3075
         End
         Begin VB.Label lblFixAssets 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
            DataField       =   "No FA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   2190
            TabIndex        =   0
            Tag             =   "ASM"
            Top             =   150
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   20
            Top             =   487
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   19
            Top             =   180
            Width           =   690
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   6510
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
      LimitRecordData =   "1"
   End
End
Attribute VB_Name = "FrmPembelianFixAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcGroup As New DBQuick
Private MyData As New clsTransaksi
Private mVarAdd As Boolean
Private RcPartner As New DBQuick

Private Sub cboRakit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Select Case DGPurchase.Col
       Case 3, 4:            HitungTotal
End Select
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mVarAdd = True Then
   Select Case DGPurchase.Col
          Case 2, 3, 4: DGPurchase.AllowUpdate = True
          Case Else: DGPurchase.AllowUpdate = False
   End Select
Else
   DGPurchase.AllowUpdate = False
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = Date
RcGroup.DBOpen "SELECT     NoAccount AS [Id Group], AccountName AS [Aktiva Group] FROM         GLAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'Aktiva Tetap Kantor' OR                      Type = N'Aktiva Tetap Produksi' OR                      Type = N'Aktiva Tetap Tak Berwujud') ORDER BY NoAccount", CNN, lckLockReadOnly
Set cboRakit(1).RowSource = RcGroup.DBRecordset
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmPembelianFixAssets
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = " SELECT     [TR Aktiva Tetap].[No FA], [TR Aktiva Tetap].PartnerID AS [Kode Supplier], PartnerDB.CompanyName AS [Nama Perusahaan],                       [TR Aktiva Tetap].DateTrans AS Tanggal, [TR Aktiva Tetap].Umur, [TR Aktiva Tetap].Unit, [TR Aktiva Tetap].Metode, [TR Aktiva Tetap].[Id Group],                       [TR Aktiva Tetap].Closed, [TR Aktiva Tetap].BankID AS [Kode Kas], GLAccount.AccountName AS [Nama Kas], [TR Aktiva Tetap].AccDep,                       [TR Aktiva Tetap].DepAktiva, [TR Aktiva Tetap].DP, [TR Aktiva Tetap].NoCheque, GLAccount_2.AccountName AS [Kelompok Akumulasi],                       GLAccount_1.AccountName AS [Depresiasi Aktiva]" & _
                    " FROM         [TR Aktiva Tetap] INNER JOIN                       PartnerDB ON [TR Aktiva Tetap].PartnerID = PartnerDB.PartnerID INNER JOIN                       GLAccount ON [TR Aktiva Tetap].BankID = GLAccount.NoAccount INNER JOIN                       GLAccount GLAccount_1 ON [TR Aktiva Tetap].AccDep = GLAccount_1.NoAccount INNER JOIN                       GLAccount GLAccount_2 ON [TR Aktiva Tetap].DepAktiva = GLAccount_2.NoAccount WHERE     ([TR Aktiva Tetap].Closed = 0) AND ([TR Aktiva Tetap].TypeTrans = 'FB') ORDER BY [TR Aktiva Tetap].DateTrans"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGroup.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
HiasForm Picture1, Me
CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPembelianFixAssets = Nothing
End Sub

Private Sub mCall_CallLinkForm()
Select Case mCall.FromTagActive
       Case "MASTER SUPPLIER": frmMasterSup.SetFocus
       Case "MASTER AKTIVA TETAP": FrmMasterFixAssets.SetFocus
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "MASTER SUPPLIER":
            With MyDDE.ActiveRecordset
                 .Fields("Kode Supplier") = mCall.GetFieldByName(0)
                 .Fields("Nama Perusahaan") = mCall.GetFieldByName(1)
            End With
       Case "MASTER KAS":
            With MyDDE
                 .GetFieldByName("Kode Kas") = mCall.GetFieldByName(0)
                 .GetFieldByName("Nama Kas") = mCall.GetFieldByName(1)
            End With
            lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
       Case "AKUMULASI DEPRESIASI":
            With MyDDE.ActiveRecordset
                 .Fields("aCCDEP") = mCall.GetFieldByName(0)
                 .Fields("Kelompok Akumulasi") = mCall.GetFieldByName(1)
            End With
       Case "AKTIVA TETAP":
            With MyDDE.ActiveRecordset
                 .Fields("depaktiva") = mCall.GetFieldByName(0)
                 .Fields("Depresiasi Aktiva") = mCall.GetFieldByName(1)
            End With
       
       Case "MASTER AKTIVA TETAP":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = 1
                 .Fields(3) = 0
                 .Fields("ppn") = 0
                 .Fields(4) = mCall.GetFieldByName("Kode Akun")
            End With
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            mVarAdd = True
            txtBox(0).Enabled = False
       Case tmbAddNew:
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            With MyDDE
                 .GetFieldByName("No FA") = MyData.PrepareIndex(tmbTransaksiBeliAktivaTetap, 5, "", TglIndex)
                 .GetFieldByName("Umur") = 1
                 .GetFieldByName("ppn") = 0
                 .GetFieldByName("DP") = 0
                 .GetFieldByName("NoCheque") = "-"
                 .GetFieldByName("Tanggal") = DTPicker1.Value
            End With
            txtBox(0).SetFocus
            mVarAdd = True
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(4) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & txtBox(0) & "') ")
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               TotalKas NoVoucher(1)
               mVarAdd = False
            End If
       Case tmbPrint:
            CallRPTReport "Pembelian Aktiva.Rpt", "Select * from [Pembelian Aktiva] Where [No Bukti]='" & lblFixAssets(0) & "'"
'       Case Else: 'mVarDataDc = False
End Select
cmdLInk(0).Enabled = mVarAdd
cmdLInk(1).Enabled = mVarAdd
cmdLInk(2).Enabled = mVarAdd
cmdLInk(3).Enabled = mVarAdd
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No FA")), MyDDE.GetFieldByName("No FA"), "XXXXX")
lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
HitungTotal
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGrid = True And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If CCur(txtBox(2)) > CCur(lblTotalKas) Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah kas tidak cukup untuk melakukan transaksi.", "Peringatan", msgOkOnly
                  Else
                     MyDDE.IsChildMemberReady = True
                  End If
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum Lengkap. Harap diisi dulu.", "Peringatan", msgOkOnly
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
            
       
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
           
       Case tmbCancel:
            
       Case tmbDetail:
            If NoVoucher(1) <> "" Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If MyDDE.ChildRecordset.Fields(4) = 0 Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly
                  Else
                     MyDDE.IsChildMemberReady = True
                     MyDDE.CancelTrans = False
                  End If
               Else
                  MyDDE.IsChildMemberReady = True
                  MyDDE.CancelTrans = False
               End If
            Else
                MessageBox "Data bank atau kas belum dipilih.", "Peringatan", msgOkOnly
                MyDDE.IsChildMemberReady = False
                MyDDE.CancelTrans = True
            End If
End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [TR Aktiva Tetap]" & _
                     " ( [No FA],   DateTrans,PartnerID , Umur,  [ID Group],BankID,TypeTrans,Periode,accdep,depaktiva,DP,noCheque)" & _
                     " VALUES  (N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & lblFixAssets(1) & "'," & CDbl(txtBox(0)) & ",N'" & cboRakit(1).BoundText & "',N'" & NoVoucher(1) & "','FB'," & mVarPeriode & ",N'" & MyDDE.GetFieldByName("Accdep") & "',N'" & MyDDE.GetFieldByName("DepAktiva") & "'," & CDbl(MyDDE.GetFieldByName("DP")) & ",'" & txtBox(1) & "')"
                     
    .PrepareUpdate = " UPDATE [TR Aktiva Tetap]" & _
                     " SET dp=" & CDbl(txtBox(0)) & ",NoCheque = '" & txtBox(1) & "' , DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), PartnerID = N'" & lblFixAssets(1) & "'," & _
                     " Umur = " & CDbl(txtBox(0)) & ",[ID Group] =N'" & cboRakit(1).BoundText & "' ,BankID=N'" & NoVoucher(1) & "', Accdep=N'" & MyDDE.GetFieldByName("Accdep") & "' , DepAktiva=N'" & MyDDE.GetFieldByName("DepAktiva") & "'" & _
                     " WHERE ([No FA] = N'" & lblFixAssets(0) & "')"
                     
    .PrepareDelete = " DELETE FROM [TR Aktiva Tetap] WHERE     ([No FA] = N'" & lblFixAssets(0) & "') "
End With
Err.Clear
End Sub

Private Sub SimpanDetail()
Dim mykey As String
Dim mNoAktiva As String
Dim mNoAcc As String
Dim StrPartic As String
Dim mVarPPn As Variant
Dim mVarTotalAktiva As Variant
Dim mVarHarga As Variant
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [DTR Aktiva Tetap] WHERE ([No FA] = N'" & lblFixAssets(0) & "')") = True Then
           .MoveFirst
           Do
             If .EOF = True Then Exit Do
             mNoAktiva = .Fields("No Aktiva")
             mNoAcc = .Fields("NoAccount")
             StrPartic = .Fields("Nama Aktiva")
             If .Fields("PPn") <> 0 Then
                mVarPPn = .Fields("PPn")
             Else
                mVarPPn = 1
             End If
             mVarHarga = .Fields("Harga")
             mVarTotalAktiva = (mVarHarga * (mVarPPn / 100)) + mVarHarga
             SendDataToServer (" INSERT INTO [DTR Aktiva Tetap]  ([No FA],NoAccount ,[No Aktiva], [Aktiva Beli], Harga,ppn) VALUES     (N'" & lblFixAssets(0) & "',N'" & .Fields("NoAccount") & "', N'" & .Fields("No Aktiva") & "', " & .Fields("Aktiva Beli") & ", " & .Fields("Harga") & "," & .Fields("PPn") & ")")
             .MoveNext
           Loop
           
'           If SendDataToServer("Delete from [Table Journal] where TransID='" & lblFixAssets(0) & "'") = True Then
'              mykey = IdxAuto
'           If SendDataToServer(" INSERT INTO [Table Journal]" & _
'                               " (JournalID, TransID,  NoAccount, PartnerID, Currency, DateTrans,  Periode, TypeTrans,Nourut,refNotes)" & _
'                               " VALUES     (N'" & mykey & "', N'" & lblFixAssets(0) & "',  N'" & NoVoucher(1) & "',N'" & lblFixAssets(1) & "',  N'IDR', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BKKAT','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "/") & "','Pembelian " & StrPartic & "')") = True Then
'              'Harga Perolehan
'              SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan) " & _
'                                " VALUES   (N'" & mykey & "', N'" & cboRakit(1).BoundText & "', N'" & mNoAktiva & "', " & CCur(LblAmount(0)) & ", 0,'Harga Beli " & StrPartic & "')")
'              'Ppn Masukan
'              SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan) " & _
'                                " VALUES   (N'" & mykey & "', N'" & CariTypeAccount(42) & "', N'xxx', " & CCur(LblAmount(1)) & ", 0,'Pajak Masukan " & StrPartic & "')")
'
'              'Pembayaran DP
'              If CDbl(txtBox(2)) <> 0 Then
'                 If CDbl(txtBox(2)) >= CDbl(LblAmount(2)) Then
'                    SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                      " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan) " & _
'                                      " VALUES   (N'" & mykey & "', N'" & NoVoucher(1) & "', N'" & txtBox(1) & "',0, " & CDbl(LblAmount(2)) & ",'Pembelian Tunai " & StrPartic & "')")
'
'                 Else
'                    SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                      " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan) " & _
'                                      " VALUES   (N'" & mykey & "', N'" & NoVoucher(1) & "', N'" & txtBox(1) & "',0, " & CDbl(txtBox(2)) & ",'Uang Muka Pembelian " & StrPartic & "')")
'
'                    SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                      " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan) " & _
'                                      " VALUES  (N'" & mykey & "', N'" & CariTypeAccount(57) & "', N'" & lblFixAssets(1) & "', 0, " & CDbl(LblAmount(2)) - CCur(txtBox(2)) & ",'Hutang Pembelian " & StrPartic & "')")
'
'                 End If
'              Else
'                 SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                   " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan) " & _
'                                   " VALUES   (N'" & mykey & "', N'" & CariTypeAccount(57) & "', N'" & lblFixAssets(1) & "', 0, " & CDbl(LblAmount(2)) & ",'Hutang Pembelian " & StrPartic & "')")
'              End If
'            End If
'           End If
           .MoveLast
        End If
     End If
End With
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen " SELECT     [DTR Aktiva Tetap].[No Aktiva], [Tabel Aktiva Tetap].[Nama Aktiva], [DTR Aktiva Tetap].[Aktiva Beli], [DTR Aktiva Tetap].Harga, [DTR Aktiva Tetap].NoAccount,[DTR Aktiva Tetap].Ppn FROM         [DTR Aktiva Tetap] INNER JOIN                       [Tabel Aktiva Tetap] ON [DTR Aktiva Tetap].[No Aktiva] = [Tabel Aktiva Tetap].[No Aktiva] WHERE     ([DTR Aktiva Tetap].[No FA] = N'" & ParamString & "')", CNN
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     PartnerID AS [Kode Supplier], CompanyName AS [Nama Perusahaan], Address AS Alamat, City AS Kota, PostalCode AS [Kode POS],                        Phone AS Telp FROM         PartnerDB WHERE     (PartnerType = N'SUPPLIER')", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     GLAccount.NoAccount as [Kode Kas], GLAccount.AccountName as [Nama Kas], ABS(SUM(ISNULL([Tabel Pembantu].CurrentDR" & PeriodeFilter & ", 0) + [Detail Journal].Debet)  - SUM(ISNULL([Tabel Pembantu].CurrentCR" & PeriodeFilter & ", 0) + [Detail Journal].Credit)) AS Saldo FROM         [Table Journal] INNER JOIN [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID INNER JOIN GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount LEFT OUTER JOIN [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     ([Table Journal].Periode = " & mVarPeriode & ") AND (GLAccount.Type = N'Kas' OR GLAccount.Type = N'Setara Kas') AND (GLAccount.[Group] = N'Detail List Account') GROUP BY GLAccount.NoAccount, GLAccount.AccountName", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun] FROM         GLAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'Akumulasi Penyusutan A.T.' OR                       Type = N'Akum. Amort. A.T. Tak Berwujud') ORDER BY NoAccount", CNN, lckLockReadOnly
       Case 3:
            RcPartner.DBOpen "SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun] FROM         GLAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'Biaya Overhead Amortisasi' OR                       Type = N'Biaya Adm. Penyusutan A.T. Kantor' OR                       Type = N'Biaya Overhead Penyusutan A.T. Produksi') ORDER BY NoAccount", CNN, lckLockReadOnly
       Case 4:
            RcPartner.DBOpen "SELECT     [Tabel Aktiva Tetap].[No Aktiva] AS [Kode Aktiva], [Tabel Aktiva Tetap].[Nama Aktiva] FROM         GLAccount INNER JOIN                       [Tabel Aktiva Tetap] ON GLAccount.NoAccount = [Tabel Aktiva Tetap].NoAccount WHERE     (GLAccount.NoAccount = N'" & cboRakit(1).BoundText & "')", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0:
                mCall.FromTagActive = "MASTER SUPPLIER"
                mCall.CaptionLink = "Supplier"
           Case 1:
                mCall.FromTagActive = "MASTER KAS"
                mCall.txtCari = NoVoucher(1)
           Case 2:
                mCall.FromTagActive = "AKUMULASI DEPRESIASI"
           Case 3:
                mCall.FromTagActive = "AKTIVA TETAP"
           Case 4:
                mCall.FromTagActive = "MASTER AKTIVA TETAP"
                mCall.CaptionLink = "Data Aktiva"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
'    If FindOwnRecordset(MyDDE.ChildRecordset, "[No Aktiva] = '" & mCall.GetFieldByName("No Aktiva") & "'") = True Then
'       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("No Aktiva") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'       CancelDetailTrans
'       DGPurchase.SetFocus
'    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   OpenPartner = True
End If
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_Change(Index As Integer)
If Index = 0 And mVarAdd = True Then
   If txtBox(0) = "" Or txtBox(0) = "0" Then txtBox(0) = "1"
End If
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub CancelDetailTrans()
If MyDDE.ChildRecordset.Recordcount <> 0 Then
  If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveNext
  If MyDDE.ChildRecordset.EOF And MyDDE.ChildRecordset.Recordcount > 0 Then MyDDE.ChildRecordset.MoveLast
End If
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "FB/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then ValidNum KeyAscii
End Sub

Private Sub HitungTotal()
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mVarTemp As Variant
Dim mTotal As Variant
Dim mPPn As Variant
Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotal = 0
mPPn = 0
With RcTotal.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            '2 QTY  3 Harga  5 ppn
            mVarTemp = (Avdata(2, I) * Avdata(3, I))
            mTotal = mTotal + mVarTemp
            If Avdata(5, I) = 0 Then
               mPPn = 0
            Else
               
               mPPn = mPPn + (mVarTemp * (Avdata(5, I) / 100))
            End If
        Next I
     End If
     LblAmount(0) = FormatNumber(mTotal, 0)
     LblAmount(1) = FormatNumber(Round(mPPn), 0)
     LblAmount(2) = FormatNumber(mTotal + mPPn, 0)
End With
Set Avdata = Nothing
Set mVarTemp = Nothing
End Sub

Private Function CekGrid() As Boolean
Dim RcGrd As New DBQuick
Set RcGrd.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
RcGrd.DBRecordset.Filter = "[Harga] <> 0"
With RcGrd.DBRecordset
     If .Recordcount <> 0 Then
        CekGrid = True
     Else
        CekGrid = False
     End If
End With
RcGrd.CloseDB
End Function

Private Function IdxAuto() As String
IdxAuto = MyData.PrepareIndex(tmbTransaksiBKKAT, 5, "", TglIndexDT)
End Function

Private Function TglIndexDT() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndexDT = "FB/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Function CariAkun(ByVal NamaForm As String, ByVal NamaValue As String) As String
Dim rcAkun As New DBQuick
rcAkun.DBOpen "SELECT [Daftar Configurasi].NoAccount FROM [Daftar Configurasi] INNER JOIN                       GLAccount ON [Daftar Configurasi].NoAccount = GLAccount.NoAccount WHERE     ([Daftar Configurasi].[Nama Form] = N'" & NamaForm & "') AND ([Daftar Configurasi].[Value Data] = N'" & NamaValue & "') GROUP BY [Daftar Configurasi].NoAccount", CNN, lckLockReadOnly
With rcAkun.DBRecordset
     If .Recordcount <> 0 Then
        CariAkun = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GLAccount.NoAccount, AccType.ID, GLAccount.AccountName FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function
