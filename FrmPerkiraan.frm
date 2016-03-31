VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPerkiraan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perkiraan"
   ClientHeight    =   6360
   ClientLeft      =   1665
   ClientTop       =   3225
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPerkiraan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11565
   Tag             =   "Chart of Account"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4020
      Left            =   5280
      TabIndex        =   15
      Top             =   1515
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   7091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15380335
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List Account"
      TabPicture(0)   =   "FrmPerkiraan.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "SaldoTrans"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtSaldo(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSaldo(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSaldo(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtSaldo(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Saldo Transaksi"
      TabPicture(1)   =   "FrmPerkiraan.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgBudget(0)"
      Tab(1).Control(1)=   "dgBudget(1)"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1170
         MaxLength       =   12
         TabIndex        =   20
         Top             =   1395
         Width           =   1665
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1170
         MaxLength       =   12
         TabIndex        =   19
         Top             =   1065
         Width           =   1665
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1170
         MaxLength       =   12
         TabIndex        =   18
         Top             =   735
         Width           =   1665
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1170
         MaxLength       =   12
         TabIndex        =   17
         Top             =   405
         Width           =   1665
      End
      Begin MSDataGridLib.DataGrid SaldoTrans 
         Height          =   3555
         Left            =   2925
         TabIndex        =   16
         Top             =   375
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   6271
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483633
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Periode"
            Caption         =   "Period"
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
            DataField       =   "Balance"
            Caption         =   "Balance"
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
            MarqueeStyle    =   3
            ScrollBars      =   0
            RecordSelectors =   0   'False
            BeginProperty Column00 
               DividerStyle    =   3
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   3
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgBudget 
         Height          =   3570
         Index           =   1
         Left            =   -71880
         TabIndex        =   21
         Top             =   390
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   6297
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483633
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Caption         =   "Budget Tahun Depan"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Periode"
            Caption         =   "Period"
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
            DataField       =   "Amount"
            Caption         =   "Amount"
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
            MarqueeStyle    =   3
            ScrollBars      =   0
            RecordSelectors =   0   'False
            BeginProperty Column00 
               DividerStyle    =   3
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   3
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgBudget 
         Height          =   3570
         Index           =   0
         Left            =   -74925
         TabIndex        =   22
         Top             =   390
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   6297
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483633
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Caption         =   "Budget Tahun Ini"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Periode"
            Caption         =   "Period"
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
            DataField       =   "Amount"
            Caption         =   "Amount"
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
            MarqueeStyle    =   3
            ScrollBars      =   0
            RecordSelectors =   0   'False
            BeginProperty Column00 
               DividerStyle    =   3
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   75
         X2              =   1995
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   75
         X2              =   1995
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   75
         X2              =   1995
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   90
         X2              =   2010
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Akhir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   26
         Top             =   1425
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kredit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   25
         Top             =   1095
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Awal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   435
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5790
      Left            =   0
      ScaleHeight     =   5790
      ScaleWidth      =   11565
      TabIndex        =   8
      Top             =   0
      Width           =   11565
      Begin MSComctlLib.ImageList imgAccount 
         Left            =   3555
         Top             =   1905
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerkiraan.frx":688A
               Key             =   "ON TOP"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerkiraan.frx":D0EC
               Key             =   "SUB ACCOUNT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerkiraan.frx":1394E
               Key             =   "LIST ACCOUNT"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerkiraan.frx":1A1B0
               Key             =   "DETAIL LIST ACCOUNT"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox mskLedger 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         DataField       =   "NoAccount"
         DataSource      =   "DataMaster"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6825
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   430
         Width           =   1860
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "type"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "DataMaster"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   6825
         MaxLength       =   200
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1090
         Width           =   3510
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   10335
         Picture         =   "FrmPerkiraan.frx":20A12
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1110
         Width           =   315
      End
      Begin VB.CheckBox chkActive 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Non Active"
         DataField       =   "Status"
         Enabled         =   0   'False
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
         Height          =   210
         Left            =   9405
         TabIndex        =   7
         Top             =   482
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtLedger 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         DataField       =   "AccountName"
         DataSource      =   "DataMaster"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   6825
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   760
         Width           =   3840
      End
      Begin MSDataListLib.DataCombo cboType 
         DataField       =   "Group"
         Height          =   315
         Index           =   0
         Left            =   6825
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   90
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Level"
         BoundColumn     =   "Group"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SemeruDC.SemeruTree MenuAccount 
         Height          =   5370
         Left            =   105
         TabIndex        =   1
         Top             =   120
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   9472
         BackColorTree   =   7159830
         BackColorBackground=   16643562
      End
      Begin VB.CheckBox chkDebet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Debet"
         DataField       =   "Default"
         Enabled         =   0   'False
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
         Height          =   210
         Left            =   9240
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblKodePerkiraan 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Perkiraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5370
         TabIndex        =   13
         Top             =   490
         Width           =   1290
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5355
         X2              =   7275
         Y1              =   730
         Y2              =   730
      End
      Begin VB.Label LBLIndexGroup 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Perkiraan"
         Height          =   195
         Left            =   8820
         TabIndex        =   12
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perkiraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5370
         TabIndex        =   11
         Top             =   820
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Perkiraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   5370
         TabIndex        =   10
         Top             =   150
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Type Perkiraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   5370
         TabIndex        =   9
         Top             =   1158
         Width           =   1290
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   5355
         X2              =   7275
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5370
         X2              =   7290
         Y1              =   1060
         Y2              =   1060
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5415
         X2              =   7335
         Y1              =   1405
         Y2              =   1405
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPerkiraan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVarnode As Nodes
Private mVarNodeNode As Node
Private mAccInduk As String
Private mVarIndex As String
Private mLastGroup As String
Private RcType As New DBQuick
Private RcGroup As New DBQuick
Private rcPrd As New Recordset
Private rcti As New Recordset
Private rcTD As New Recordset
Private rsAccType As New DBQuick

Public Enum IndexAccount
    sGroupAccount = 1
    sListAccount = 2
    sSubAccount = 3
    sDetailListAccount = 4
End Enum

Private Type GroupMaxAcc
        idGroupAcc As String
        idDetailAcc As String
        idSubDetailAcc As String
        idListAcc As String
        idDetailListAcc As String
        idIndukAcc As String
End Type

Private Type LenGroupMaxAcc
        lenidGroupAcc As Byte
        lenidDetailAcc As Byte
        lenidSubDetailAcc As Byte
        lenidListAcc As Byte
        lenidDetailListAcc As Byte
        lenPreGroupAcc As String
        lenPreDetailAcc As String
        lenPreSubDetailAcc As String
        lenPreListAcc As String
        lenPreDetailListAcc As String
End Type

Private mIndex As GroupMaxAcc
Private mLenIndex As LenGroupMaxAcc
Private mAdd, MEdit As Boolean
Private mOpen, mLokCbo As Boolean
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim IDNYa As String
Public MyAcc As clsMaster

Private Sub cmdLink_Click(Index As Integer)
Set mCall = New frmCaller
    rsAccType.DBOpen " SELECT ID, Tipe From AccType Where (status = 1) ORDER BY Tipe", CNN, lckLockReadOnly
    If rsAccType.Recordcount <> 0 Then
        mCall.FromTagActive = "Tipe Rekening"
        Set mCall.FormData = rsAccType.DBRecordset
        mCall.LookUp Me
    Else
        MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Kontrol Entry Rekening", msgOkOnly, msgExclamation
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Rc As New DBQuick
MyDDE.SetPermissions = aksess.MayDo("Daftar Perkiraan") 'Set Akses Tombol
SSTab1.Tab = 0
mAccInduk = "ON TOP"
mIndex.idIndukAcc = mAccInduk
Rc.DBOpen "SELECT [Length Per Account], Prefix FROM [Account Setup] ORDER BY [No Index]", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Do
            If .EOF Then Exit Do
            Select Case .AbsolutePosition
                   Case 1:
                        mLenIndex.lenidGroupAcc = .Fields(0)
                        mLenIndex.lenPreGroupAcc = .Fields(1)
'                   Case 2:
'                        mLenIndex.lenidDetailAcc = .Fields(0)
'                        mLenIndex.lenPreDetailAcc = .Fields(1)
                   Case 2:
                        mLenIndex.lenidSubDetailAcc = .Fields(0)
                        mLenIndex.lenPreSubDetailAcc = .Fields(1)
                   Case 3:
                        mLenIndex.lenidListAcc = .Fields(0)
                        mLenIndex.lenPreListAcc = .Fields(1)
                   Case 4:
                        mLenIndex.lenidDetailListAcc = .Fields(0)
                        mLenIndex.lenPreDetailListAcc = .Fields(1)
            End Select
            .MoveNext
        Loop
     End If
End With

RcGroup.DBOpen "SELECT [Index Group] AS [Group], [Group Name] AS [Level] " & _
            " FROM [Account Setup] ORDER BY [No Index]", CNN, lckLockReadOnly
cboType(0).ListField = "Level"
Set cboType(0).RowSource = RcGroup.DBRecordset
cboType(0).MatchEntry = dblExtendedMatching
cboType(0).Text = RcGroup.DBRecordset.Fields(1)

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me

SSTab1.BackColor = Picture2.BackColor
chkActive.BackColor = Picture2.BackColor
chkDebet.BackColor = Picture2.BackColor
MenuAccount.BackColorBackground = Picture2.BackColor

'RcType.DBOpen " SELECT AccType.Tipe,AccType.[Id] FROM AccType ORDER BY AccType.Tipe; ", CNN, lckLockReadOnly
'Set txtBox(13).DataSource = RcType.DBRecordset
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmPerkiraan
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "Select * from [GLAccount] order by NoAccount"
End With
LoadTree
Set mVarNodeNode = MenuAccount.MenuTreeView.Nodes(1)
mVarNodeNode.Selected = True
MenuAccount.MenuTreeView.SelectedItem.EnsureVisible
Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CompareAccount
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      CompareAccount
      Cancel = False
      MyDDE.ClearRecordset
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPerkiraan = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
    txtBox(13).Text = mCall.GetFieldByName("Tipe")
    IDNYa = mCall.GetFieldByName("ID")
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
            
            cboType(0).Text = RcGroup.DBRecordset.Fields(1)
            RcGroup.DBRecordset.AbsolutePosition = cboType(0).SelectedItem
            Call cboType_Click(0, 0)
            MyDDE.GetFieldByName("Status") = False
            MyDDE.GetFieldByName("Default") = False
            MenuAccount.MenuTreeView.Enabled = False
            
            txtLedger(0).SetFocus
            mskLedger.Enabled = False
            cmdLink(0).Enabled = True
       Case tmbEdit:
            mAdd = True
            MEdit = True
            MenuAccount.MenuTreeView.Enabled = False
            mskLedger.Enabled = False
            cboType(0).Enabled = False
            cmdLink(0).Enabled = True
           ' mIndex.idIndukAcc = MyDDE.GetFieldByName("GroupAccount")
            'LoadTree
       Case tmbCancel:
            mAdd = False
            MEdit = False
            MenuAccount.MenuTreeView.Enabled = True
            mLastGroup = ""
       Case tmbDelete:
            mAdd = False
            MEdit = False
            NodeAdo "delete"
            MenuAccount.MenuTreeView.Enabled = True
            mLastGroup = ""
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               MenuAccount.MenuTreeView.Enabled = True
               If MEdit = True Then
                  NodeAdo "edit"
               Else
                  NodeAdo "new"
               End If
               MenuAccount.SetFocus
               mAdd = False
               MEdit = False
               cmdLink(0).Enabled = False
               mLastGroup = ""
            End If
       Case tmbPrint:
            CallRPTReport "Laporan Perkiraan.rpt"
            
       Case Else:
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
LoadPeriode
OpenSaldo IIf(Not IsNull(MyDDE.GetFieldByName("NoAccount")), MyDDE.GetFieldByName("NoAccount"), "xxxxx")

IDNYa = IIf(IsNull(MyDDE.GetFieldByName("id")), IDNYa, MyDDE.GetFieldByName("id"))
'IDNYa = MyDDE.GetFieldByName("id")
'Debug.Print IDNYa
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbAddNew:
            
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
                mIndex.idIndukAcc = mskLedger
            Else
                mIndex.idIndukAcc = "ON TOP"
            End If
            
            OpenListAccount
       Case tmbEdit:
            If MyDDE.CheckEmptyControl = False Then
               If mVarIndex = "" Then
                  MyDDE.CancelTrans = True
                  MessageBox "Perkiraan Datatree belum dipilih.Silahkan anda pilih dulu.", "Peringatan", msgOkOnly, msgCrtical
               Else
                  MyDDE.CancelTrans = False
               End If
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
                If mDel.CekDelete(mskLedger.Text, reDelMasterAccount) = False Then
                    MyDDE.IsChildMemberReady = True
                    If HirarkiDelete(mskLedger, LBLIndexGroup.Caption) Then
                        PrepareQuery
                    Else
                        MyDDE.CancelTrans = True
                    End If
                Else
                    MyDDE.CancelTrans = True
                    MessageBox "Kode akun (" & mskLedger.Text & ") sedang digunakan." & vbCrLf & "Tidak Bisa DiHapus.", "Account Control", msgOkOnly, msgExclamation
                    MyDDE.IsChildMemberReady = False
                End If
            Else
                MyDDE.IsChildMemberReady = False
            End If
       Case tmbCancel: mAdd = False
         
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               mAdd = False
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub
Private Function HirarkiDelete(sKodeAcc As String, sGroup As String) As Boolean
Dim strSQL As String
Select Case sGroup
    Case "Group Account"
        strSQL = "DELETE FROM GLAccount WHERE ([Group] = N'Sub Account') AND " & _
            " (LEFT(NoAccount, 1) = '" & Left(sKodeAcc, 1) & "') OR ([Group] = N'List Account') AND " & _
            " (LEFT(NoAccount, 1) = '" & Left(sKodeAcc, 1) & "') OR ([Group] = N'Detail List Account') AND " & _
            " (LEFT(NoAccount, 1) = '" & Left(sKodeAcc, 1) & "')"
    Case "Sub Account"
        strSQL = "DELETE FROM GLAccount WHERE (LEFT(NoAccount, 2) = '" & Left(sKodeAcc, 2) & "')"
        strSQL = "DELETE FROM GLAccount WHERE (LEFT(NoAccount, 2) = '" & Left(sKodeAcc, 2) & "') AND " & _
            " ([Group] = N'List Account') OR (LEFT(NoAccount, 2) = '" & Left(sKodeAcc, 2) & "') AND ([Group] = N'Detail List Account')"
    Case "List Account"
        strSQL = "DELETE FROM GLAccount WHERE (LEFT(NoAccount, 3) = '" & Left(sKodeAcc, 3) & "') AND ([Group] = N'Detail List Account')"
End Select
If sGroup = "Detail List Account" Then
    HirarkiDelete = True
Else
    If MsgBox("Kode akun dan sub akun semua akun akan dihapus", vbExclamation + vbYesNo, "Control Account") = vbYes Then
            HirarkiDelete = True
            SendDataToServer strSQL
    Else
        HirarkiDelete = False
    End If
End If
End Function
Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO GLAccount" & _
                     " (ID,NoAccount, Type, [Group], AccountName, GroupAccount, Status,[Default])" & _
                     " VALUES (" & IDNYa & ",N'" & mskLedger & "', N'" & txtBox(13).Text & "', N'" & cboType(0).BoundText & "', N'" & txtLedger(0) & "', N'" & mIndex.idIndukAcc & "', " & chkActive.Value & "," & chkDebet.Value & ")"
    .PrepareUpdate = " UPDATE GLAccount" & _
                     " SET [ID]=" & IDNYa & ",Type = N'" & txtBox(13).Text & "', [Group] = N'" & cboType(0).BoundText & "', AccountName = N'" & txtLedger(0) & "', Status =  " & chkActive.Value & _
                     " ,[Default] = " & chkDebet.Value & " WHERE  (NoAccount = N'" & mskLedger.Text & "')"
    .PrepareDelete = " DELETE FROM [GLAccount] WHERE   (NoAccount = N'" & ValidString(mskLedger) & "') "
End With
Err.Clear
End Sub

Private Sub LoadTree()
'On Error Resume Next
Dim rcNode As New DBQuick
Dim I As Long
Dim Avdata As Variant
rcNode.DBOpen "Select * from [GLAccount] order by NoAccount", CNN, lckLockReadOnly

Set mVarnode = Nothing
Set MenuAccount.MenuTreeView.ImageList = imgAccount
With rcNode.DBRecordset
     If .Recordcount <> 0 Then
        
        mAccInduk = IIf(Not IsNull(.Fields("GroupAccount")), .Fields("GroupAccount"), "")
        If mAccInduk = "ON TOP" Then
           mAccInduk = IIf(Not IsNull(.Fields("NoAccount")), .Fields("NoAccount"), "")
        Else
        End If
        Set mVarnode = MenuAccount.MenuTreeView.Nodes
        mVarnode.Clear
        MenuAccount.MenuTreeView.Nodes.Clear
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            Select Case UCase(Avdata(2, I))
                Case "GROUP ACCOUNT":
                     If Avdata(4, I) = "ON TOP" Then
                        MenuAccount.NodeAdd , , "[" & Avdata(0, I) & "]", Left(Avdata(0, I), mLenIndex.lenidGroupAcc) & " - " & Avdata(3, I), "ON TOP", , , True, , True, True, , &HFCF1ED, &H6D4016
                     Else
                        MenuAccount.NodeAdd Avdata(4, I), tvwChild, "[" & Avdata(0, I) & "]", Left(Avdata(0, I), mLenIndex.lenidGroupAcc) & " - " & Avdata(3, I), "ON TOP", , , True, , True, True, , &HFCF1ED, &H6D4016
                     End If
                Case "SUB ACCOUNT":
                    MenuAccount.NodeAdd "[" & Avdata(4, I) & "]", tvwChild, "[" & Avdata(0, I) & "]", Mid(Avdata(0, I), mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreGroupAcc) + 1, mLenIndex.lenidSubDetailAcc) & " - " & Avdata(3, I), "SUB ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
                
                Case "LIST ACCOUNT"
'                    Debug.Print Mid(Avdata(0, i), mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreGroupAcc) + mLenIndex.lenidSubDetailAcc + Len(mLenIndex.lenPreSubDetailAcc) + 1, mLenIndex.lenidListAcc) & " - " & Avdata(3, i)
'                    Debug.Print Mid(Avdata(0, i), mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreGroupAcc) + mLenIndex.lenidSubDetailAcc + Len(mLenIndex.lenPreSubDetailAcc) + 1, mLenIndex.lenidListAcc) & " - " & Avdata(3, i)
                      MenuAccount.NodeAdd "[" & Avdata(4, I) & "]", tvwChild, "[" & Avdata(0, I) & "]", Mid(Avdata(0, I), mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreGroupAcc) + mLenIndex.lenidSubDetailAcc + Len(mLenIndex.lenPreSubDetailAcc) + 1, mLenIndex.lenidListAcc) & " - " & Avdata(3, I), "LIST ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
                Case "DETAIL LIST ACCOUNT":
                      MenuAccount.NodeAdd "[" & Avdata(4, I) & "]", tvwChild, "[" & Avdata(0, I) & "]", Right(Avdata(0, I), mLenIndex.lenidDetailListAcc) & " - " & Avdata(3, I), "DETAIL LIST ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
            End Select
        Next I
     Else
     End If
End With

rcNode.CloseDB
Set Avdata = Nothing
Err.Clear
End Sub

Private Function KasiZero(ByVal SourceDigiteSam As String, ByVal PirangDigitSam As Byte) As String
SourceDigiteSam = Replace(Trim(SourceDigiteSam), " ", "")
If Len(Trim(Str(SourceDigiteSam))) > PirangDigitSam Then
   KasiZero = Trim(Str(SourceDigiteSam))
Else
   Select Case PirangDigitSam - Len(Trim(Str(SourceDigiteSam)))
          Case 1: KasiZero = Trim(Str(SourceDigiteSam)) & "0"
          Case 2: KasiZero = Trim(Str(SourceDigiteSam)) & "00"
          Case 3: KasiZero = Trim(Str(SourceDigiteSam)) & "000"
          Case 4: KasiZero = Trim(Str(SourceDigiteSam)) & "0000"
          Case 5: KasiZero = Trim(Str(SourceDigiteSam)) & "00000"
          Case 6: KasiZero = Trim(Str(SourceDigiteSam)) & "000000"
          Case 7: KasiZero = Trim(Str(SourceDigiteSam)) & "0000000"
          Case 8: KasiZero = Trim(Str(SourceDigiteSam)) & "00000000"
          Case 9: KasiZero = Trim(Str(SourceDigiteSam)) & "000000000"
          Case 10: KasiZero = Trim(Str(SourceDigiteSam)) & "0000000000"
          Case 11: KasiZero = Trim(Str(SourceDigiteSam)) & "00000000000"
          Case 12: KasiZero = Trim(Str(SourceDigiteSam)) & "000000000000"
          Case 13: KasiZero = Trim(Str(SourceDigiteSam)) & "0000000000000"
          Case 14: KasiZero = Trim(Str(SourceDigiteSam)) & "00000000000000"
          Case 15: KasiZero = Trim(Str(SourceDigiteSam)) & "000000000000000"
          Case Else: KasiZero = Trim(Str(SourceDigiteSam))
   End Select
End If
End Function

Private Sub MenuAccount_NodeClick(ByVal Node As MSComctlLib.Node)
MyDDE.FindStringData "NoAccount ='" & Replace(Mid(Node.Key, 2, Len(Node.Key)), "]", "") & "'"
LBLIndexGroup.Caption = MyDDE.GetFieldByName("Group")
Select Case MyDDE.GetFieldByName("Group")
       Case "Group Account":
            mAccInduk = Replace(Mid(Node.Key, 2, Len(Node.Key)), "]", "")
       Case "Detail Account", "Sub Account", "List Account":
            mAccInduk = MyDDE.GetFieldByName("NoAccount")
       Case "Detail List Account":
            mAccInduk = Replace(Mid(Node.Parent.Key, 2, Len(Node.Parent.Key)), "]", "")
       Case Else: mAccInduk = "ON TOP"
End Select
Set mVarNodeNode = Node
mVarIndex = Trim(Str(Node.Index))
Err.Clear
End Sub

Private Sub cboType_Click(Index As Integer, Area As Integer)
If Index = 0 Then
   If mAdd = True Then
        'mLastGroup = ""
        If mLastGroup = "" Then mLastGroup = cboType(0).BoundText
        'mLastGroup = RcGroup.Fields(0)
        Select Case UCase(mLastGroup)
               Case "GROUP ACCOUNT":
                    'If mAccInduk <> "ON TOP" Then mIndex.idIndukAcc = "ON TOP"
'               Case "DETAIL ACCOUNT":
'                    If UCase(cboType(0).BoundText) <> "GROUP ACCOUNT" Then RcGroup.AbsolutePosition = 2
               Case "SUB ACCOUNT":
                    If UCase(cboType(0).BoundText) <> "GROUP ACCOUNT" Then RcGroup.AbsolutePosition = 2
               Case "LIST ACCOUNT":
                    If UCase(cboType(0).BoundText) <> "GROUP ACCOUNT" Then RcGroup.AbsolutePosition = 3
               Case "DETAIL LIST ACCOUNT":
                    If UCase(cboType(0).BoundText) <> "GROUP ACCOUNT" Then RcGroup.AbsolutePosition = 4
        End Select
        If UCase(cboType(0).BoundText) <> "GROUP ACCOUNT" Then cboType(0).Text = RcGroup.Fields(1)
        Select Case UCase(cboType(0).BoundText)
               Case "GROUP ACCOUNT":
                    If mAccInduk <> "ON TOP" Then mIndex.idIndukAcc = "ON TOP"
'               Case "DETAIL ACCOUNT":
'                    mIndex.idIndukAcc = mAccInduk
               Case "SUB ACCOUNT":
                    mIndex.idIndukAcc = mAccInduk
               Case "LIST ACCOUNT":
                    mIndex.idIndukAcc = mAccInduk
               Case "DETAIL LIST ACCOUNT":
                    mIndex.idIndukAcc = mAccInduk
        End Select
        If mAccInduk <> "" Then
           mskLedger = AutoIndexAcc(cboType(0).BoundText, mAccInduk)
        Else
           MessageBox "Data induk belum dipilih,maka default akun yang aktif menjadi Group Account"
        End If
   End If
End If
End Sub

Private Function AutoIndexAcc(ByVal GroupAcc As String, ByVal NoAccountAcc As String) As String
Dim Rckode As New DBQuick
Dim mVarTotalDigit As Long
Dim mIndexAuto As Long
Dim GroupAcct As Integer
Dim DetailAcct As Integer
Dim SubDetailAcct As Integer
Dim ListAcct As Integer
Dim DListAcct As Integer
Dim mTampung As String

If NoAccountAcc = "ON TOP" Then NoAccountAcc = KasiZero("0", mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + mLenIndex.lenidListAcc + mLenIndex.lenidDetailListAcc)
mVarTotalDigit = mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + mLenIndex.lenidListAcc + mLenIndex.lenidDetailListAcc

GroupAcct = 0
SubDetailAcct = mLenIndex.lenidGroupAcc + Val(Len(mLenIndex.lenPreGroupAcc)) + Len(mLenIndex.lenidSubDetailAcc)
ListAcct = mLenIndex.lenidGroupAcc + Val(Len(mLenIndex.lenPreGroupAcc)) + mLenIndex.lenidSubDetailAcc + Val(Len(mLenIndex.lenPreSubDetailAcc)) + mLenIndex.lenidListAcc + Val(Len(mLenIndex.lenPreListAcc))
DListAcct = mLenIndex.lenidDetailListAcc
Select Case UCase(GroupAcc)
       Case "GROUP ACCOUNT":
            Rckode.DBOpen "SELECT  MAX(LEFT(NoAccount, " & mLenIndex.lenidGroupAcc & ")) AS MaxNom FROM GLAccount ", CNN, lckLockReadOnly
'       Case "DETAIL ACCOUNT":
'            Rckode.DBOpen "SELECT  MAX(SUBSTRING(NoAccount," & DetailAcct & ", " & mLenIndex.lenidDetailAcc & ")) AS MaxNom FROM GLAccount Where left(noaccount," & mLenIndex.lenidGroupAcc & ")=N'" & Left(NoAccountAcc, mLenIndex.lenidGroupAcc) & "' and [Group]=N'" & GroupAcc & "'", CNN, lckLockReadOnly
       Case "SUB ACCOUNT":
            Rckode.DBOpen "SELECT  MAX(SUBSTRING(NoAccount," & SubDetailAcct & ", " & mLenIndex.lenidSubDetailAcc & ")) AS MaxNom FROM GLAccount where  left(noaccount," & mLenIndex.lenidGroupAcc & ")=N'" & Left(NoAccountAcc, mLenIndex.lenidGroupAcc) & "' And [Group]=N'" & GroupAcc & "'", CNN, lckLockReadOnly
       Case "LIST ACCOUNT":
            Rckode.DBOpen "SELECT  MAX(SUBSTRING(NoAccount," & ListAcct & ", " & mLenIndex.lenidListAcc & ")) AS MaxNom FROM GLAccount where  left(noaccount," & mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc & ")=N'" & Left(NoAccountAcc, mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc) & "' And [Group]=N'" & GroupAcc & "'", CNN, lckLockReadOnly
       Case "DETAIL LIST ACCOUNT":
            Rckode.DBOpen "SELECT  MAX(right(NoAccount," & mLenIndex.lenidDetailListAcc & ")) AS MaxNom FROM GLAccount where   left(noaccount," & mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + mLenIndex.lenidListAcc + Len(mLenIndex.lenPreGroupAcc) + Len(mLenIndex.lenPreSubDetailAcc) + Len(mLenIndex.lenPreListAcc) & ")=N'" & Left(NoAccountAcc, mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + mLenIndex.lenidListAcc + Len(mLenIndex.lenPreGroupAcc) + Len(mLenIndex.lenPreSubDetailAcc) + Len(mLenIndex.lenPreListAcc)) & "' And  [Group]=N'" & GroupAcc & "'", CNN, lckLockReadOnly
End Select
With Rckode.DBRecordset
     If .Recordcount <> 0 Then
         mIndexAuto = Val(IIf(Not IsNull(.Fields(0)), .Fields(0), 0))
     Else
        If NoAccountAcc = "ON TOP" Then NoAccountAcc = KasiZero("0", mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + mLenIndex.lenidListAcc + mLenIndex.lenidDetailListAcc)
        GroupAcct = Left(NoAccountAcc, mLenIndex.lenidGroupAcc)
'        DetailAcct = Val(Mid(NoAccountAcc, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreGroupAcc) + 1, mLenIndex.lenidDetailAcc))
        SubDetailAcct = Val(Mid(NoAccountAcc, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreGroupAcc), mLenIndex.lenidSubDetailAcc))
        ListAcct = Val(Mid(NoAccountAcc, mLenIndex.lenidGroupAcc + mLenIndex.lenidListAcc + mLenIndex.lenidSubDetailAcc, mLenIndex.lenidListAcc))
        DListAcct = Val(Right(NoAccountAcc, mLenIndex.lenidDetailListAcc))
        mIndexAuto = 0
     End If
     mIndex.idGroupAcc = Left(NoAccountAcc, mLenIndex.lenidGroupAcc)
'     mIndex.idDetailAcc = Mid(NoAccountAcc, DetailAcct, mLenIndex.lenidDetailAcc)
     mIndex.idSubDetailAcc = Mid(NoAccountAcc, SubDetailAcct, mLenIndex.lenidSubDetailAcc)
     mIndex.idListAcc = Mid(NoAccountAcc, ListAcct, mLenIndex.lenidListAcc)
     mIndex.idDetailListAcc = Right(NoAccountAcc, mLenIndex.lenidDetailListAcc)
     Select Case UCase(GroupAcc)
            Case "GROUP ACCOUNT":
                 mTampung = KasiZero(Trim(Str(mIndexAuto + 1)), mVarTotalDigit)
'            Case "DETAIL ACCOUNT":
'                 mTampung = mIndex.idGroupAcc & KasiZero(Str(mIndexAuto + 1), mLenIndex.lenidDetailAcc)
'                 mTampung = KasiZero(mTampung, mVarTotalDigit)
            Case "SUB ACCOUNT":
                 mTampung = mIndex.idGroupAcc & KasiZero(Trim(Str(mIndexAuto + 1)), mLenIndex.lenidSubDetailAcc)
                 mTampung = KasiZero(mTampung, mVarTotalDigit)
            Case "LIST ACCOUNT":
                 mTampung = mIndex.idGroupAcc & mIndex.idSubDetailAcc & KasiZero(Str(mIndexAuto + 1), mLenIndex.lenidListAcc)
                 mTampung = KasiZero(mTampung, mVarTotalDigit)
            Case "DETAIL LIST ACCOUNT":
                 mTampung = mIndex.idGroupAcc & mIndex.idSubDetailAcc & mIndex.idListAcc & KasiZero(Str(mIndexAuto + 1), mLenIndex.lenidDetailListAcc)
                 mTampung = KasiZero(mTampung, mVarTotalDigit)
    End Select
End With

AutoIndexAcc = Left(mTampung, mLenIndex.lenidGroupAcc) & Trim(mLenIndex.lenPreGroupAcc) & _
               Mid(mTampung, mLenIndex.lenidGroupAcc + Len(Trim(mLenIndex.lenPreGroupAcc)) + 1, mLenIndex.lenidSubDetailAcc) & Trim(mLenIndex.lenPreSubDetailAcc) & _
               Mid(mTampung, mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + 1, mLenIndex.lenidSubDetailAcc) & Trim(mLenIndex.lenPreListAcc) & _
               Right(mTampung, mLenIndex.lenidDetailListAcc) & Trim(mLenIndex.lenPreDetailListAcc)
               
End Function

Private Sub LoadPeriode()
'Exit Sub
On Error Resume Next
Dim RcBaca As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Dim mVarPer As Integer
RcBaca.DBOpen "SELECT SettingPeriod.GlFile FROM SettingPeriod WHERE (((Left([SettingPeriod].[GlFile],4))='" & TahunFiskalYear & "')) ORDER BY SettingPeriod.GlFile", CNN, lckLockReadOnly

'PANCINGAN
I = IIf(Not IsNull(MyDDE.GetFieldByName(mVarPer)), MyDDE.GetFieldByName(mVarPer), 0)
'Buka Grid Saldo
'Set rcPrd = New DBQuick
CloseDB rcPrd
Set rcPrd = New Recordset
With rcPrd
    .Fields.Append "Periode", adBSTR
    .Fields.Append "Balance", adCurrency
    .Open
End With
CloseDB rcti
Set rcti = New Recordset
With rcti
    .Fields.Append "Periode", adBSTR
    .Fields.Append "Amount", adCurrency
    .Open
End With
CloseDB rcTD
Set rcTD = New Recordset
With rcTD
     .Fields.Append "Periode", adBSTR
     .Fields.Append "Amount", adCurrency
     .Open
End With

With RcBaca.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            rcPrd.AddNew 0, Avdata(0, I)
            mVarPer = 6 + Val(Right(Avdata(0, I), 2))
            rcPrd.Fields(1) = IIf(Not IsNull(MyDDE.GetFieldByName(mVarPer)), MyDDE.GetFieldByName(mVarPer), 0)

            rcti.AddNew 0, Avdata(0, I)
            mVarPer = 20 + Val(Right(Avdata(0, I), 2))
            rcti.Fields(1) = IIf(Not IsNull(MyDDE.GetFieldByName(mVarPer)), MyDDE.GetFieldByName(mVarPer), 0)

            rcTD.AddNew 0, Avdata(0, I)
            mVarPer = 34 + Val(Right(Avdata(0, I), 2))
            rcTD.Fields(1) = IIf(Not IsNull(MyDDE.GetFieldByName(mVarPer)), MyDDE.GetFieldByName(mVarPer), 0)
        Next I
     End If
End With
RcBaca.CloseDB
Set SaldoTrans.DataSource = rcPrd
Set dgBudget(0).DataSource = rcti
Set dgBudget(1).DataSource = rcTD
Err.Clear
End Sub

Private Sub OpenSaldo(ByVal noAccount As String)
Dim RcSaldo As New DBQuick
Dim mSaldo As Variant
Dim mAkhir As Variant
Dim mPer As Integer
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
txtSaldo(0) = 0
txtSaldo(1) = 0
txtSaldo(2) = 0
txtSaldo(3) = 0
RcSaldo.DBOpen "SELECT  CurrentDR" & mPer & ", CurrentCR" & mPer & ", CurrentDR" & mVarPeriode & _
" - CurrentDR" & mPer & " AS CurrentDR, CurrentCR" & mVarPeriode & " - CurrentCR" & mPer & _
" AS CurrentCR FROM [Tabel Pembantu] WHERE     (NoAccount = N'" & noAccount & "')", CNN, lckLockReadOnly
'Debug.Print RcSaldo.DBRecordset.Source
With RcSaldo.DBRecordset
     If .Recordcount <> 0 Then
        mSaldo = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) - IIf(Not IsNull(.Fields(1)), .Fields(1), 0)
        txtSaldo(0) = FormatNumber(Abs(mSaldo), 0)
        txtSaldo(1) = FormatNumber(Abs(IIf(Not IsNull(.Fields(2)), .Fields(2), 0)), 0)
        txtSaldo(2) = FormatNumber(Abs(IIf(Not IsNull(.Fields(3)), .Fields(3), 0)), 0)
        mAkhir = (mSaldo + CCur(txtSaldo(1))) - CCur(txtSaldo(2))
        txtSaldo(3) = FormatNumber(Abs(mAkhir), 0)
     End If
End With
'RcSaldo.CloseDB
End Sub

Private Sub NodeAdo(ByVal NodeaddString As String)
On Error GoTo Hell
Select Case UCase(NodeaddString)
       Case "NEW":
            Select Case UCase(cboType(0).BoundText)
                   Case "GROUP ACCOUNT":
                        If mIndex.idIndukAcc = "ON TOP" Then
                           MenuAccount.NodeAdd , , "[" & mskLedger & "]", Left(mskLedger, mLenIndex.lenidGroupAcc) & " - " & txtLedger(0), "ON TOP", , , True, , True, True, , &HFCF1ED, &H6D4016
                        Else
                           MenuAccount.NodeAdd mIndex.idIndukAcc, tvwChild, "[" & mskLedger & "]", Left(mskLedger, mLenIndex.lenidGroupAcc) & " - " & txtLedger(0), "ON TOP", , , True, , True, True, , &HFCF1ED, &H6D4016
                        End If
'                   Case "DETAIL ACCOUNT":
'                         MenuAccount.NodeAdd "[" & mIndex.idIndukAcc & "]", tvwChild, "[" & mskLedger & "]", Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreDetailAcc), mLenIndex.lenidDetailAcc) & " - " & txtLedger(0), "DETAIL ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
                   Case "SUB ACCOUNT"
'                         MenuAccount.NodeAdd "[" & mIndex.idIndukAcc & "]", tvwChild, "[" & mskLedger & "]", Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreDetailAcc) + mLenIndex.lenidDetailAcc + Len(mLenIndex.lenPreDetailAcc), mLenIndex.lenidSubDetailAcc) & " - " & txtLedger(0), "SUB ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
'                        Debug.Print Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreSubDetailAcc) + mLenIndex.lenidSubDetailAcc, mLenIndex.lenidSubDetailAcc) & " - " & txtLedger(0)
                         MenuAccount.NodeAdd "[" & mIndex.idIndukAcc & "]", tvwChild, "[" & mskLedger & "]", Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreSubDetailAcc) + mLenIndex.lenidSubDetailAcc, mLenIndex.lenidSubDetailAcc) & " - " & txtLedger(0), "SUB ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
                   Case "LIST ACCOUNT":
'                         MenuAccount.NodeAdd "[" & mIndex.idIndukAcc & "]", tvwChild, "[" & mskLedger & "]", Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreDetailAcc) + mLenIndex.lenidDetailAcc + Len(mLenIndex.lenPreDetailAcc) + mLenIndex.lenidSubDetailAcc + Len(mLenIndex.lenidSubDetailAcc), mLenIndex.lenidListAcc) & " - " & txtLedger(0), "LIST ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
                         MenuAccount.NodeAdd "[" & mIndex.idIndukAcc & "]", tvwChild, "[" & mskLedger & "]", Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreSubDetailAcc) + mLenIndex.lenidSubDetailAcc + Len(mLenIndex.lenPreListAcc) + Len(mLenIndex.lenidListAcc), mLenIndex.lenidListAcc) & " - " & txtLedger(0), "LIST ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
                   Case "DETAIL LIST ACCOUNT":
                         MenuAccount.NodeAdd "[" & mIndex.idIndukAcc & "]", tvwChild, "[" & mskLedger & "]", Right(mskLedger, mLenIndex.lenidDetailListAcc) & " - " & txtLedger(0), "DETAIL LIST ACCOUNT", , , , , True, True, , &HFCF1ED, &H6D4016
            End Select
            If MyDDE.ActiveRecordset.Recordcount > 1 Then
               Set mVarNodeNode = MenuAccount.MenuTreeView.Nodes(mVarNodeNode.LastSibling.Index)
               mVarNodeNode.Selected = True
               MyDDE.ActiveRecordset.AbsolutePosition = mVarNodeNode.LastSibling.Index
            End If
       Case "EDIT":
            Select Case UCase(cboType(0).BoundText)
                   Case "GROUP ACCOUNT":
                        If mIndex.idIndukAcc = "ON TOP" Then
                           mVarnode.Item(Val(mVarIndex)).Text = Left(mskLedger, mLenIndex.lenidGroupAcc) & " - " & txtLedger(0)
                        Else
                           mVarnode.Item(Val(mVarIndex)).Text = Left(mskLedger, mLenIndex.lenidGroupAcc) & " - " & txtLedger(0)
                        End If
'                   Case "DETAIL ACCOUNT":
'                         mVarnode.Item(Val(mVarIndex)).Text = Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreDetailAcc), mLenIndex.lenidDetailAcc) & " - " & txtLedger(0)
                   Case "SUB ACCOUNT"
'                         mVarnode.Item(Val(mVarIndex)).Text = Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenPreDetailAcc) + mLenIndex.lenidDetailAcc + Len(mLenIndex.lenPreDetailAcc), mLenIndex.lenidSubDetailAcc) & " - " & txtLedger(0)
                         mVarnode.Item(Val(mVarIndex)).Text = Mid(mskLedger, mLenIndex.lenidGroupAcc + Len(mLenIndex.lenidSubDetailAcc), mLenIndex.lenidSubDetailAcc) & " - " & txtLedger(0)
                   Case "LIST ACCOUNT":
                         mVarnode.Item(Val(mVarIndex)).Text = Mid(mskLedger, mLenIndex.lenidGroupAcc + mLenIndex.lenidSubDetailAcc + Len(mLenIndex.lenidSubDetailAcc), mLenIndex.lenidListAcc) & " - " & txtLedger(0)
                   Case "DETAIL LIST ACCOUNT":
                         mVarnode.Item(Val(mVarIndex)).Text = Right(mskLedger, mLenIndex.lenidDetailListAcc) & " - " & txtLedger(0)
            End Select
       Case "DELETE":
            mVarnode.Remove (Val(mVarIndex))
End Select

Exit Sub
Hell:
MessageBox Err.Description, "frmPerkiraan : NodeAdo", msgOkOnly, msgExclamation
mVarIndex = ""
Err.Clear
End Sub

Private Sub CompareAccount()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
If Not CNN Is Nothing Then
   If CNN.State = 1 Then
      If RcCom.DBOpen("SELECT GLAccount.NoAccount, [Tabel Pembantu].NoAccount AS NoAccountB FROM         GLAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly) = True Then
        With RcCom.DBRecordset
             If .Recordcount <> 0 Then
                mVarData = .Getrows(.Recordcount, adBookmarkFirst)
                For I = 0 To UBound(mVarData, 2)
                    If IsNull(mVarData(1, I)) Then
                        SendDataToServer ("INSERT INTO [Tabel Pembantu] (NoAccount) VALUES (N'" & mVarData(0, I) & "')")
                    End If
                Next I
             End If
        End With
      End If
   End If
End If
Set mVarData = Nothing
Set RcCom = Nothing
End Sub

'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''MoveForm Picture1.Parent.hwnd
'End Sub

Private Sub OpenListAccount()
On Error GoTo Hell
Select Case UCase(cboType(0).BoundText)
    Case "GROUP ACCOUNT":
        RcGroup.DBRecordset.AbsolutePosition = 2
         'mIndex.idIndukAcc = "ON TOP"
'       Case "Detail Account":
'            RcGroup.DBRecordset.AbsolutePosition = 3
    Case "SUB ACCOUNT":
        RcGroup.DBRecordset.AbsolutePosition = 3
    Case "LIST ACCOUNT":
        RcGroup.DBRecordset.AbsolutePosition = 4
    Case "DETAIL LIST ACCOUNT":
        RcGroup.DBRecordset.AbsolutePosition = 4
    Case Else
        RcGroup.DBRecordset.AbsolutePosition = 1
End Select
Exit Sub

Hell:
    If RcGroup.DBRecordset.AbsolutePosition > 1 Then
       RcGroup.DBRecordset.AbsolutePosition = RcGroup.DBRecordset.AbsolutePosition - 1
    Else
       RcGroup.DBRecordset.AbsolutePosition = 1
    End If
End Sub

Private Function OpenAccountInduk(ByVal GetNoaccount As String) As String
Dim Rc As New DBQuick
Rc.DBOpen "SELECT GroupAccount FROM GLAccount WHERE (NoAccount = N'" & GetNoaccount & "')", CNN, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        OpenAccountInduk = IIf(Not IsNull(.Fields(0)), .Fields(0), "ON TOP")
     Else
        OpenAccountInduk = "ON TOP"
     End If
End With
Rc.CloseDB
Set Rc = Nothing
End Function







