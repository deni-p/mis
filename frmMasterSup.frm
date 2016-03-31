VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{EB0E2EAE-5969-4167-B57F-56BCD8266DF2}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMasterSup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Card"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterSup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Tag             =   "Supplier"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6090
      Left            =   0
      ScaleHeight     =   6090
      ScaleWidth      =   9855
      TabIndex        =   23
      Top             =   0
      Width           =   9855
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Petty cash Supplier"
         DataField       =   "default"
         DataSource      =   "MyDDE"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7515
         TabIndex        =   40
         Tag             =   "Partner"
         Top             =   3510
         Width           =   2070
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "term_desc"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   13
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "Partner"
         Top             =   2280
         Width           =   2850
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   9210
         Picture         =   "frmMasterSup.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2288
         Width           =   315
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   9195
         Picture         =   "frmMasterSup.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Country"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   7
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "Partner"
         Top             =   120
         Width           =   2835
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4635
         Picture         =   "frmMasterSup.frx":6F66
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1928
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   5
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1920
         Width           =   2820
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "URL"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   12
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   15
         Tag             =   "Partner"
         Top             =   1920
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Email"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   11
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "Partner"
         Top             =   1560
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Fax"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   10
         Left            =   6360
         MaxLength       =   24
         TabIndex        =   13
         Tag             =   "Partner"
         Top             =   1200
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Mobile"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   9
         Left            =   6360
         MaxLength       =   24
         TabIndex        =   12
         Tag             =   "Partner"
         Top             =   840
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Phone"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   8
         Left            =   6360
         MaxLength       =   24
         TabIndex        =   11
         Tag             =   "Partner"
         Top             =   480
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PostalCode"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   6
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Partner"
         Top             =   2280
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Address"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   4
         Left            =   1815
         MaxLength       =   60
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1560
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ContactTitle"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   3
         Left            =   1815
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1200
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "ContactName"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   2
         Left            =   1815
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   840
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   1
         Left            =   1815
         MaxLength       =   40
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   480
         Width           =   3195
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "PartnerID"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1815
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   120
         Width           =   3195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAAF6F&
         Height          =   795
         Left            =   105
         TabIndex        =   24
         Top             =   2640
         Width           =   9465
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "TAX Rate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   315
            Index           =   14
            Left            =   4890
            MaxLength       =   3
            TabIndex        =   19
            Tag             =   "Partner"
            Top             =   300
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "WHT"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   315
            Index           =   15
            Left            =   6030
            MaxLength       =   3
            TabIndex        =   20
            Tag             =   "Partner"
            Top             =   300
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Vat"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   315
            Index           =   16
            Left            =   7110
            MaxLength       =   3
            TabIndex        =   21
            Tag             =   "Partner"
            Top             =   300
            Visible         =   0   'False
            Width           =   570
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   330
            Left            =   915
            TabIndex        =   18
            Top             =   285
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   20
            Format          =   "##.###.###.#-###.###"
            Mask            =   "##.###.###.#-###.###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N.P.W.P"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   14
            Left            =   90
            TabIndex        =   25
            Top             =   345
            Width           =   690
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   75
            X2              =   975
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            Index           =   1
            Visible         =   0   'False
            X1              =   4035
            X2              =   4935
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            Index           =   2
            Visible         =   0   'False
            X1              =   5520
            X2              =   6420
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            Index           =   3
            X1              =   0
            X2              =   1920
            Y1              =   -15
            Y2              =   -15
         End
         Begin VB.Line Line2 
            Index           =   4
            Visible         =   0   'False
            X1              =   6675
            X2              =   7575
            Y1              =   600
            Y2              =   600
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2415
         Left            =   105
         TabIndex        =   22
         Top             =   3570
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4260
         _Version        =   393216
         Style           =   1
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
         TabCaption(0)   =   "Supplier"
         TabPicture(0)   =   "frmMasterSup.frx":72F0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Bank Partner"
         TabPicture(1)   =   "frmMasterSup.frx":730C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Item"
         TabPicture(2)   =   "frmMasterSup.frx":7328
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture4"
         Tab(2).ControlCount=   1
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000000&
            Height          =   1950
            Left            =   -74925
            ScaleHeight     =   1890
            ScaleWidth      =   9270
            TabIndex        =   45
            Top             =   375
            Width           =   9330
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frmMasterSup.frx":7344
               Height          =   1920
               Index           =   2
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   3387
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   6
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "Item/Service"
                  Caption         =   "Item"
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
                  DataField       =   "Nama Item/Service"
                  Caption         =   "Nama Item"
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
                  DataField       =   "Merk"
                  Caption         =   "Merk"
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
               BeginProperty Column03 
                  DataField       =   "Serial Supplier Code"
                  Caption         =   "Serial Supplier Code"
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
                  DataField       =   "UOM"
                  Caption         =   "UOM"
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   4
                  BeginProperty Column00 
                     ColumnWidth     =   1800
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   1800
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000000&
            Height          =   1950
            Left            =   -74925
            ScaleHeight     =   1890
            ScaleWidth      =   9270
            TabIndex        =   43
            Top             =   375
            Width           =   9330
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frmMasterSup.frx":7359
               Height          =   1920
               Index           =   1
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   3387
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   6
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "Bank Name"
                  Caption         =   "Nama Bank"
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
                  DataField       =   "Account"
                  Caption         =   "No Rekening"
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
                  DataField       =   "Address"
                  Caption         =   "Alamat"
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
               BeginProperty Column03 
                  DataField       =   "Currency"
                  Caption         =   "Currency"
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
                  DataField       =   "Default"
                  Caption         =   "Default"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "YES"
                     FalseValue      =   "NO"
                     NullValue       =   "YES"
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   4
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000000&
            Height          =   1950
            Left            =   75
            ScaleHeight     =   1890
            ScaleWidth      =   9270
            TabIndex        =   41
            Top             =   375
            Width           =   9330
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "frmMasterSup.frx":736E
               Height          =   1920
               Index           =   0
               Left            =   0
               TabIndex        =   42
               Tag             =   "Partner"
               Top             =   0
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   3387
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   6
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
               ColumnCount     =   13
               BeginProperty Column00 
                  DataField       =   "PartnerID"
                  Caption         =   "Partner ID"
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
                  DataField       =   "CompanyName"
                  Caption         =   "Nama Perusahaan"
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
                  DataField       =   "ContactName"
                  Caption         =   "Nama Kontak"
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
               BeginProperty Column03 
                  DataField       =   "ContactTitle"
                  Caption         =   "Jabatan"
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
                  DataField       =   "Address"
                  Caption         =   "Alamat"
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
               BeginProperty Column05 
                  DataField       =   "City"
                  Caption         =   "Kota"
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
               BeginProperty Column06 
                  DataField       =   "PostalCode"
                  Caption         =   "Kode Pos"
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
               BeginProperty Column07 
                  DataField       =   "Country"
                  Caption         =   "Kode Negara"
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
               BeginProperty Column08 
                  DataField       =   "Phone"
                  Caption         =   "Telp"
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
               BeginProperty Column09 
                  DataField       =   "Mobile"
                  Caption         =   "Mobile No"
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
               BeginProperty Column10 
                  DataField       =   "Fax"
                  Caption         =   "Faximile"
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
               BeginProperty Column11 
                  DataField       =   "Email"
                  Caption         =   "Email"
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
               BeginProperty Column12 
                  DataField       =   "URL"
                  Caption         =   "Web Site"
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   4
                  BeginProperty Column00 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column01 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column02 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column03 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column04 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column05 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column06 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column07 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column08 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column09 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column10 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column11 
                     DividerStyle    =   6
                  EndProperty
                  BeginProperty Column12 
                     DividerStyle    =   6
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   13
         Left            =   5130
         TabIndex        =   39
         Top             =   188
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   12
         Left            =   5130
         TabIndex        =   38
         Top             =   548
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   11
         Left            =   5130
         TabIndex        =   37
         Top             =   908
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   5130
         TabIndex        =   36
         Top             =   1268
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   5130
         TabIndex        =   35
         Top             =   1628
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   5130
         TabIndex        =   34
         Top             =   1988
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   33
         Top             =   2355
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   32
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   31
         Top             =   1635
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   30
         Top             =   1275
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   29
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   555
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   27
         Top             =   195
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1845
         X2              =   75
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1845
         X2              =   75
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1845
         X2              =   75
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1845
         X2              =   75
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1845
         X2              =   75
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   1845
         X2              =   75
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   6855
         X2              =   5085
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   6855
         X2              =   5085
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   6855
         X2              =   5085
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   6855
         X2              =   5085
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   6855
         X2              =   5085
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   6855
         X2              =   5085
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   1845
         X2              =   75
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   5130
         TabIndex        =   26
         Top             =   2348
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   5085
         X2              =   6930
         Y1              =   2595
         Y2              =   2595
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6090
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmMasterSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyData As clsMaster
Private RcBank As New Recordset
Private Rcitem As Recordset
Private MProp As Boolean
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner As DBQuick
Dim strSQL As String

Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub Form_Activate()
'If Me.WindowState = 0 Then If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
MyDDE.SetPermissions = aksess.MayDo("Supplier Card")

HiasFormManTell Picture2, Me
GridLayout
HiasFormManTell Picture2, Me
SSTab1.BackColor = Picture2.BackColor
Frame2.BackColor = Picture2.BackColor
Set MyData = New clsMaster
SSTab1.Tab = 0
With MyDDE
     .EditModeReplace = False
     Set .BindForm = frmMasterSup
     .BindFormTAG = "Partner"
     Set .ActiveConnection = CNN
     strSQL = "SELECT PartnerDB.PartnerID, PartnerDB.CompanyName, PartnerDB.ContactName, PartnerDB.ContactTitle, " & _
            " PartnerDB.Address, PartnerDB.City, PartnerDB.PostalCode, PartnerDB.Country, PartnerDB.Phone, PartnerDB.Mobile, " & _
            " PartnerDB.Fax, PartnerDB.Email, PartnerDB.URL, PartnerDB.NPWP, PartnerDB.WHT, PartnerDB.[TAX Rate], " & _
            " PartnerDB.Vat, PartnerDB.PartnerType, PartnerDB.NoAccount, PartnerDB.Term_code, " & _
            " PartnerDB.[default], termpayment.Description AS term_desc " & _
            " FROM PartnerDB INNER JOIN termpayment ON PartnerDB.Term_code = termpayment.Code " & _
            " WHERE (PartnerDB.PartnerType = 'SUPPLIER') ORDER BY PartnerDB.PartnerID"
            
     .PrepareQuery = strSQL '"Select * from PartnerDB where PartnerType='SUPPLIER' Order By PartnerID"
End With
Set mCall = New frmCaller

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
    Set MyData = Nothing
    MyDDE.ClearRecordset
    Set mCall = Nothing
If Me.Tag = "Supplier" Then
   IsFrmSup = False
Else
   IsFrmCus = False
End If
   Else
      Cancel = True
   End If
Else
    Set MyData = Nothing
    MyDDE.ClearRecordset
    Set mCall = Nothing
If Me.Tag = "Supplier" Then
   IsFrmSup = False
Else
   IsFrmCus = False
End If
End If
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'SSTab1.BackColor = Picture2.BackColor
'Frame2.BackColor = Picture2.BackColor
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMasterSup = Nothing
End Sub

Private Sub MaskEdBox1_Change()
If MProp = True Then MyDDE.GetFieldByName("NPWP") = MaskEdBox1
End Sub

Private Sub MaskEdBox1_GotFocus()
On Error Resume Next
Block MaskEdBox1
Err.Clear
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.FromTagActive
       Case "DAFTAR NAMA KOTA":
           If txtBox(6).Enabled = True Then txtBox(6).SetFocus
       Case "DAFTAR NAMA NEGARA":
           If txtBox(8).Enabled = True Then txtBox(8).SetFocus
End Select
End Sub

Private Sub mCall_CallLinkForm()
Select Case mCall.FromTagActive
       Case "DAFTAR NAMA KOTA":
            frmRegional.SetFocus
            frmRegional.OptCity(0).Value = True
            frmRegional.ZOrder (0)
            
       Case "DAFTAR NAMA NEGARA":
            frmRegional.SetFocus
            frmRegional.OptCity(1).Value = True
            frmRegional.ZOrder
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case mCall.FromTagActive
       Case "DAFTAR NAMA KOTA":
             With mCall
                  MyDDE.GetFieldByName("City") = .GetFieldByName(1)
             End With
       Case "DAFTAR NAMA NEGARA":
             With mCall
                  MyDDE.GetFieldByName("Country") = .GetFieldByName(1)
             End With
       Case "SYARAT PEMBAYARAN":
             With mCall
                  MyDDE.GetFieldByName("Term_code") = .GetFieldByName(0)
                  MyDDE.GetFieldByName("term_desc") = .GetFieldByName("Description")
             End With
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterPartner) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly, msgCrtical
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            MaskEdBox1.Mask = "##.###.###.#-###.###"
            MaskEdBox1.Format = "##.###.###.#-###.###"
            If Me.Tag = "Supplier" Then
               MyDDE.GetFieldByName("PartnerID") = MyData.PrepareIndex(tmbSupplier, 7, "", "SP")
            Else
               MyDDE.GetFieldByName("PartnerID") = MyData.PrepareIndex(tmbCustomer, 7, "", "CS")
            End If
            MyDDE.GetFieldByName("ContactName") = "-"
            MyDDE.GetFieldByName("ContactTitle") = "-"
            MyDDE.GetFieldByName("Phone") = "-"
            MyDDE.GetFieldByName("Mobile") = "-"
            MyDDE.GetFieldByName("Fax") = "-"
            MyDDE.GetFieldByName("email") = "-"
            MyDDE.GetFieldByName("URL") = "-"
            MyDDE.GetFieldByName("PostalCode") = "-"
            MyDDE.GetFieldByName("TAX Rate") = 0
            MyDDE.GetFieldByName("WHT") = 0
            MyDDE.GetFieldByName("Vat") = 0
            txtBox(0).Enabled = False
            txtBox(1).SetFocus
            MProp = True
            MaskEdBox1.Enabled = MProp
            
       Case tmbEdit:
            MaskEdBox1.Mask = "##.###.###.#-###.###"
            MaskEdBox1.Format = "##.###.###.#-###.###"
            txtBox(0).Enabled = False
            txtBox(1).SetFocus
            MProp = True
            MaskEdBox1.Enabled = MProp
       Case tmbCancel:
            MaskEdBox1.Mask = ""
            MaskEdBox1.Format = ""
            MProp = False
            MaskEdBox1.Enabled = MProp
       Case tmbSave:
            MaskEdBox1.Mask = ""
            MaskEdBox1.Format = ""
            MProp = False
            MaskEdBox1.Enabled = MProp
       Case tmbPrint:
            If Me.Tag = "Supplier" Then
               CallRPTReport "Daftar Supplier.rpt"
            Else
               CallRPTReport "Daftar Customer.rpt"
            End If
       Case tmbQuit:
'            Unload Me
'            Set MyDDE.BindForm = Nothing
End Select
txtBox(5).Enabled = False
txtBox(7).Enabled = False
txtBox(13).Enabled = False
MaskEdBox1.Enabled = txtBox(1).Enabled
cmdLink(0).Enabled = txtBox(1).Enabled
cmdLink(1).Enabled = txtBox(1).Enabled
cmdLink(2).Enabled = txtBox(1).Enabled
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim mVarNpwp As String
CloseDB RcBank
Set RcBank = MyData.OpenBankAccount(IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), ""))
Set DataGrid1(1).DataSource = RcBank
OpenItem IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "")
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   mVarNpwp = IIf(Not IsNull(MyDDE.GetFieldByName("NPWP")), MyDDE.GetFieldByName("NPWP"), "")
   If mVarNpwp = "" Then
      MaskEdBox1.Mask = ""
      MaskEdBox1.Format = ""
      MaskEdBox1 = ""
   Else
      MaskEdBox1 = mVarNpwp
   End If
Else
   MaskEdBox1.Mask = ""
   MaskEdBox1.Format = ""
   MaskEdBox1 = ""
End If

'If MyDDE.GetFieldByName("default") Then chk.Value = 1 Else chk.Value = 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyTab Then
   If DataGrid1(SSTab1.Tab).Enabled = True Then
      DataGrid1(SSTab1.Tab).SetFocus
   Else
      MyDDE.SetFocus
   End If
End If
End Sub

Private Sub txtBox_Change(Index As Integer)
If MProp = True Then
   If Index = 14 Or Index = 15 Or Index = 16 Then
      If txtBox(Index) = "-" Then txtBox(Index) = 0
   End If
End If
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
Dim TypTr As String
With MyDDE
   If Me.Tag = "Supplier" Then
      TypTr = "SUPPLIER"
   Else
      TypTr = "CUSTOMER"
   End If
   .PrepareAppend = " INSERT INTO PartnerDB (PartnerID, CompanyName, ContactName, ContactTitle, Address, City, PostalCode, Country, Phone, Mobile, Fax, Email, URL, NPWP, WHT, [TAX Rate],Vat, PartnerType,term_code,[default]) " & _
                    " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "', N'" & ValidString(txtBox(3)) & "', N'" & ValidString(txtBox(4)) & "', N'" & ValidString(txtBox(5)) & "', N'" & ValidString(txtBox(6)) & "', N'" & ValidString(txtBox(7)) & "' ," & _
                    "  N'" & ValidString(txtBox(8)) & "', N'" & ValidString(txtBox(9)) & "', N'" & ValidString(txtBox(10)) & "', N'" & ValidString(txtBox(11)) & "', N'" & ValidString(txtBox(12)) & "', N'" & ValidString(MaskEdBox1) & "', " & CDbl(txtBox(14)) & ", " & CDbl(txtBox(15)) & ", " & CDbl(txtBox(15)) & ", N'" & TypTr & "','" & .GetFieldByName("term_code") & "'," & chk.Value & ")"
                    
   .PrepareUpdate = " UPDATE PartnerDB Set CompanyName=N'" & ValidString(txtBox(1)) & "', ContactName = N'" & ValidString(txtBox(2)) & "', ContactTitle = N'" & ValidString(txtBox(3)) & "', Address = N'" & ValidString(txtBox(4)) & "', City = N'" & ValidString(txtBox(5)) & "', PostalCode = N'" & ValidString(txtBox(6)) & "', Country = N'" & ValidString(txtBox(7)) & "', Phone = N'" & ValidString(txtBox(8)) & "', Mobile = N'" & ValidString(txtBox(9)) & "', Fax = N'" & ValidString(txtBox(10)) & "'," & _
                    " Email = N'" & ValidString(txtBox(11)) & "', URL = N'" & ValidString(txtBox(12)) & "', NPWP = N'" & ValidString(MaskEdBox1) & "', WHT = " & CDbl(txtBox(14)) & ", [TAX Rate] = " & CDbl(txtBox(15)) & ", Vat = " & CDbl(txtBox(16)) & ",term_code='" & .GetFieldByName("term_code") & "',[default]=" & chk.Value & "  WHERE (PartnerID = N'" & ValidString(txtBox(0)) & "') AND (PartnerType = N'" & TypTr & "')"
                    
   .PrepareDelete = " DELETE FROM PartnerDB WHERE   (PartnerType = N'" & TypTr & "') AND (PartnerID = N'" & ValidString(txtBox(0)) & "')"
End With
End Sub

Private Sub OpenItem(ByVal NoPartnerID As String)
CloseDB Rcitem
Set Rcitem = New Recordset
Rcitem.CursorLocation = adUseClient
Rcitem.Open "SELECT [Detail PO].NoItem AS [Item/Service], Inventory.ItemName AS [Nama Item/Service], Inventory.Merk, Inventory.[Serial Supplier] AS [Serial Supplier Code], Inventory.UOM FROM [PO Order] INNER JOIN [Detail PO] ON [PO Order].PurchaseID = [Detail PO].PurchaseID INNER JOIN Inventory ON [Detail PO].NoItem = Inventory.NoItem WHERE     ([PO Order].PartnerID = N'" & NoPartnerID & "') ORDER BY [Detail PO].NoItem", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
Set Rcitem.ActiveConnection = Nothing
Set DataGrid1(2).DataSource = Rcitem
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 14 Or Index = 15 Or Index = 16 Then
   ValidNum KeyAscii
End If
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:

Set RcPartner = New DBQuick

Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT RG AS [Kode Kota], [RG Name] AS [Nama Kota], [Code RG] AS [Kode Regional] " & _
            " FROM Regional WHERE ([Type RG] = N'CITY') ORDER BY [RG Name]", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen "SELECT RG AS [Kode Negara], [RG Name] AS [Nama Negara], [Code RG] AS [Kode Regional] " & _
            " FROM Regional WHERE ([Type RG] = N'COUNTRY') ORDER BY [RG Name]", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT Code AS Kode , [Due Date Calculation] , [Discount Date Calculation], " & _
            " [Discount %], Description   FROM   TermPayment ", CNN, lckLockReadOnly
End Select

If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0:
               mCall.FromTagActive = "DAFTAR NAMA KOTA"
               mCall.CaptionLink = "Kota"
           Case 1:
               mCall.FromTagActive = "DAFTAR NAMA NEGARA"
               mCall.CaptionLink = "Negara"
           Case 2:
               mCall.FromTagActive = "SYARAT PEMBAYARAN"
               mCall.CaptionLink = "Term Payment"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    If mCall.FromTagActive = "MASTER BANK" Then mCall.SetFormat(3) = "YES/NO"
    mCall.LookUp Me
    
'    If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
'       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'       DGPurchase.SetFocus
'    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
End If

Exit Sub
Hell:
    'messagebox Err.Description
    Err.Clear
End Sub
Private Sub GridLayout()
DataGrid1(0).Columns(0).width = 1709.858
DataGrid1(0).Columns(1).width = 2670.236
DataGrid1(0).Columns(2).width = 2475.213
DataGrid1(0).Columns(3).width = 1874.835
DataGrid1(0).Columns(4).width = 1170.142
DataGrid1(0).Columns(5).width = 1514.835
DataGrid1(0).Columns(6).width = 1514.835
DataGrid1(0).Columns(7).width = 1514.835
DataGrid1(0).Columns(8).width = 1514.835
DataGrid1(0).Columns(9).width = 1514.835
DataGrid1(0).Columns(10).width = 1514.835
DataGrid1(0).Columns(11).width = 1514.835
DataGrid1(0).Columns(12).width = 1514.835
DataGrid1(1).Columns(0).width = 1514.835
DataGrid1(1).Columns(1).width = 2654.929
DataGrid1(1).Columns(2).width = 1514.835
DataGrid1(1).Columns(3).width = 1514.835
DataGrid1(1).Columns(4).width = 1514.835
DataGrid1(2).Columns(0).width = 1800
DataGrid1(2).Columns(1).width = 2489.953
DataGrid1(2).Columns(2).width = 1800
DataGrid1(2).Columns(3).width = 1739.906
DataGrid1(2).Columns(4).width = 900.2835

End Sub


