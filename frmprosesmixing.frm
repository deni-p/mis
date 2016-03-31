VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMixingMilling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catatan Proses Mixing Chips & Milling Powder"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmprosesmixing.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin SemeruDC.SemeruOleDC DDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5700
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1005
      BindFormTAG     =   "mixing"
      InitControlSet  =   1
      ActiveLanguage  =   1
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
      Height          =   5685
      Left            =   0
      ScaleHeight     =   5685
      ScaleWidth      =   10740
      TabIndex        =   24
      Top             =   0
      Width           =   10740
      Begin TabDlg.SSTab SSTab1 
         Height          =   5550
         Left            =   45
         TabIndex        =   1
         Top             =   60
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   9790
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   15380335
         TabCaption(0)   =   "Mixing Chips"
         TabPicture(0)   =   "frmprosesmixing.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line1(4)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Line1(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(14)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(15)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Line1(6)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(16)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Line1(7)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Line1(8)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label1(17)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Line1(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label1(18)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label1(19)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label1(20)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Line1(10)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Line1(11)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label1(21)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label1(2)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Line1(12)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label4"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "tgl(1)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "tgl(0)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "GridDetail"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtMixing(1)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtMixing(0)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtMixing(2)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtMixing(3)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtMixing(4)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "cmdLink(0)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "cmdLink(1)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "cmdLink(2)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtMixing(5)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).ControlCount=   31
         TabCaption(1)   =   "Milling"
         TabPicture(1)   =   "frmprosesmixing.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1(0)"
         Tab(1).Control(1)=   "Line1(0)"
         Tab(1).Control(2)=   "Label1(13)"
         Tab(1).Control(3)=   "Label1(12)"
         Tab(1).Control(4)=   "Label1(11)"
         Tab(1).Control(5)=   "Label1(8)"
         Tab(1).Control(6)=   "Line1(2)"
         Tab(1).Control(7)=   "Line1(1)"
         Tab(1).Control(8)=   "Label1(7)"
         Tab(1).Control(9)=   "Label1(6)"
         Tab(1).Control(10)=   "Label1(5)"
         Tab(1).Control(11)=   "Label1(4)"
         Tab(1).Control(12)=   "lblMesh150"
         Tab(1).Control(13)=   "lblMesh100"
         Tab(1).Control(14)=   "Label1(1)"
         Tab(1).Control(15)=   "Line1(3)"
         Tab(1).Control(16)=   "GridMilling"
         Tab(1).Control(17)=   "Frame1"
         Tab(1).Control(18)=   "txtMilling(0)"
         Tab(1).Control(19)=   "txtMilling(7)"
         Tab(1).Control(20)=   "txtMilling(3)"
         Tab(1).Control(21)=   "txtMilling(6)"
         Tab(1).Control(22)=   "txtMilling(5)"
         Tab(1).Control(23)=   "txtMilling(4)"
         Tab(1).Control(24)=   "txtMilling(2)"
         Tab(1).Control(25)=   "txtMilling(1)"
         Tab(1).Control(26)=   "viewDate"
         Tab(1).ControlCount=   27
         TabCaption(2)   =   "Test Granulometry && Benda Asing Mesh 100"
         TabPicture(2)   =   "frmprosesmixing.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3"
         Tab(2).Control(1)=   "GridMesh100"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Test Granulometry && Benda Asing Mesh 150"
         TabPicture(3)   =   "frmprosesmixing.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label2"
         Tab(3).Control(1)=   "GridMesh150"
         Tab(3).ControlCount=   2
         Begin VB.TextBox txtMixing 
            Appearance      =   0  'Flat
            DataField       =   "itemName"
            DataSource      =   "DDE"
            Height          =   330
            Index           =   5
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   4
            Tag             =   "mixing"
            Top             =   1395
            Width           =   1740
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   4380
            Picture         =   "frmprosesmixing.frx":68C2
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1410
            Width           =   405
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   4380
            Picture         =   "frmprosesmixing.frx":6C4C
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1763
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   4380
            Picture         =   "frmprosesmixing.frx":6FD6
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1050
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox txtMixing 
            Appearance      =   0  'Flat
            DataSource      =   "DDE"
            Height          =   330
            Index           =   4
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1020
            Visible         =   0   'False
            Width           =   1740
         End
         Begin MSComCtl2.DTPicker viewDate 
            Height          =   345
            Left            =   -69900
            TabIndex        =   14
            Top             =   885
            Visible         =   0   'False
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60096515
            CurrentDate     =   39538
         End
         Begin VB.TextBox txtMixing 
            Appearance      =   0  'Flat
            DataField       =   "total_sesudah"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   3
            Left            =   2640
            TabIndex        =   10
            Tag             =   "mixing"
            Top             =   2790
            Width           =   705
         End
         Begin VB.TextBox txtMixing 
            Appearance      =   0  'Flat
            DataField       =   "total_sebelum"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   9
            Tag             =   "mixing"
            Top             =   2445
            Width           =   705
         End
         Begin VB.TextBox txtMixing 
            Appearance      =   0  'Flat
            DataField       =   "prelot"
            DataSource      =   "DDE"
            Height          =   330
            Index           =   0
            Left            =   2640
            TabIndex        =   6
            Tag             =   "mixing"
            Top             =   1755
            Width           =   1740
         End
         Begin VB.TextBox txtMixing 
            Appearance      =   0  'Flat
            DataField       =   "grup"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   8
            Tag             =   "mixing"
            Top             =   2100
            Width           =   705
         End
         Begin VB.TextBox txtMilling 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   -72480
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0"
            Top             =   4245
            Width           =   930
         End
         Begin VB.TextBox txtMilling 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   -72480
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0"
            Top             =   4605
            Width           =   930
         End
         Begin VB.TextBox txtMilling 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "sakMesh100"
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   -70890
            TabIndex        =   18
            Tag             =   "mixing"
            Text            =   "0"
            Top             =   4245
            Width           =   915
         End
         Begin VB.TextBox txtMilling 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "sakMesh150"
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   -70890
            TabIndex        =   20
            Tag             =   "mixing"
            Text            =   "0"
            Top             =   4605
            Width           =   915
         End
         Begin VB.TextBox txtMilling 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "pre_lot_powder"
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   -70890
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "0"
            Top             =   4965
            Width           =   915
         End
         Begin VB.TextBox txtMilling 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   -72480
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "0"
            Top             =   4965
            Width           =   930
         End
         Begin VB.TextBox txtMilling 
            Appearance      =   0  'Flat
            DataField       =   "desc"
            DataSource      =   "DDE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Index           =   7
            Left            =   -69135
            TabIndex        =   23
            Tag             =   "mixing"
            Top             =   4230
            Width           =   4620
         End
         Begin VB.TextBox txtMilling 
            Appearance      =   0  'Flat
            DataField       =   "prelot"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   0
            Left            =   -73020
            Locked          =   -1  'True
            TabIndex        =   13
            Tag             =   "mixing"
            Top             =   735
            Width           =   1950
         End
         Begin VB.Frame Frame1 
            Caption         =   "Kondisi Magnet"
            Height          =   630
            Left            =   -66525
            TabIndex        =   32
            Tag             =   "mixing"
            Top             =   405
            Width           =   2130
            Begin VB.OptionButton OpKondisi 
               Caption         =   "Kotor"
               Height          =   345
               Index           =   1
               Left            =   1275
               TabIndex        =   16
               Top             =   210
               Width           =   825
            End
            Begin VB.OptionButton OpKondisi 
               Caption         =   "Bersih"
               Height          =   345
               Index           =   0
               Left            =   165
               TabIndex        =   15
               Top             =   210
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin MSDataGridLib.DataGrid GridMesh150 
            Height          =   4680
            Left            =   -74910
            TabIndex        =   28
            Top             =   450
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   8255
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "No"
               Caption         =   "No Bag"
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
               DataField       =   "L80"
               Caption         =   "100 (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "p80"
               Caption         =   "150 (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "p100"
               Caption         =   "100 (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Total"
               Caption         =   "Total (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "kondisiPowder"
               Caption         =   "Kondisi Powder*"
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
               DataField       =   "kuantitas"
               Caption         =   "Kuantitas (Kg)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "operator"
               Caption         =   "Operator"
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
               BeginProperty Column07 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid GridMesh100 
            Height          =   4680
            Left            =   -74910
            TabIndex        =   29
            Top             =   450
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   8255
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "No"
               Caption         =   "No Bag"
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
               DataField       =   "L80"
               Caption         =   ">80 (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "p80"
               Caption         =   "80 (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "p100"
               Caption         =   "100 (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Total"
               Caption         =   "Total (%)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "kondisiPowder"
               Caption         =   "Kondisi Powder*"
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
               DataField       =   "kuantitas"
               Caption         =   "Kuantitas (Kg)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "operator"
               Caption         =   "Operator"
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   720
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
               BeginProperty Column07 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid GridMilling 
            Height          =   2580
            Left            =   -74880
            TabIndex        =   31
            Top             =   1290
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   4551
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   3
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "grup"
               Caption         =   "Grup"
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
               DataField       =   "namaMesin"
               Caption         =   "Nama Mesin"
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
               DataField       =   "tanggal_mulai"
               Caption         =   "Tgl && Waktu Mulai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy hh:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "tanggal_selesai"
               Caption         =   "Tgl && Waktu Selesai"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd MMM yyyy hh:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "mesh100"
               Caption         =   "Total Mesh 100 (Kg)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "mesh150"
               Caption         =   "Total Mesh 150 (Kg)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "total"
               Caption         =   "Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "rata2"
               Caption         =   "Rata2 per Jam"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
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
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2160
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid GridDetail 
            Height          =   4245
            Left            =   5685
            TabIndex        =   45
            Top             =   720
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   7488
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
               DataField       =   "noEkstraksi"
               Caption         =   "No Ekstraksi"
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
               DataField       =   "kuantitas"
               Caption         =   "Kuantitas (Kg)"
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
               BeginProperty Column00 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker tgl 
            DataField       =   "tanggal_mulai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   11
            Tag             =   "mixing"
            Top             =   3135
            Width           =   2775
            _ExtentX        =   4895
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60096515
            CurrentDate     =   39426
         End
         Begin MSComCtl2.DTPicker tgl 
            DataField       =   "tanggal_selesai"
            DataSource      =   "DDE"
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   12
            Tag             =   "mixing"
            Top             =   3495
            Width           =   2775
            _ExtentX        =   4895
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   60096515
            CurrentDate     =   39426
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7695
            TabIndex        =   54
            Top             =   5040
            Width           =   2640
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   8160
            X2              =   5670
            Y1              =   5370
            Y2              =   5370
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By"
            Height          =   255
            Index           =   2
            Left            =   5670
            TabIndex        =   55
            Top             =   5100
            Width           =   1920
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Produk"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   53
            Top             =   1425
            Width           =   2190
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   2670
            X2              =   240
            Y1              =   1695
            Y2              =   1695
         End
         Begin VB.Line Line1 
            Index           =   10
            Visible         =   0   'False
            X1              =   2670
            X2              =   240
            Y1              =   1335
            Y2              =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Manufature Order"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   52
            Top             =   1065
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Chips sebelum diblender                       Kg"
            Height          =   255
            Index           =   19
            Left            =   255
            TabIndex        =   51
            Top             =   2505
            Width           =   3360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal dan waktu selesai"
            Height          =   255
            Index           =   18
            Left            =   255
            TabIndex        =   50
            Top             =   3525
            Width           =   1920
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   2745
            X2              =   255
            Y1              =   3795
            Y2              =   3795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal dan waktu mulai"
            Height          =   255
            Index           =   17
            Left            =   255
            TabIndex        =   49
            Top             =   3180
            Width           =   1920
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2745
            X2              =   255
            Y1              =   3435
            Y2              =   3435
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   2745
            X2              =   255
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Chips sesudah diblender                       Kg"
            Height          =   255
            Index           =   16
            Left            =   255
            TabIndex        =   48
            Top             =   2850
            Width           =   3855
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   2745
            X2              =   255
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre Lot Chips No"
            Height          =   255
            Index           =   15
            Left            =   255
            TabIndex        =   47
            Top             =   1785
            Width           =   1290
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Grup"
            Height          =   255
            Index           =   14
            Left            =   255
            TabIndex        =   46
            Top             =   2160
            Width           =   480
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   2685
            X2              =   255
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2715
            X2              =   255
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   -72405
            X2              =   -74850
            Y1              =   5265
            Y2              =   5265
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil Milling Powder"
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
            Left            =   -74835
            TabIndex        =   44
            Top             =   4005
            Width           =   1830
         End
         Begin VB.Label lblMesh100 
            BackStyle       =   0  'Transparent
            Caption         =   "PL______(Mesh 100)   :"
            Height          =   255
            Left            =   -74850
            TabIndex        =   43
            Top             =   4290
            Width           =   2025
         End
         Begin VB.Label lblMesh150 
            BackStyle       =   0  'Transparent
            Caption         =   "PL______(Mesh 150)   :"
            Height          =   255
            Left            =   -74850
            TabIndex        =   42
            Top             =   4665
            Width           =   2025
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   4
            Left            =   -71445
            TabIndex        =   41
            Top             =   4290
            Width           =   345
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   5
            Left            =   -71430
            TabIndex        =   40
            Top             =   4635
            Width           =   345
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sak"
            Height          =   255
            Index           =   6
            Left            =   -69840
            TabIndex        =   39
            Top             =   4290
            Width           =   345
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sak"
            Height          =   255
            Index           =   7
            Left            =   -69840
            TabIndex        =   38
            Top             =   4650
            Width           =   345
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   -72420
            X2              =   -74865
            Y1              =   4545
            Y2              =   4545
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   -72405
            X2              =   -74850
            Y1              =   4905
            Y2              =   4905
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total                            "
            Height          =   255
            Index           =   8
            Left            =   -74850
            TabIndex        =   37
            Top             =   5040
            Width           =   2025
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   255
            Index           =   11
            Left            =   -71430
            TabIndex        =   36
            Top             =   4995
            Width           =   345
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sak"
            Height          =   255
            Index           =   12
            Left            =   -69840
            TabIndex        =   35
            Top             =   5010
            Width           =   345
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
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
            Index           =   13
            Left            =   -69135
            TabIndex        =   34
            Top             =   3975
            Width           =   1830
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   -72390
            X2              =   -74835
            Y1              =   1035
            Y2              =   1035
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre Lot Powder No"
            Height          =   255
            Index           =   0
            Left            =   -74835
            TabIndex        =   33
            Top             =   780
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "* Kondisi Powder Setelah Melewati mesin Magnet "
            Height          =   285
            Left            =   -74880
            TabIndex        =   30
            Top             =   5235
            Width           =   3885
         End
         Begin VB.Label Label3 
            Caption         =   "* Kondisi Powder Setelah Melewati mesin Magnet "
            Height          =   285
            Left            =   -74880
            TabIndex        =   27
            Top             =   5235
            Width           =   3885
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sak"
      Height          =   255
      Index           =   10
      Left            =   1590
      TabIndex        =   26
      Top             =   15
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "frmMixingMilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsMixing As New DBQuick
Private RsMilling As New DBQuick
Private RsMesh100 As New DBQuick
Private RsMesh150 As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RsLookup As New DBQuick
Dim tabel As String
Private MEdit As Boolean


Private Sub LoadExtraksi()
   RsLookup.DBOpen "select * from StatusProduksi where posisi='CRUSHER' and status=1", CNN
   If RsLookup.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      mCall.FromTagActive = "No Ekstraksi"
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
      DDE.ChildRecordset.MoveLast
      DDE.ChildRecordset.Delete
   End If
End Sub




Private Sub cmdLink_Click(Index As Integer)
   Select Case Index
      Case 0: RsLookup.DBOpen "select OrderID,OrderName,Type,RequireDate from [MAnufacture Order] where status='RELEASED'", CNN
      Case 1: RsLookup.DBOpen "Select sl_no,creation_date from item_tracking_list where sl_code='prelot' and blocked=0", CNN
      Case 2: RsLookup.DBOpen "select NoItem, ItemName as [Nama Barang] from Inventory where left(Noitem,2) ='FG'", CNN
   End Select
   If RsLookup.DBRecordset.Recordcount > 0 Then
      Set mCall.FormData = RsLookup.DBRecordset
      Select Case Index
         Case 0: mCall.FromTagActive = "Manufacture Order"
         Case 1: mCall.FromTagActive = "Prelot No"
         Case 2: mCall.FromTagActive = "Produk"
      End Select
   Else
      MessageBox "Data Tidak Tersedia", "Peringatan", msgOkOnly, msgCrtical
   End If
End Sub

Private Sub DDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
   Case tmbAddNew:
      With DDE
         RsLookup.DBOpen "Select sl_no from item_tracking_list where sl_code='prelot' and blocked=0 order by sl_no ", CNN
         If RsLookup.DBRecordset.Recordcount > 0 Then
            .GetFieldByName("prelot") = RsLookup.DBRecordset.Fields(0)
         Else
            MessageBox "Data Stock Prelot Tidak ditemukan", "Peringatan", msgOkOnly, msgCrtical
         End If
         .GetFieldByName("grup") = "-"
         .GetFieldByName("total_sebelum") = 0
         .GetFieldByName("total_sesudah") = 0
         .GetFieldByName("tanggal_mulai") = Now
         .GetFieldByName("tanggal_selesai") = Now
         .GetFieldByName("kondisiMagnet") = " "
         .GetFieldByName("desc") = "-"
         .GetFieldByName("sakMesh100") = 0
         .GetFieldByName("sakMesh150") = 0
         tgl(0).Value = Now
         tgl(1).Value = Now
      End With
   Case tmbSave:
      If DDE.IsChildMemberReady = True Then
         SendDataToServer "insert into StatusProduksi (NoEkstraksi,posisi,status,tanggal) values ('" & _
                           txtMixing(0).Text & "','MIXING',1,'" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "')"
         SimpanDetail
         UpdatePrelot
         'MakeStock
      End If
   Case tmbDetail
      Select Case SSTab1.Tab
         Case 0
            LoadExtraksi
         Case 1
            DDE.ChildRecordset.Fields("tanggal_mulai") = Now
            DDE.ChildRecordset.Fields("tanggal_selesai") = Now
            DDE.ChildRecordset.Fields("mesh100") = 0
            DDE.ChildRecordset.Fields("mesh150") = 0
            DDE.ChildRecordset.Fields("total") = 0
            DDE.ChildRecordset.Fields("rata2") = 0
         Case 2
            DDE.ChildRecordset.Fields("L80") = 0
            DDE.ChildRecordset.Fields("P80") = 0
            DDE.ChildRecordset.Fields("P100") = 0
            DDE.ChildRecordset.Fields("total") = 0
            DDE.ChildRecordset.Fields("kuantitas") = 0
         Case 3
            DDE.ChildRecordset.Fields("L80") = 0
            DDE.ChildRecordset.Fields("P80") = 0
            DDE.ChildRecordset.Fields("P100") = 0
            DDE.ChildRecordset.Fields("total") = 0
            DDE.ChildRecordset.Fields("kuantitas") = 0
      End Select
   Case tmbPrint:
      Dim lRep As New utility
      lRep.CallReportView "select * from view_milling where prelot='" & txtMixing(0).Text & "'", "mixing_and _milling.rpt", ReportPath, "Catatan Proses Mixing"
End Select

End Sub

Private Sub UpdatePrelot()
   SendDataToServer "update item_tracking_list set blocked=1 where sl_no='" & txtMixing(0).Text & "'"
End Sub


Private Sub MakeStock()
   If Not MEdit Then
      SendDataToServer "insert into [inventory tabel] (NoIdx,NoItem,Qty_In,stockTmp,typeTrans,ln_no) values (newID(),'" & _
                                                      DDE.GetFieldByName("NoItem") & "'," & FQty(DDE.GetFieldByName("total_sesudah")) & _
                                                      "," & FQty(DDE.GetFieldByName("total_sesudah")) & ",'IP','" & DDE.GetFieldByName("prelot") & "')"
   End If
End Sub


Function simpan()

   DDE.PrepareAppend = "insert into mixing_header (prelot, grup, total_sebelum,total_sesudah,tanggal_mulai, " & _
                       " tanggal_selesai,  kondisiMagnet, [desc], sakMesh100, sakMesh150,issued_by) values ('" & _
                        txtMixing(0) & "','" & txtMixing(1) & "', " & FQty(txtMixing(2)) & ", " & _
                        FQty(txtMixing(3)) & ",'" & Format(tgl(0).Value, "yyyy-MM-dd h:mm:ss") & "','" & _
                        Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "','" & IIf(OpKondisi(0).Value = True, "Bersih", "Kotor") & "','" & _
                        txtMilling(7) & "', " & FQty(txtMilling(4)) & "," & FQty(txtMilling(5)) & ",'" & MainMenu.StatusBar1.Panels(1) & "')"
                           
   DDE.PrepareUpdate = " update mixing_header set grup = '" & txtMixing(1).Text & "',total_sebelum = " & FQty(txtMixing(2)) & _
                       ", total_sesudah = " & FQty(txtMixing(3)) & ", tanggal_mulai = '" & Format(tgl(0).Value, "yyyy-MM-dd hh:mm:ss") & _
                       "', tanggal_selesai = '" & Format(tgl(1).Value, "yyyy-MM-dd hh:mm:ss") & "', kondisiMagnet ='" & _
                       IIf(OpKondisi(0).Value = True, "Bersih", "Kotor") & "', [desc] ='" & txtMilling(7) & _
                       "', sakMesh100 =" & FQty(txtMilling(4)) & ",sakMesh150=" & FQty(txtMilling(5)) & _
                       " where prelot = '" & txtMixing(0) & "'"

End Function


Private Sub DDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   Select Case AdReasonActiveDb
      Case tmbDelete:
         If SSTab1.Tab = 0 Then
            SendDataToServer "update statusProduksi set posisi='CRUSHER',status=1 where NoEkstraksi='" & DDE.ChildRecordset.Fields("noEkstraksi") & "'"
         End If
   End Select
End Sub

Private Sub DDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   loadDetail SSTab1.Tab
   Label4.Caption = IIf(IsNull(DDE.GetFieldByName("approved_by")), "", DDE.GetFieldByName("approved_by"))
End Sub

Private Sub loadDetail(aTab As Integer)
Dim jmlM100 As Double
Dim jmlM150 As Double
RsMixing.DBOpen "select noEkstraksi,kuantitas from mixing_detail where prelot='" & DDE.GetFieldByName("prelot") & "'", CNN
RsMilling.DBOpen "select namaMesin,grup,tanggal_mulai,tanggal_selesai,mesh100,mesh150,total,rata2 from MillingPowder where prelot ='" & DDE.GetFieldByName("prelot") & "'", CNN
RsMesh100.DBOpen "select no,L80,P80,P100,Total,KondisiPowder,kuantitas,operator from MillingMesh where prelot ='" & DDE.GetFieldByName("prelot") & "' and mesh = '100'", CNN
RsMesh150.DBOpen "select no,L80,P80,P100,Total,KondisiPowder,kuantitas,operator from MillingMesh where prelot ='" & DDE.GetFieldByName("prelot") & "' and mesh = '150'", CNN

With DDE.ChildRecordset
   Select Case aTab
      Case 0
         Set DDE.ChildRecordset = RsMixing.DBRecordset
         Set gridDetail.DataSource = DDE.ChildRecordset
      Case 1
         Set DDE.ChildRecordset = RsMilling.DBRecordset
         Set GridMilling.DataSource = DDE.ChildRecordset
         If DDE.ChildRecordset.Recordcount > 0 Then
            jmlM100 = 0
            jmlM150 = 0
            While Not DDE.ChildRecordset.EOF
               jmlM100 = jmlM100 + Val(DDE.ChildRecordset.Fields("Mesh100"))
               jmlM150 = jmlM150 + Val(DDE.ChildRecordset.Fields("Mesh150"))
               DDE.ChildRecordset.MoveNext
            Wend
            txtMilling(1).Text = jmlM100
            txtMilling(2).Text = jmlM150
            txtMilling(3).Text = jmlM100 + jmlM150
            txtMilling(6).Text = Val(txtMilling(4)) + Val(txtMilling(5))
            If DDE.GetFieldByName("KondisiMagnet") = "Bersih" Then
               OpKondisi(0).Value = True
               OpKondisi(1).Value = False
            Else
               OpKondisi(0).Value = False
               OpKondisi(1).Value = True
            End If
         End If
      Case 2
         Set DDE.ChildRecordset = RsMesh100.DBRecordset
         Set GridMesh100.DataSource = DDE.ChildRecordset
      Case 3
         Set DDE.ChildRecordset = RsMesh150.DBRecordset
         Set GridMesh150.DataSource = DDE.ChildRecordset
   End Select
End With
End Sub

Private Sub DDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
cmdLink(0).Enabled = False
cmdLink(1).Enabled = False
cmdLink(2).Enabled = False
Select Case AdReasonActiveDb
   Case tmbAddNew:
      cmdLink(0).Enabled = True
      cmdLink(1).Enabled = True
      cmdLink(2).Enabled = True
      MEdit = False
   Case tmbEdit:
      cmdLink(0).Enabled = True
      cmdLink(1).Enabled = True
      cmdLink(2).Enabled = True
      MEdit = True
   Case tmbSave:
       If RsMixing.DBRecordset.Recordcount > 0 Then
         DDE.IsChildMemberReady = True
       Else
         DDE.IsChildMemberReady = False
         MessageBox "Belum Ada No Ekstraksi yang dimasukkan", "Peringatan", msgOkOnly, msgCrtical
       End If
       simpan
       
   Case tmbDelete:
      Dim rsCek As New DBQuick
      rsCek.DBOpen "select lockFIFO,Qty_out from [Inventory tabel] where noItem='" & DDE.GetFieldByName("NoItem") & "' and ln_no ='" & DDE.GetFieldByName("prelot") & "'"
      If rsCek.DBRecordset.Recordcount > 0 Then
         If rsCek.DBRecordset.Fields(1) = 0 Then
             DDE.PrepareDelete = "delete from Mixing_header where prelot = '" & DDE.GetFieldByName("prelot") & "'"
         Else
            MessageBox "Data Tidak Bisa Dihapus", "Error", msgOkOnly, msgExclamation
            DDE.CancelTrans = True
         End If
      Else
         MessageBox "Data Tidak Bisa Dihapus", "Error", msgOkOnly, msgCrtical
         DDE.CancelTrans = True
      End If
      
End Select
End Sub



Private Sub DGDETAIL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DDE.ChildRecordset.AddNew
End If
End Sub

Private Sub Form_Load()
Set mCall = New frmCaller
With DDE
Set .BindForm = Me
    .BindFormTAG = "mixing"
Set .ActiveConnection = CNN
    .PrepareQuery = " select mixing_header.*,inventory.itemName as itemName from mixing_header left outer join inventory on mixing_header.noItem = inventory.noItem "
End With
HiasFormManTell Picture2, Me
GridMilling.RowHeight = 300
GridMilling.HeadLines = 3
End Sub

Function SimpanDetail()

'*** Save to Mixing Detil ***'
With RsMixing.DBRecordset
   If .Recordcount > 0 Then
      If SendDataToServer("delete from mixing_detail where prelot='" & txtMixing(0) & "'") = True Then
         .MoveFirst
         While Not .EOF
            SendDataToServer "Insert into Mixing_detail (prelot,noEkstraksi,kuantitas) values ('" & _
                             txtMixing(0) & "','" & .Fields("NoEkstraksi") & "'," & _
                             FQty(.Fields("kuantitas")) & ")"
            '*** Update Status produksi
            SendDataToServer "update StatusProduksi set posisi='EKSMIXING',status=1,tanggal='" & _
                             Format(Now, "yyyy-MM-dd hh:mm:ss") & "' where NoEkstraksi = '" & .Fields("NoEkstraksi") & "'"
            .MoveNext
         Wend
      End If
   End If
End With

'*** Save to MillingPowder ***'
With RsMilling.DBRecordset
   If .Recordcount > 0 Then
      If SendDataToServer("delete from MillingPowder where prelot ='" & txtMixing(0) & "'") = True Then
         .MoveFirst
         While Not .EOF
            SendDataToServer "insert into MillingPowder (prelot,namaMesin,grup,tanggal_mulai,tanggal_selesai,Mesh100,mesh150,total,rata2) values ('" & _
                             txtMixing(0) & "','" & .Fields("NamaMesin") & "','" & .Fields("grup") & _
                             "','" & Format(.Fields("tanggal_mulai"), "yyyy-MM-dd hh:mm:ss") & _
                             "','" & Format(.Fields("tanggal_selesai"), "yyyy-MM-dd hh:mm:ss") & _
                             "', " & FQty(.Fields("Mesh100")) & "," & FQty(.Fields("MEsh150")) & _
                             " , " & FQty(.Fields("total")) & "," & FQty(.Fields("rata2")) & ")"
            .MoveNext
         Wend
      End If
   End If
End With

'*** Save to MillingMesh ***'
With RsMesh100.DBRecordset
   If .Recordcount > 0 Then
      If SendDataToServer(" delete from MillingMesh where prelot ='" & txtMixing(0) & "' and Mesh ='100'") Then
         .MoveFirst
         While Not .EOF
            SendDataToServer " insert into MillingMesh (prelot,L80,P80,P100,Total,kondisiPowder,kuantitas,operator,Mesh) values ('" & _
                              txtMixing(0) & "'," & FQty(.Fields("L80")) & "," & FQty(.Fields("P80")) & "," & _
                              FQty(.Fields("P100")) & "," & FQty(.Fields("total")) & ",'" & .Fields("KondisiPowder") & _
                              "'," & FQty(.Fields("kuantitas")) & ",'" & .Fields("operator") & "','100')"
            .MoveNext
         Wend
      End If
   End If
End With

'*** Save to MillingMesh ***'
With RsMesh150.DBRecordset
   If .Recordcount > 0 Then
      If SendDataToServer(" delete from MillingMesh where prelot ='" & txtMixing(0) & "' and Mesh ='150'") Then
         .MoveFirst
         While Not .EOF
            SendDataToServer " insert into MillingMesh (prelot,L80,P80,P100,Total,kondisiPowder,kuantitas,operator,Mesh) values ('" & _
                              txtMixing(0) & "'," & FQty(.Fields("L80")) & "," & FQty(.Fields("P80")) & "," & _
                              FQty(.Fields("P100")) & "," & FQty(.Fields("total")) & ",'" & .Fields("KondisiPowder") & _
                              "'," & FQty(.Fields("kuantitas")) & ",'" & .Fields("operator") & "','150')"
            .MoveNext
         Wend
      End If
   End If
End With

End Function


Private Sub GridMilling_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   viewDate.Visible = False
   Select Case GridMilling.col
      Case 2
         viewDate.Value = DDE.ChildRecordset.Fields("tanggal_mulai")
         viewDate.Move GridMilling.Left + GridMilling.Columns(2).Left, _
                       GridMilling.Top + GridMilling.RowTop(GridMilling.row), _
                       GridMilling.Columns(2).width, _
                       GridMilling.RowHeight
         viewDate.Visible = True
      Case 3
         viewDate.Value = DDE.ChildRecordset.Fields("tanggal_selesai")
         viewDate.Move GridMilling.Left + GridMilling.Columns(3).Left, _
                       GridMilling.Top + GridMilling.RowTop(GridMilling.row), _
                       GridMilling.Columns(3).width, _
                       GridMilling.RowHeight
         viewDate.Visible = True
      Case 4, 5
         DDE.ChildRecordset.Fields("total") = Val(DDE.ChildRecordset.Fields("mesh150")) + Val(DDE.ChildRecordset.Fields("mesh100"))
         
   End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
   Select Case UCase(mCall.FromTagActive)
      Case "NO EKSTRAKSI":
         With DDE.ChildRecordset
            .Fields("NoEkstraksi") = mCall.GetFieldByName(0)
            .Fields("kuantitas") = 0
         End With
      Case "MANUFACTURE ORDER":
         txtMixing(4).Text = mCall.GetFieldByName(0)
      Case "PRELOT NO":
         DDE.GetFieldByName("prelot") = mCall.GetFieldByName(0)
      Case "PRODUK":
         DDE.GetFieldByName("noItem") = mCall.GetFieldByName(0)
         DDE.GetFieldByName("ItemName") = mCall.GetFieldByName(1)
   End Select
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim jmlM100 As Double
Dim jmlM150 As Double

   With DDE.ChildRecordset
      Select Case SSTab1.Tab
         Case 0
            Set DDE.ChildRecordset = RsMixing.DBRecordset
            Set gridDetail.DataSource = DDE.ChildRecordset
         Case 1
            Set DDE.ChildRecordset = RsMilling.DBRecordset
            Set GridMilling.DataSource = DDE.ChildRecordset
            lblMesh100.Caption = "PL " & DDE.GetFieldByName("prelot") & " AF (Mesh 100) "
            lblMesh150.Caption = "PL " & DDE.GetFieldByName("prelot") & " AB (Mesh 150) "
            
            If RsMilling.DBRecordset.Recordcount > 0 Then
               RsMilling.DBRecordset.MoveFirst
               jmlM100 = 0
               jmlM150 = 0
               While Not .EOF
                  jmlM100 = jmlM100 + Val(RsMilling.DBRecordset.Fields("Mesh100"))
                  jmlM150 = jmlM150 + Val(RsMilling.DBRecordset.Fields("Mesh150"))
                  .MoveNext
               Wend
               txtMilling(1).Text = jmlM100
               txtMilling(2).Text = jmlM150
               txtMilling(3).Text = jmlM100 + jmlM150
               txtMilling(6).Text = Val(txtMilling(4)) + Val(txtMilling(5))
               If DDE.GetFieldByName("KondisiMagnet") = "Bersih" Then
                  OpKondisi(0).Value = True
                  OpKondisi(1).Value = False
               Else
                  OpKondisi(0).Value = False
                  OpKondisi(1).Value = True
               End If
            End If
         Case 2
            Set DDE.ChildRecordset = RsMesh100.DBRecordset
            Set GridMesh100.DataSource = DDE.ChildRecordset
         Case 3
            Set DDE.ChildRecordset = RsMesh150.DBRecordset
            Set GridMesh150.DataSource = DDE.ChildRecordset
      End Select
   End With
End Sub


Private Sub viewDate_Change()
On Error GoTo xErr
   Select Case GridMilling.col
      Case 2: DDE.ChildRecordset.Fields("tanggal_mulai") = viewDate.Value
      Case 3: DDE.ChildRecordset.Fields("tanggal_selesai") = viewDate.Value
   End Select
   DDE.ChildRecordset.Fields("rata2") = Val(DDE.ChildRecordset.Fields("total")) / Val(SelisihHariJam(DDE.ChildRecordset.Fields("tanggal_mulai"), DDE.ChildRecordset.Fields("tanggal_selesai"), 2))
Exit Sub
xErr:
   If Err.Number = 11 Then
      DDE.ChildRecordset.Fields("rata2") = 0
      Err.Clear
   Else
      MessageBox Err.Description, "Stop", msgOkOnly, msgExclamation
   End If
End Sub
