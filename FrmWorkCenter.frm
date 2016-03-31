VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMOrder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order"
   ClientHeight    =   6780
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmWorkCenter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Tag             =   "Manufacturing Order"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   6195
      Left            =   0
      ScaleHeight     =   6195
      ScaleWidth      =   11700
      TabIndex        =   26
      Top             =   0
      Width           =   11700
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "no_rekomendasi"
         Height          =   315
         Index           =   5
         Left            =   7035
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   60
         Tag             =   "Partner"
         Top             =   1830
         Width           =   2310
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9345
         Picture         =   "FrmWorkCenter.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1845
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "EmpID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   7035
         MaxLength       =   25
         TabIndex        =   9
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   1500
         Width           =   3345
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4935
         Picture         =   "FrmWorkCenter.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   795
         Width           =   330
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataField       =   "StatusOrder"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "FrmWorkCenter.frx":6F66
         Left            =   7035
         List            =   "FrmWorkCenter.frx":6F7F
         TabIndex        =   8
         Tag             =   "Partner"
         Text            =   "Combo1"
         Top             =   1140
         Width           =   3345
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "No Kontrak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   7035
         MaxLength       =   25
         TabIndex        =   7
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   788
         Width           =   3345
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "QTY Order"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0;(#.##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   1485
         Width           =   2145
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Nama Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "Partner"
         Text            =   "Text1"
         Top             =   780
         Width           =   3585
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataField       =   "Tipe Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "FrmWorkCenter.frx":6FBF
         Left            =   1350
         List            =   "FrmWorkCenter.frx":6FC9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Partner"
         Top             =   60
         Width           =   2190
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Catatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   1
         Left            =   1350
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Partner"
         Text            =   "FrmWorkCenter.frx":6FE5
         Top             =   1845
         Width           =   3945
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3330
         Left            =   75
         TabIndex        =   11
         Top             =   2745
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   5874
         _Version        =   393216
         Style           =   1
         Tabs            =   9
         TabsPerRow      =   9
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
         TabCaption(0)   =   "Order Schedule"
         TabPicture(0)   =   "FrmWorkCenter.frx":6FEB
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Work Centers"
         TabPicture(1)   =   "FrmWorkCenter.frx":7007
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture4(0)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Material"
         TabPicture(2)   =   "FrmWorkCenter.frx":7023
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Routing"
         TabPicture(3)   =   "FrmWorkCenter.frx":703F
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture1"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Cost"
         TabPicture(4)   =   "FrmWorkCenter.frx":705B
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Picture6"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Bill of Manufacture"
         TabPicture(5)   =   "FrmWorkCenter.frx":7077
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txt(1)"
         Tab(5).Control(1)=   "txt(4)"
         Tab(5).Control(2)=   "txt(0)"
         Tab(5).Control(3)=   "txt(3)"
         Tab(5).Control(4)=   "txt(5)"
         Tab(5).Control(5)=   "txt(2)"
         Tab(5).Control(6)=   "SemeruTree1"
         Tab(5).Control(7)=   "Line1(21)"
         Tab(5).Control(8)=   "Label1(20)"
         Tab(5).Control(9)=   "Line1(20)"
         Tab(5).Control(10)=   "Label1(19)"
         Tab(5).Control(11)=   "Line1(19)"
         Tab(5).Control(12)=   "Label1(18)"
         Tab(5).Control(13)=   "Line1(18)"
         Tab(5).Control(14)=   "Label1(17)"
         Tab(5).Control(15)=   "Line1(17)"
         Tab(5).Control(16)=   "Label1(16)"
         Tab(5).Control(17)=   "lblUOM"
         Tab(5).Control(18)=   "Label1(15)"
         Tab(5).Control(19)=   "Line1(16)"
         Tab(5).ControlCount=   20
         TabCaption(6)   =   "Where Used"
         TabPicture(6)   =   "FrmWorkCenter.frx":7093
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Picture7"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Rekomendasi"
         TabPicture(7)   =   "FrmWorkCenter.frx":70AF
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Picture8"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Raw Material (RL)"
         TabPicture(8)   =   "FrmWorkCenter.frx":70CB
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Picture9"
         Tab(8).ControlCount=   1
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00EAAF6F&
            Height          =   2910
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11370
            TabIndex        =   69
            Top             =   360
            Width           =   11430
            Begin VB.CommandButton btnRemove 
               Caption         =   "Batalkan Permintaan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1575
               TabIndex        =   72
               Tag             =   "Partner"
               Top             =   2505
               Width           =   1905
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "Tambah RL"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   45
               TabIndex        =   71
               Tag             =   "Partner"
               Top             =   2505
               Width           =   1500
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWorkCenter.frx":70E7
               Height          =   2460
               Index           =   4
               Left            =   -15
               TabIndex        =   70
               Top             =   0
               Width           =   11385
               _ExtentX        =   20082
               _ExtentY        =   4339
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
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
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "batch_lot"
                  Caption         =   "Kode Batch"
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
                  DataField       =   "Qty required"
                  Caption         =   "Qty Diminta"
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
                  DataField       =   "UOM"
                  Caption         =   "Satuan"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Qty received"
                  Caption         =   "Qty Diterima"
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
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture8 
            Height          =   2910
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11370
            TabIndex        =   67
            Top             =   360
            Width           =   11430
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWorkCenter.frx":70FC
               Height          =   2745
               Index           =   3
               Left            =   0
               TabIndex        =   68
               Top             =   0
               Width           =   10995
               _ExtentX        =   19394
               _ExtentY        =   4842
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BorderStyle     =   0
               HeadLines       =   2
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
                  DataField       =   "formid"
                  Caption         =   "Proses ID"
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
                  DataField       =   "formname"
                  Caption         =   "Proses"
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
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture7 
            Height          =   2910
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11355
            TabIndex        =   65
            Top             =   360
            Width           =   11415
            Begin MSDataGridLib.DataGrid GrdMOUsed 
               Bindings        =   "FrmWorkCenter.frx":7111
               Height          =   2880
               Left            =   0
               TabIndex        =   66
               Tag             =   "SL"
               Top             =   0
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   5080
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               BackColor       =   16577005
               BorderStyle     =   0
               ForeColor       =   7159830
               HeadLines       =   1
               RowHeight       =   16
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
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   10
               BeginProperty Column00 
                  DataField       =   "No_MO"
                  Caption         =   "No.MO"
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
                  DataField       =   "ItemName"
                  Caption         =   "Nama Barang"
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
                  DataField       =   "Status"
                  Caption         =   "Status"
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
                  DataField       =   "OrderName"
                  Caption         =   "OrderName"
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
                  DataField       =   "Type"
                  Caption         =   "Tipe MO"
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
                  DataField       =   "NoItem"
                  Caption         =   "NoItem"
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
                  DataField       =   "StatusMO"
                  Caption         =   "Status MO"
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
                  DataField       =   "StartDate"
                  Caption         =   "Tgl Mulai"
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
                  DataField       =   "EndDate"
                  Caption         =   "Tgl Selesai"
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
                  DataField       =   "CompanyName"
                  Caption         =   "Klien"
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
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
                  BeginProperty Column02 
                     Object.Visible         =   0   'False
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                  EndProperty
                  BeginProperty Column06 
                  EndProperty
                  BeginProperty Column07 
                  EndProperty
                  BeginProperty Column08 
                  EndProperty
                  BeginProperty Column09 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   2910
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11355
            TabIndex        =   64
            Top             =   360
            Width           =   11415
         End
         Begin VB.PictureBox Picture1 
            Height          =   2910
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11355
            TabIndex        =   62
            Top             =   360
            Width           =   11415
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWorkCenter.frx":7126
               Height          =   2880
               Index           =   0
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   11310
               _ExtentX        =   19950
               _ExtentY        =   5080
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BorderStyle     =   0
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
               ColumnCount     =   11
               BeginProperty Column00 
                  DataField       =   "seqNo"
                  Caption         =   "SeqNo"
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
                  DataField       =   "stageID"
                  Caption         =   "Work Center"
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
                  DataField       =   "Keterangan"
                  Caption         =   "Description"
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
                  DataField       =   "unit_run"
                  Caption         =   "Unit Run (Minutes)"
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
                  DataField       =   "Extended_Run"
                  Caption         =   "Extended Run (Hours)"
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
                  DataField       =   "Setup_Time"
                  Caption         =   "Setup Time (Sec)"
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
                  DataField       =   "Queue_Time"
                  Caption         =   "Queue Time (Sec)"
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
                  DataField       =   "Wait_Time"
                  Caption         =   "Wait Time (Sec)"
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
               BeginProperty Column08 
                  DataField       =   "Total_Run_Time"
                  Caption         =   "Total Run Time (Hours)"
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
               BeginProperty Column09 
                  DataField       =   "Actual_Time"
                  Caption         =   "Actual Time (Hours)"
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
               BeginProperty Column10 
                  DataField       =   "Overlap"
                  Caption         =   "Overlap"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "Ya"
                     FalseValue      =   "Tidak"
                     NullValue       =   ""
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   7
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
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column06 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column07 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column08 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column09 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column10 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Keterangan"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   -68955
            TabIndex        =   21
            Text            =   " - Nama Barang -"
            Top             =   885
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "UOM"
            Height          =   315
            Index           =   4
            Left            =   -68955
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Text            =   " - Jumlah -"
            Top             =   2280
            Width           =   2115
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "BOM Id"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   -68955
            TabIndex        =   20
            Text            =   " - Kode Barang -"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Keterangan"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   -68955
            TabIndex        =   25
            Text            =   " - Kelompok -"
            Top             =   1815
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "UOM"
            Height          =   315
            Index           =   5
            Left            =   -68955
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Text            =   " - Lead Time -"
            Top             =   2745
            Width           =   3000
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            DataField       =   "Keterangan"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   -68955
            TabIndex        =   22
            Text            =   " - Kategori -"
            Top             =   1350
            Width           =   3000
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00EAAF6F&
            Height          =   2910
            Left            =   60
            ScaleHeight     =   2850
            ScaleWidth      =   11370
            TabIndex        =   32
            Top             =   360
            Width           =   11430
            Begin VB.ComboBox Combo2 
               DataField       =   "Priority"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "FrmWorkCenter.frx":713B
               Left            =   1710
               List            =   "FrmWorkCenter.frx":7145
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Tag             =   "Partner"
               Top             =   555
               Width           =   3210
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "CreateDate"
               Height          =   330
               Index           =   0
               Left            =   1710
               TabIndex        =   13
               Tag             =   "Partner"
               Top             =   210
               Width           =   3210
               _ExtentX        =   5662
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
               Format          =   77070339
               CurrentDate     =   38484
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "RequireDate"
               Height          =   330
               Index           =   1
               Left            =   1710
               TabIndex        =   15
               Tag             =   "Partner"
               Top             =   900
               Width           =   3210
               _ExtentX        =   5662
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
               Format          =   77070339
               CurrentDate     =   38484
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "StartDate"
               Height          =   330
               Index           =   3
               Left            =   1710
               TabIndex        =   16
               Tag             =   "Partner"
               Top             =   1245
               Width           =   3210
               _ExtentX        =   5662
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
               Format          =   77070339
               CurrentDate     =   38484
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "FinishedDate"
               Height          =   330
               Index           =   4
               Left            =   1710
               TabIndex        =   17
               Tag             =   "Partner"
               Top             =   1605
               Width           =   3210
               _ExtentX        =   5662
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
               Format          =   77070339
               CurrentDate     =   38484
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "EarliesDate"
               Height          =   330
               Index           =   2
               Left            =   1710
               TabIndex        =   33
               Tag             =   "Partner"
               Top             =   1245
               Visible         =   0   'False
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   582
               _Version        =   393216
               CustomFormat    =   "dd/MMMM/yyyy"
               Format          =   77070339
               CurrentDate     =   38484
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "released_date"
               Height          =   330
               Index           =   6
               Left            =   7500
               TabIndex        =   18
               Tag             =   "Partner"
               Top             =   1260
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
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
               Format          =   77070339
               CurrentDate     =   -106283
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               DataField       =   "closed_date"
               Height          =   330
               Index           =   7
               Left            =   7485
               TabIndex        =   19
               Tag             =   "Partner"
               Top             =   1620
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
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
               Format          =   77070339
               CurrentDate     =   38484
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Released Date"
               Height          =   195
               Index           =   23
               Left            =   6075
               TabIndex        =   58
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Line Line1 
               Index           =   24
               X1              =   6075
               X2              =   7500
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Closed Date"
               Height          =   195
               Index           =   22
               Left            =   6075
               TabIndex        =   57
               Top             =   1665
               Width           =   870
            End
            Begin VB.Line Line1 
               Index           =   23
               X1              =   6075
               X2              =   7500
               Y1              =   1575
               Y2              =   1575
            End
            Begin VB.Line Line1 
               Index           =   11
               X1              =   300
               X2              =   1725
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Finished Date"
               Height          =   195
               Index           =   10
               Left            =   300
               TabIndex        =   39
               Top             =   1650
               Width           =   975
            End
            Begin VB.Line Line1 
               Index           =   10
               X1              =   300
               X2              =   1725
               Y1              =   1905
               Y2              =   1905
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start Date"
               Height          =   195
               Index           =   9
               Left            =   300
               TabIndex        =   38
               Top             =   1305
               Width           =   750
            End
            Begin VB.Line Line1 
               Index           =   9
               Visible         =   0   'False
               X1              =   1845
               X2              =   3270
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Earliest Date"
               Height          =   210
               Index           =   8
               Left            =   1845
               TabIndex        =   37
               Top             =   1305
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Line Line1 
               Index           =   8
               X1              =   300
               X2              =   1725
               Y1              =   1215
               Y2              =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Require Date"
               Height          =   195
               Index           =   7
               Left            =   300
               TabIndex        =   36
               Top             =   960
               Width           =   945
            End
            Begin VB.Line Line1 
               Index           =   7
               X1              =   300
               X2              =   1725
               Y1              =   870
               Y2              =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Priority"
               Height          =   195
               Index           =   6
               Left            =   300
               TabIndex        =   35
               Top             =   615
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order Date"
               Height          =   195
               Index           =   5
               Left            =   300
               TabIndex        =   34
               Top             =   300
               Width           =   810
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   300
               X2              =   1725
               Y1              =   525
               Y2              =   525
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00EAAF6F&
            Height          =   2910
            Index           =   0
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11370
            TabIndex        =   29
            Top             =   360
            Width           =   11430
            Begin MSComCtl2.DTPicker TglStart 
               Height          =   315
               Left            =   5805
               TabIndex        =   30
               Top             =   240
               Visible         =   0   'False
               Width           =   1740
               _ExtentX        =   3069
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
               CustomFormat    =   "dd/MM/yy hh:mm"
               Format          =   77070339
               CurrentDate     =   36494
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWorkCenter.frx":7159
               Height          =   2880
               Index           =   1
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   11370
               _ExtentX        =   20055
               _ExtentY        =   5080
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   21
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
                  DataField       =   "SeqNo"
                  Caption         =   "No"
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
                  DataField       =   "StageID"
                  Caption         =   "Stage"
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
                  DataField       =   "Keterangan"
                  Caption         =   "Keterangan"
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
                  DataField       =   "StartDate"
                  Caption         =   "Start Date"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "m/d/yy h:nn AM/PM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   4
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "EndDate"
                  Caption         =   "End Date"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "m/d/yy h:nn AM/PM"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   3
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "ResourcesID"
                  Caption         =   "Resources"
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
                  DataField       =   "Status"
                  Caption         =   "Status"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   5
                     Format          =   ""
                     HaveTrueFalseNull=   1
                     TrueValue       =   "Closed"
                     FalseValue      =   "Open"
                     NullValue       =   "Open"
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   7
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "Warehouse"
                  Caption         =   "GD. Tujuan"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
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
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   2
                  EndProperty
                  BeginProperty Column05 
                  EndProperty
                  BeginProperty Column06 
                  EndProperty
                  BeginProperty Column07 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00EAAF6F&
            Height          =   2910
            Left            =   -74940
            ScaleHeight     =   2850
            ScaleWidth      =   11370
            TabIndex        =   27
            Top             =   360
            Width           =   11430
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmWorkCenter.frx":716E
               Height          =   2880
               Index           =   2
               Left            =   -15
               TabIndex        =   28
               Top             =   -15
               Width           =   11385
               _ExtentX        =   20082
               _ExtentY        =   5080
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               BorderStyle     =   0
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
               ColumnCount     =   17
               BeginProperty Column00 
                  DataField       =   "SeqNo"
                  Caption         =   "SeqNo"
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
                  DataField       =   "Work Center"
                  Caption         =   "Work Center"
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
                  DataField       =   "Routing"
                  Caption         =   "Routing"
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
                  DataField       =   "Kode Barang"
                  Caption         =   "Kode Barang"
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
                  DataField       =   "Nama Barang"
                  Caption         =   "Nama Barang"
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
               BeginProperty Column06 
                  DataField       =   "Quote Qty"
                  Caption         =   "Quote Qty"
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
                  DataField       =   "Actual Qty"
                  Caption         =   "Actual Qty"
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
                  DataField       =   "scrap_qty"
                  Caption         =   "scrap_qty"
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
                  DataField       =   "wip_qty"
                  Caption         =   "wip_qty"
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
                  DataField       =   "completed_qty"
                  Caption         =   "completed_qty"
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
                  DataField       =   "Partner ID"
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
               BeginProperty Column12 
                  DataField       =   "Nama Perusahaan"
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
               BeginProperty Column13 
                  DataField       =   "No PO"
                  Caption         =   "No PO"
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
               BeginProperty Column14 
                  DataField       =   "Phantom"
                  Caption         =   "Phantom"
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
               BeginProperty Column15 
                  DataField       =   "Complete"
                  Caption         =   "Complete"
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
               BeginProperty Column16 
                  DataField       =   "comment"
                  Caption         =   "comment"
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
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                  EndProperty
                  BeginProperty Column06 
                  EndProperty
                  BeginProperty Column07 
                  EndProperty
                  BeginProperty Column08 
                  EndProperty
                  BeginProperty Column09 
                  EndProperty
                  BeginProperty Column10 
                  EndProperty
                  BeginProperty Column11 
                  EndProperty
                  BeginProperty Column12 
                  EndProperty
                  BeginProperty Column13 
                     ColumnWidth     =   720
                  EndProperty
                  BeginProperty Column14 
                  EndProperty
                  BeginProperty Column15 
                  EndProperty
                  BeginProperty Column16 
                  EndProperty
               EndProperty
            End
         End
         Begin SemeruDC.SemeruTree SemeruTree1 
            Height          =   2820
            Left            =   -74895
            TabIndex        =   49
            Top             =   390
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   4974
            BackColorTree   =   7159830
            BackColorBackground=   -2147483648
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   -70185
            X2              =   -68760
            Y1              =   2580
            Y2              =   2580
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            Height          =   195
            Index           =   20
            Left            =   -70185
            TabIndex        =   56
            Top             =   2340
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   -70185
            X2              =   -68760
            Y1              =   2115
            Y2              =   2115
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelompok"
            Height          =   195
            Index           =   19
            Left            =   -70185
            TabIndex        =   55
            Top             =   1875
            Width           =   675
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   -70170
            X2              =   -68745
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
            Height          =   195
            Index           =   18
            Left            =   -70185
            TabIndex        =   54
            Top             =   480
            Width           =   915
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   -70185
            X2              =   -68760
            Y1              =   1185
            Y2              =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
            Height          =   195
            Index           =   17
            Left            =   -70185
            TabIndex        =   53
            Top             =   945
            Width           =   960
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   -70185
            X2              =   -68760
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Time"
            Height          =   195
            Index           =   16
            Left            =   -70185
            TabIndex        =   52
            Top             =   2805
            Width           =   720
         End
         Begin VB.Label lblUOM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PCS"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   -66795
            TabIndex        =   51
            Top             =   2280
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori"
            Height          =   195
            Index           =   15
            Left            =   -70185
            TabIndex        =   50
            Top             =   1410
            Width           =   600
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   -70185
            X2              =   -68760
            Y1              =   1650
            Y2              =   1650
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "CreateDate"
         Height          =   315
         Index           =   5
         Left            =   7035
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   435
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   77070339
         CurrentDate     =   38272
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "Pelanggan"
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Top             =   1140
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "[Nama Perusahaan]"
         BoundColumn     =   "Pelanggan"
         Text            =   "DataCombo1"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rekomendasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   5745
         TabIndex        =   61
         Top             =   1905
         Width           =   1185
      End
      Begin VB.Line Line1 
         Index           =   25
         X1              =   5745
         X2              =   7170
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         DataField       =   "No Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1350
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   420
         Width           =   2145
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5745
         X2              =   7170
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   105
         X2              =   1530
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   48
         Top             =   2370
         Width           =   420
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   105
         X2              =   1530
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   47
         Top             =   120
         Width           =   825
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   105
         X2              =   1530
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   1530
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   46
         Top             =   450
         Width           =   60
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   105
         X2              =   1530
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmWorkCenter.frx":7183
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   45
         Top             =   848
         Width           =   6705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Qty"
         DataField       =   "Qty Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   44
         Top             =   1553
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   105
         X2              =   1530
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   5745
         TabIndex        =   43
         Top             =   495
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Line Line1 
         Index           =   13
         Visible         =   0   'False
         X1              =   5745
         X2              =   7170
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   5745
         X2              =   7170
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         DataField       =   "Qty Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   105
         TabIndex        =   42
         Top             =   488
         Width           =   660
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   5745
         X2              =   7170
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approval By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   5745
         TabIndex        =   41
         Top             =   1545
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmWorkCenter.frx":720B
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   40
         Top             =   1200
         Width           =   6615
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   6210
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmMOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Private mAdd As Boolean
Private RcIssued As New DBQuick
Private RcPart As New DBQuick
Private RcPartner As New DBQuick
Private RcCompDetail As New DBQuick
Private RcCompData As New DBQuick
Private RsTeeBOM As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mFirstCaller, mAddOnly As Boolean
Private PrevStatus As String
Private RsRekomendasi As New DBQuick
Private RsRL As New DBQuick
Dim RcMOWhere As DBQuick

Private Sub cmdEkstraksi_Click()
OpenDetailPartner 7
End Sub

Private Sub btnRemove_Click()
   If RsRL.DBRecordset.Fields("qty received") = 0 Then
      If MessageBox("Data Akan Dihapus ?", "Konfirmasi", msgYesNo, msgQuestion) <> 0 Then
         RsRL.DBRecordset.Delete
      End If
   Else
      MessageBox "Data Tidak Bisa Dibatalkan", "Perhatian"
   End If
End Sub

Private Sub cmdAdd_Click()
   RsRL.DBRecordset.AddNew
   OpenDetailPartner 8
End Sub

Private Sub cmdLink_Click()
   OpenDetailPartner 3
End Sub

'Private Sub cmdLink1_Click()
'   OpenDetailPartner 5
'End Sub

Private Sub Combo1_Click(Index As Integer)
Select Case Index
       Case 0:
            Combo1(1).Clear
            Select Case UCase(Combo1(0).Text)
                   Case "ASSEMBLY ORDER":
                        Combo1(1).AddItem "QUOTED"
                        Combo1(1).AddItem "RELEASED"
                        Combo1(1).AddItem "FINISHED"
'                        Combo1(1).ListIndex = 0
                        Combo1(1).Text = MyDDE.GetFieldByName("StatusOrder")
                        DataCombo1.Enabled = True
                   Case "MAKE ORDER"
                        Combo1(1).AddItem "ORDERED"
                        Combo1(1).AddItem "RELEASED"
                        Combo1(1).AddItem "FINISHED"
                        Combo1(1).AddItem "INVOICED"
                        Combo1(1).AddItem "CLOSED"
                        'Combo1(1).ListIndex = 0
                        Combo1(1).Text = MyDDE.GetFieldByName("StatusOrder")
                        DataCombo1.Enabled = True
                        DataCombo1.Tag = "Partner"
                   Case "MAKE STOCK"
                        Combo1(1).AddItem "NEW"
                        Combo1(1).AddItem "RELEASED"
                        Combo1(1).AddItem "FINISHED"
                        'Combo1(1).ListIndex = 0
                        Combo1(1).Text = MyDDE.GetFieldByName("StatusOrder")
                        DataCombo1.Text = "MAKE TO STOCK"
                        DataCombo1.Enabled = False
                        DataCombo1.Tag = ""
            End Select
       Case 1:
       Case 2:
End Select
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub


Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
If Index = 0 And mAddOnly = True Then
   If Combo1(Index) = "ASSEMBLE ORDER" Then
        MessageBox "Tipe order harus selain ASSEMBLE ORDER", "Peringatan", msgOkOnly
        Cancel = True
   End If
End If
End Sub

Private Sub Command1_Click()
   OpenDetailPartner 6
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataCombo1_Validate(Cancel As Boolean)
If mAdd = True Then
   If Combo1(0).Text = "MAKE ORDER" Then
      If DataCombo1.Text = "MAKE TO STOCK" Then
         MessageBox "Harus diisi selain MAKE TO STOCK.", "Peringatan", msgOkOnly
         Cancel = True
      End If
   End If
End If
End Sub

Private Sub DataGrid1_AfterColEdit(Index As Integer, ByVal ColIndex As Integer)
If mAdd = True And (Combo1(1).Text <> "RELEASED") Or (Combo1(0).Text <> "MAKE STOCK") Then
Select Case Index
       Case 0:
       Case 1:
            Select Case ColIndex:
                   Case 3:
                        If CDate(DataGrid1(Index).Columns(ColIndex)) > CDate(DTPicker1(1).Value) Then
                           MessageBox "Tanggal Mulai Proses Stage tidak boleh lebih besar dari tanggal Require Date", "Peringatan", msgOkOnly
                        End If
                   Case 4:
            End Select
End Select
End If
End Sub

Private Sub DataGrid1_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
Select Case Index
       Case 1:
            Select Case ColIndex
                   Case 6:
                        If CBool(DataGrid1(Index).Columns(6).Value) = True Then
                           DataGrid1(Index).Columns(6).Value = False
                        Else
                           DataGrid1(Index).Columns(6).Value = True
                        End If
                        MyDDE.ChildRecordset.Fields("Status") = CBool(DataGrid1(Index).Columns(6).Value)
                   Case 7: OpenDetailPartner 4
            End Select
End Select
End Sub

Private Sub DataGrid1_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If mAdd = True And (Combo1(1).Text <> "RELEASED") Or (Combo1(0).Text <> "MAKE STOCK") Then
   Select Case DataGrid1(Index).col
          Case 3, 4, 7:
'               DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
'               DataGrid1(Index).AllowUpdate = mAdd
          Case Else
               DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
               DataGrid1(Index).AllowUpdate = False
   End Select
'   MoveCtrl
Else
   DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
   DataGrid1(Index).AllowUpdate = False
'   CmdStage.Enabled = False
'   CmdStage.Visible = False
End If
End Sub

Private Sub DTPicker1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
       Case 0:
       Case 1:
            If DTPicker1(1).Value < DTPicker1(0).Value Then
               MessageBox "Tanggal Required Date tidak boleh lebih kecil dari tanggal order", "Peringatan", msgOkOnly
               DTPicker1(1).Value = DTPicker1(0).Value + 1
               Cancel = True
            End If
       Case 2:

       Case 3:
            If DTPicker1(3).Value < DTPicker1(0).Value Then
               MessageBox "Tanggal Start Date tidak boleh lebih kecil dari tanggal order", "Peringatan", msgOkOnly
               DTPicker1(3).Value = DTPicker1(0).Value + 1
               Cancel = True
            End If
       Case 4:
            If DTPicker1(4).Value < DTPicker1(3).Value Then
               MessageBox "Tanggal Finish Date tidak boleh lebih kecil dari tanggal order", "Peringatan", msgOkOnly
               DTPicker1(4).Value = DTPicker1(3).Value + 1
               Cancel = True
            End If
End Select
End Sub



Private Sub Form_Load()
Combo1(1).Clear
Combo1(0).ListIndex = 0


'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout

DTPicker1(0).Value = Date
DTPicker1(1).Value = Date
DTPicker1(2).Value = Date
DTPicker1(3).Value = Date
DTPicker1(4).Value = Date
SSTab1.Tab = 0
OpenPartner
With MyDDE
    .EditModeReplace = False
    Set .BindForm = Me
    .SetPermissions = UserDeleteDenied
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = " SELECT OrderID AS [No Order], PartnerID AS Pelanggan, OrderName AS [Nama Order], no_rekomendasi," & _
                " Type AS [Tipe Order], [manufacture Order].Status AS StatusOrder, Note AS Catatan, ContractID AS [No Kontrak], " & _
                " CreateDate, Priority, RequireDate, EarliesDate, StartDate AS StartDate, FinishedDate AS FinishedDate, " & _
                " [QTY Order], EmpID, NoItem AS [Kode BOM],[manufacture Order].ekstraksi_no,[manufacture Order].released_date,[manufacture Order].closed_date " & _
                " FROM [Manufacture Order] ORDER BY OrderID"
End With
Set mCall = New frmCaller
'Check1.BackColor = &HEAAF6F
Label2.ForeColor = txtBox(0).ForeColor
DataGrid1(0).HeadLines = 3
LoadTreeBOM
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If MyDDE.CheckRecordPendinged = True Then
'   ScanKey vbKeyF5, 0, MyDDE
'   If MyDDE.IsSucces = True Then
'      Cancel = False
'      MyDDE.ClearRecordset
'      Set FrmWorkCenter = Nothing
'   Else
'      Cancel = True
'   End If
'Else
Set RcIssued = Nothing
Set RcPart = Nothing
Set RcPartner = Nothing
Set RcCompDetail = Nothing
Set RcCompData = Nothing
Set mCall = Nothing
MyDDE.ClearRecordset
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMOrder = Nothing
End Sub

Private Sub mCall_BeforeUnload()
   Select Case UCase(mCall.FromTagActive)
      Case "NO REKOMENDASI":
         RsRekomendasi.DBOpen "Select formid,formname from  view_rekomekstraksi_proses where splno ='" & MyDDE.GetFieldByName("no_rekomendasi") & "'", CNN
         Set DataGrid1(3).DataSource = RsRekomendasi.DBRecordset
   End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
'If MyDDE.ChildRecordset.Recordcount > 0 Then
Select Case UCase(mCall.FromTagActive)
       Case "GUDANG":
            With MyDDE.ChildRecordset
                 .Fields("Warehouse") = mCall.GetFieldByName(0)
            End With
       Case "BOM LISTING":
            With MyDDE.ActiveRecordset
                 .Fields("Nama Order") = mCall.GetFieldByName(1)
                 .Fields("Kode BOM") = mCall.GetFieldByName(0)
                  Dim Rc As New DBQuick
                  Rc.DBOpen "SELECT [BOM Stage Detail].NoLine, [BOM Stage Detail].WCID AS SeqStageID, [BOM Stage Detail].Description AS Keterangan, [BOM Stage Detail].ResourcesID,  [Resources Table].Description AS Resources, [BOM Stage Detail].StageNote AS Catatan FROM         [BOM Stage Detail] INNER JOIN  Inventory ON [BOM Stage Detail].NoItem = Inventory.NoItem AND [BOM Stage Detail].BomReff = Inventory.BomReff LEFT OUTER JOIN   [Resources Table] ON [BOM Stage Detail].ResourcesID = [Resources Table].ResourcesID WHERE (Inventory.NoItem = N'" & mCall.GetFieldByName(0) & "') ORDER BY [BOM Stage Detail].NoLine, [BOM Stage Detail].SeqStageID", CNN, lckLockBatch
'                  Debug.Print "SELECT [BOM Stage Detail].NoLine, [BOM Stage Detail].WCID AS SeqStageID, [BOM Stage Detail].Description AS Keterangan, [BOM Stage Detail].ResourcesID,  [Resources Table].Description AS Resources, [BOM Stage Detail].StageNote AS Catatan FROM         [BOM Stage Detail] INNER JOIN  Inventory ON [BOM Stage Detail].NoItem = Inventory.NoItem AND [BOM Stage Detail].BomReff = Inventory.BomReff LEFT OUTER JOIN   [Resources Table] ON [BOM Stage Detail].ResourcesID = [Resources Table].ResourcesID WHERE (Inventory.NoItem = N'" & mCall.GetFieldByName(0) & "') ORDER BY [BOM Stage Detail].NoLine, [BOM Stage Detail].SeqStageID"
                  With Rc.DBRecordset
                       If .Recordcount > 0 Then
                           OpenDetail "xxx"
                           Do
                             If .EOF Then Exit Do
                             MyDDE.ChildRecordset.AddNew
                             MyDDE.ChildRecordset.Fields("SeqNo") = .Fields(0)
                             MyDDE.ChildRecordset.Fields("StageID") = .Fields(1)
                             MyDDE.ChildRecordset.Fields("Keterangan") = .Fields(2)
                             If Not IsNull(.Fields(3)) Then MyDDE.ChildRecordset.Fields("ResourcesID") = .Fields(3)
                             MyDDE.ChildRecordset.Fields("StartDate") = Date
                             MyDDE.ChildRecordset.Fields("EndDate") = Date
                             MyDDE.ChildRecordset.Fields("Status") = False
                             .MoveNext
                           Loop
                           MyDDE.ChildRecordset.MoveFirst
                       Else
                          MessageBox "Data bom belum lengkap. Transaksi dibatalkan", "Peringatan", msgOkOnly
                          MyDDE.CallButtonActive (tmbCancel)
                          Exit Sub
                       End If
                  End With
                  Dim RcCompData As New DBQuick
                  Dim I As Integer
                  'RcCompData.DBOpen "SELECT     [BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [Kode Barang], Inventory.ItemName AS [Nama Komponen],  [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan],  [BOM Component Detail].QTYUsage FROM         Inventory INNER JOIN  PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID INNER JOIN [BOM Component Detail] ON Inventory.NoItem = [BOM Component Detail].Component WHERE     ([BOM Component Detail].NoItem = N'" & .Fields("Kode Bom") & "') GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID, PartnerDB.CompanyName, [BOM Component Detail].QTYUsage ORDER BY [BOM Component Detail].SeqStageID", Cnn, lckLockReadOnly
'                  RcCompData.DBOpen " SELECT [BOM Component Detail].SeqStageID, [BOM Component Detail].Component AS [Kode Barang], Inventory.ItemName AS [Nama Komponen], [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], [BOM Component Detail].QTYUsage" & _
                                    " FROM [BOM Component Detail] INNER JOIN [BOM Stage Detail] ON [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND [BOM Component Detail].BomReff = [BOM Stage Detail].BomReff AND [BOM Component Detail].WCID = [BOM Stage Detail].WCID INNER JOIN Inventory INNER JOIN   PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].BomReff = Inventory.BomReff AND  [BOM Component Detail].Component = Inventory.NoItem WHERE     ([BOM Component Detail].NoItem = N'" & .Fields("Kode Bom") & "') GROUP BY [BOM Component Detail].SeqStageID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,   PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine", Cnn, lckLockReadOnly
                                    
                 RcCompData.DBOpen " SELECT [BOM Stage Detail].WCID, [BOM Component Detail].Component AS [Kode Barang], Inventory.ItemName AS [Nama Komponen], [BOM Component Detail].UOM, Inventory.PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan],  [BOM Component Detail].QTYUsage, [BOM Component Detail].SeqStageID" & _
                                   " FROM [BOM Component Detail] INNER JOIN [BOM Stage Detail] ON [BOM Component Detail].NoItem = [BOM Stage Detail].NoItem AND [BOM Component Detail].WCID = [BOM Stage Detail].WCID INNER JOIN Inventory INNER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID ON [BOM Component Detail].Component = Inventory.NoItem" & _
                                   " WHERE ([BOM Component Detail].NoItem = N'" & .Fields("Kode Bom") & "') " & _
                                   " GROUP BY [BOM Stage Detail].WCID, [BOM Component Detail].Component, Inventory.ItemName, [BOM Component Detail].UOM, Inventory.PartnerID,PartnerDB.CompanyName, [BOM Component Detail].QTYUsage, [BOM Stage Detail].NoLine,[BOM Component Detail].SeqStageID ORDER BY [BOM Stage Detail].NoLine", CNN, lckLockReadOnly
'                Debug.Print RcCompData.DBRecordset.Source
                  With RcCompData.DBRecordset
                       If .Recordcount > 0 Then
                           I = 0
                           OpenComponentDetail "xxx" 'IIf(Not IsNull(MyDDE.GetFieldByName("No Order")), MyDDE.GetFieldByName("No Order"), "xxx")
                           Do
                             I = I + 1
                             If .EOF Then Exit Do
                             RcCompDetail.DBRecordset.AddNew
                             RcCompDetail.DBRecordset.Fields("SeqNo") = I
                             'MsgBox .Fields(0)
                             RcCompDetail.DBRecordset.Fields("Work Center") = .Fields(0)
                             RcCompDetail.DBRecordset.Fields("Kode Barang") = .Fields(1)
                             RcCompDetail.DBRecordset.Fields("Nama Barang") = .Fields(2)
                             RcCompDetail.DBRecordset.Fields("UOM") = .Fields(3)
                             RcCompDetail.DBRecordset.Fields("Partner ID") = .Fields(4)
                             RcCompDetail.DBRecordset.Fields("Nama Perusahaan") = .Fields(5)
                             RcCompDetail.DBRecordset.Fields("Quote QTY") = .Fields(6)
                             RcCompDetail.DBRecordset.Fields("Actual QTY") = LoadQty(.Fields(1))
                             RcCompDetail.DBRecordset.Fields("No PO") = "-"
                             RcCompDetail.DBRecordset.Fields("Phantom") = 0
                             RcCompDetail.DBRecordset.Fields("Complete") = 0
                             RcCompDetail.DBRecordset.Fields("Routing") = .Fields("SeqStageID").Value
                             RcCompDetail.DBRecordset.Fields("wip_qty") = 0
                             RcCompDetail.DBRecordset.Fields("scrap_qty") = 0
                             RcCompDetail.DBRecordset.Fields("comment") = "-"
                             .MoveNext
                           Loop
                           RcCompDetail.DBRecordset.MoveFirst
                       Else
                          MessageBox "Data bom belum lengkap. Transaksi dibatalkan", "Peringatan", msgOkOnly
                          MyDDE.CallButtonActive (tmbCancel)
                          Exit Sub
                       End If
                  End With
                  
            End With
       Case "MASTER STAGE":
            With MyDDE.ChildRecordset
                 .Fields("Operation") = mCall.GetFieldByName(0)
            End With
      Case "NO REKOMENDASI":
         MyDDE.GetFieldByName("no_rekomendasi") = mCall.GetFieldByName(0)
      Case "DAFTAR RL BATCH":
         RsRL.DBRecordset.Fields("batch_lot") = mCall.GetFieldByName("sl_no")
         RsRL.DBRecordset.Fields("UOM") = mCall.GetFieldByName("uom")
         RsRL.DBRecordset.Fields("qty required") = mCall.GetFieldByName("stockTMp")
         RsRL.DBRecordset.Fields("qty Received") = 0
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim z As Integer
Command1.Enabled = False
Combo1(1).Enabled = False
Select Case AdReasonActiveDb
       Case tmbAddNew:
            Dim IDGen As New IDGenerator
            
            RsRL.DBOpen "select Idtrans, NoItem, batch_lot, UOM, [Qty Required], [Qty Received] from backflush_line where orderID ='xxxxxx'", CNN, lckLockBatch
            Set DataGrid1(4).DataSource = RsRL.DBRecordset
            
            Command1.Enabled = True
            mAdd = True
            txtBox(0).SetFocus
            Label2 = IDGen.GetID("MO")   'IndexAuto
            'txtBox(11).Text = IDGen.GetID("EKS")
            MyDDE.GetFieldByName("QTY Order") = 1
            MyDDE.GetFieldByName("No Kontrak") = "-"
            MyDDE.GetFieldByName("Catatan") = "-"
            MyDDE.GetFieldByName("CreateDate") = Date
            MyDDE.GetFieldByName("RequireDate") = Date
            MyDDE.GetFieldByName("StartDate") = Date
            MyDDE.GetFieldByName("FinishedDate") = Date
            MyDDE.GetFieldByName("EmpID") = GetSetting("Manufacturing Intelligent", "Server", "UserName Active")
            
            DTPicker1(3).Enabled = False
            DTPicker1(4).Enabled = False
            Picture5.Enabled = False
            txtBox(0).Locked = True
            mAddOnly = True
            Combo2.ListIndex = 0
            Combo1(0).ListIndex = 1
            Combo1(0).SetFocus
          '  txtBox(11).Enabled = False
           ' cmdEkstraksi.Enabled = True
            Combo1(1).Enabled = True
            
            mAdd = True
            'SSTab1.Tab = 0
            

       Case tmbEdit:
            Command1.Enabled = True
            txtBox(0).Enabled = False
            mAdd = True
            txtBox(0).Locked = True
            'cmdEkstraksi.Enabled = True
            txtBox(1).SetFocus
            
            'SSTab1.Tab = 0
            DTPicker1(3).Enabled = False
            DTPicker1(4).Enabled = False
            Combo1(1).Enabled = True
            
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
              ' mAdd = True
            Else
               mAdd = False
            End If
            mAddOnly = Combo1(1).Enabled
            'cmdEkstraksi.Enabled = False
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then
               'SSTab1.Tab = 1
'               OpenDetailPartner 1
            End If
       Case tmbSave:
'            If ((Combo1(1).Text <> "RELEASED" Or Combo1(1).Text <> "FINISHED") And Combo1(0).Text <> "MAKE STOCK") Then
            If MyDDE.IsChildMemberReady = True Then
                If MyDDE.ChildRecordset.Recordcount <> 0 Then
'                MsgBox MyDDE.ChildRecordset.Source
                   With MyDDE.ChildRecordset
                        .MoveFirst
                        z = 0
                        If SendDataToServer("Delete From [Order Output Detail] WHERE     (OrderID = N'" & Label2 & "')") = True Then
                             Do
                             If MyDDE.ChildRecordset.EOF Then Exit Do
                                z = z + 1
                                SendDataToServer " INSERT INTO [Order Output Detail]" & _
                                                 " (Warehouse,EndDate,StartDate,OrderID, SeqNo, wcid,  ResourcesID,Status)" & _
                                                 " VALUES  (N'" & MyDDE.ChildRecordset.Fields("Warehouse") & "','" & Format(MyDDE.ChildRecordset.Fields("EndDate"), "yyyy-MM-dd hh:mm:ss") & "','" & Format(MyDDE.ChildRecordset.Fields("StartDate"), "yyyy-MM-dd hh:mm:ss") & "',N'" & Label2 & "', " & MyDDE.ChildRecordset.Fields("SeqNo") & ", N'" & MyDDE.ChildRecordset.Fields("sTAGEId") & "', N'" & MyDDE.ChildRecordset.Fields("ResourcesID") & "'," & BoolToInt(MyDDE.ChildRecordset.Fields("Status")) & ")"
                                .MoveNext
                             Loop
                        End If
                        .MoveLast
                   End With
                   With RcCompDetail.DBRecordset
                       If .Recordcount <> 0 Then
                        .MoveFirst
                        If SendDataToServer("DELETE FROM [Ord Comp Detail] WHERE (OrderID = N'" & Label2 & "')") = True Then
                           Do
                             If .EOF Then Exit Do
                             SendDataToServer (" INSERT INTO [Ord Comp Detail]" & _
                                               " (SeqNo, StageID, SeqStageID, OrderID, NoItem, [DESC], UOM, [Quote Qty], [Actual Qty], Phantom, Complete, PartnerID, PurchaseID, wip_qty, scrap_qty)" & _
                                               " VALUES     (" & .Fields("SeqNo") & ", N'" & .Fields("Work Center") & "', N'" & .Fields("Routing") & "', N'" & Label2 & "', N'" & .Fields("Kode Barang") & "', " & _
                                               " N'" & .Fields("Nama Barang") & "', N'" & .Fields("UOM") & "', " & FQty(.Fields("Quote Qty")) & ", " & FQty(.Fields("Actual Qty")) & "," & _
                                               BoolToInt(.Fields("Phantom")) & ", " & BoolToInt(.Fields("Complete")) & ", N'" & .Fields("Partner ID") & "', N'" & .Fields("No PO") & "', " & FQty(.Fields("wip_qty")) & ", " & FQty(.Fields("scrap_qty")) & ")")
                             .MoveNext
                           Loop
                           .MoveLast
                        End If
                       End If
                   End With
                End If
                
                
                '*** Update Nomor Ekstraksi yang baru
                ' SendDataToServer "UPDATE labrekomekstraksi SET  status = 1 Where splno = '" & txtBox(5).Text & "'"
                mAdd = False
                mAddOnly = False
            End If
'            End If
            
       Case tmbPrint:
            CallRPTReport "Manufacture Order Table.rpt", "select * from [Manufacture Order Table] where OrderID =N'" & Label2 & "'"
       Case Else: 'mVarDataDc = False
End Select


If (Combo1(1).Text <> "RELEASED" Or Combo1(1).Text <> "FINISHED") Or (Combo1(0).Text <> "MAKE STOCK") Then
'    MoveCtrl
    cmdLink.Enabled = mAdd
    'cmdLink1.Enabled = mAdd
'    TglStart.Enabled = mAdd
'    TglStart.Visible = mAdd
    DataGrid1(1).Columns(6).Button = mAdd
    DataGrid1(1).Columns(7).Button = mAdd
    Picture5.Enabled = True
Else
   cmdLink.Enabled = mAdd
  ' cmdLink1.Enabled = mAdd
   DataCombo1.Enabled = False
   txtBox(3).Enabled = False
   txtBox(1).Enabled = False
   'DataCombo2.Enabled = False
   txtBox(4).Enabled = False
   Picture5.Enabled = False
End If
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   PrevStatus = MyDDE.GetFieldByName("StatusOrder")
   
   If IsNull(MyDDE.GetFieldByName("released_date")) Then
      Label1(23).Visible = False
      Line1(23).Visible = False
      DTPicker1(6).Visible = False
   Else
      Label1(23).Visible = True
      Line1(23).Visible = True
      DTPicker1(6).Visible = True
   End If
   
   If IsNull(MyDDE.GetFieldByName("closed_date")) Then
      Line1(24).Visible = False
      Label1(22).Visible = False
      DTPicker1(7).Visible = False
   Else
      Line1(24).Visible = True
      Label1(22).Visible = True
      DTPicker1(7).Visible = True
   End If
   
   OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Order")), MyDDE.GetFieldByName("No Order"), "xxx")
   
   OpenComponentDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Order")), MyDDE.GetFieldByName("No Order"), "xxx")
   
   If Not IsNull(MyDDE.GetFieldByName("Kode Bom")) Then
      OpenMOWhereUsed MyDDE.GetFieldByName("Kode Bom"), MyDDE.GetFieldByName("No Order")
   End If
   
   LoadTreeBOM
   
   RsRekomendasi.DBOpen "Select formid,formname from  view_rekomekstraksi_proses where splno ='" & IIf(IsNull(MyDDE.GetFieldByName("no_rekomendasi")), "x", MyDDE.GetFieldByName("no_rekomendasi")) & "'", CNN
   Set DataGrid1(3).DataSource = RsRekomendasi.DBRecordset
   
   RsRL.DBOpen "select Idtrans, NoItem, batch_lot, UOM, [Qty Required], [Qty Received] from backflush_line where orderID ='" & Label2.Caption & "'", CNN, lckLockBatch
   Set DataGrid1(4).DataSource = RsRL.DBRecordset

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
DTPicker1(6).Enabled = False
DTPicker1(7).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               If Combo1(1).Text = "CLOSED" Or Combo1(1).Text = "FINISHED" Then
                  MyDDE.IsChildMemberReady = False
                  MyDDE.CancelTrans = True
                  MessageBox "Data sudah ditutup dan tidak bisa diedit lagi.", "Peringatan", msgOkOnly
               End If
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
'               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                If Combo1(1) <> "ORDERED" Or Combo1(1) <> "NEW" Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Transaksi berstatus ORDER dan NEW yang boleh dihapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  TanggalStart
                  With RcCompDetail.DBRecordset
                       If .Recordcount <> 0 Then
                       .MoveFirst
                       Do
                            If .EOF Then Exit Do
                            .Fields("Quote Qty") = CekStock(.Fields("Kode Barang")) * CDbl(txtBox(3))
                            .MoveNext
                       Loop
                       .MoveFirst
                       End If
                  End With
'                  LihatTanggal Label2, True
                  PrepareQuery
               Else
                  MessageBox "Data stage detail belum ada.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
                  'PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDetail:
            MyDDE.CancelTrans = mFirstCaller
            If MyDDE.CancelTrans = True Then Exit Sub
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
End Select
Set mDel = Nothing
End Sub

Private Sub SemeruTree1_NodeClick(ByVal Node As MSComctlLib.INode)
   If RcCompDetail.DBRecordset.Recordcount > 0 Then
      RcCompDetail.DBRecordset.MoveFirst
      RcCompDetail.DBRecordset.Find "[Kode Barang]='" & Node.Key & "'", , adSearchForward, 1
      If Not RcCompDetail.DBRecordset.EOF Then
            txt(0).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("Kode Barang")), "", RcCompDetail.DBRecordset.Fields("Kode Barang"))
            txt(1).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("Nama Barang")), "", RcCompDetail.DBRecordset.Fields("Nama Barang"))
            txt(2).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("kategori")), "", RcCompDetail.DBRecordset.Fields("kategori"))
            txt(3).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("kelompok")), "", RcCompDetail.DBRecordset.Fields("kelompok"))
            txt(4).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("Quote Qty")), "", RcCompDetail.DBRecordset.Fields("Quote Qty"))
            txt(5).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("LeadTimeDays")), "", RcCompDetail.DBRecordset.Fields("LeadTimeDays"))
            lblUOM.Caption = IIf(IsNull(RcCompDetail.DBRecordset.Fields("UOM")), "", RcCompDetail.DBRecordset.Fields("UOM"))
      End If
   End If
End Sub

'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''MoveForm Picture1.Parent.hwnd
'End Sub

Private Sub TglStart_Change()
'If DataGrid1(1).Col = 3 Or DataGrid1(1).Col = 4 Then DataGrid1(1).Columns(DataGrid1(1).Col).Value = TglStart.Value
Call TglStart_Click
End Sub

Private Sub OpenMOWhereUsed(ByVal Param1 As String, ByVal Param2 As String)
    Set RcMOWhere = New DBQuick

    strSQL = "SELECT [Manufacture Order].OrderID AS No_MO, Inventory.ItemName, [Order Output Detail].Status, " & _
        " [Manufacture Order].OrderName, [Manufacture Order].Type, [Manufacture Order].NoItem, [Manufacture Order].Status AS StatusMO, " & _
        " MIN([Order Output Detail].StartDate) AS StartDate, MAX([Order Output Detail].EndDate) AS EndDate, PartnerDB.CompanyName " & _
        " FROM [Order Output Detail] INNER JOIN [Manufacture Order] ON [Order Output Detail].OrderID = [Manufacture Order].OrderID " & _
        " INNER JOIN Inventory ON [Manufacture Order].NoItem = Inventory.NoItem LEFT OUTER JOIN " & _
        " PartnerDB ON [Manufacture Order].PartnerID = PartnerDB.PartnerID " & _
        " GROUP BY [Manufacture Order].OrderID, Inventory.ItemName, [Manufacture Order].OrderName, [Manufacture Order].Type, " & _
        " [Manufacture Order].NoItem , [Order Output Detail].status, [Manufacture Order].status, PartnerDB.CompanyName " & _
        " HAVING ([Manufacture Order].NoItem = N'" & Param1 & "') AND (dbo.[Manufacture Order].OrderID <> N'" & Param2 & "')"
        
'    Debug.Print RcMOWhere.DBRecordset.Source
    RcMOWhere.DBOpen strSQL, CNN, lckLockReadOnly
    Set GrdMOUsed.DataSource = RcMOWhere.DBRecordset
End Sub

Private Sub TglStart_Click()
'If DataGrid1(1).col <= 0 Then MoveCtrl
DataGrid1(1).Columns(DataGrid1(1).col).Value = TglStart.Value
Select Case DataGrid1(1).col
       Case 3:
            If CDate(DataGrid1(1).Columns(3).Value) < CDate(DTPicker1(0).Value) Then
               DataGrid1(1).Columns(3).Value = DTPicker1(0).Value
               TglStart.Value = DataGrid1(1).Columns(3).Value
               'If CDate(DataGrid1(1).Columns(3).Value) > CDate(DTPicker1(3).Value) Then
'               DTPicker1(3).Value = CDate(DataGrid1(1).Columns(3).Value)
               'End If
            End If
       Case 4:
            If CDate(DataGrid1(1).Columns(4).Value) < CDate(DataGrid1(1).Columns(3).Value) Then
               DataGrid1(1).Columns(4).Value = DataGrid1(1).Columns(3).Value
               TglStart.Value = DataGrid1(1).Columns(4).Value
'               If CDate(DataGrid1(1).Columns(4).Value) > CDate(DTPicker1(4).Value) Then
'                  DTPicker1(4).Value = CDate(DataGrid1(1).Columns(4).Value)
'               End If
            End If
End Select
MyDDE.ChildRecordset.Fields("StartDate") = CDate(DataGrid1(1).Columns(3).Value)
MyDDE.ChildRecordset.Fields("EndDate") = CDate(DataGrid1(1).Columns(4).Value)
End Sub

Private Sub TglStart_DropDown()
Call TglStart_Click
End Sub

Private Sub TglStart_Validate(Cancel As Boolean)
'Call TglStart_Click
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
'if MyDDE.GetFieldByName("")
With MyDDE
    .PrepareAppend = " INSERT INTO [Manufacture Order]" & _
                     " (Noitem,DateIssued,EmpID,OrderID, [QTY Order],PartnerID, OrderName, Type, Status, Note, ContractID,[CreateDate], [Priority], [RequireDate], [EarliesDate], [StartDate], [FinishedDate],no_rekomendasi)" & _
                     " VALUES (N'" & MyDDE.GetFieldByName("Kode Bom") & "','" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "',N'" & txtBox(4).Text & "', N'" & Label2 & "'," & FQty(txtBox(3)) & ", N'" & DataCombo1.BoundText & "', N'" & txtBox(0) & "', N'" & Combo1(0) & "', N'" & Combo1(1).Text & "', N'" & txtBox(1) & "', N'" & txtBox(2) & "','" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & "',N'" & Combo2.Text & "','" & Format(DTPicker1(1).Value, "yyyy-MM-dd") & "','" & Format(DTPicker1(2).Value, "yyyy-MM-dd") & "','" & Format(DTPicker1(3).Value, "yyyy-MM-dd") & "','" & Format(DTPicker1(4).Value, "yyyy-MM-dd") & "','" & MyDDE.GetFieldByName("no_rekomendasi") & "')"
                     
'    .PrepareUpdate = " UPDATE [Manufacture Order] Set Noitem=N'" & MyDDE.GetFieldByName("Kode Bom") & _
'                                                "',empID = N'" & txtBox(4).Text & "', [QTY Order]= " & CDbl(txtBox(3)) & _
'                                                ", [OrderName] = N'" & txtBox(0) & _
'                                                "',PartnerID=N'" & DataCombo1.BoundText & _
'                                                "',Type=N'" & Combo1(0) & _
'                                                "',Status=N'" & Combo1(1).Text & _
'                                                "',Note=N'" & txtBox(1) & _
'                                                "',ContractID=N'" & txtBox(2) & _
'                                                "',[CreateDate]='" & Format(DTPicker1(0).value, "yyyy-MM-dd") & _
'                                                "',[Priority]='" & Combo2.Text & _
'                                                "',[RequireDate]='" _
'                                                & Format(DTPicker1(1).value, "yyyy-MM-dd") & _
'                                                "',[EarliesDate]='" & Format(DTPicker1(2).value, "yyyy-MM-dd") & _
'                                                "',[StartDate]='" & Format(DTPicker1(3).value, "yyyy-MM-dd") & _
'                                                "',[FinishedDate]='" & Format(DTPicker1(4).value, "yyyy-MM-dd") & _
'                                                "',ekstraksi_no='" & txtBox(11).Text & _
'                                                "',released_date=" & dReleased & _
'                                                " ,closed_date=" & dClosed & _
'                     " WHERE ([OrderID] = N'" & Label2 & "')"

    .PrepareUpdate = " UPDATE [Manufacture Order] Set Noitem=N'" & MyDDE.GetFieldByName("Kode Bom") & _
                                                "',empID = N'" & txtBox(4).Text & "', [QTY Order]= " & CDbl(txtBox(3)) & _
                                                ", [OrderName] = N'" & txtBox(0) & _
                                                "',PartnerID=N'" & DataCombo1.BoundText & _
                                                "',Type=N'" & Combo1(0) & _
                                                "',Status=N'" & Combo1(1).Text & _
                                                "',Note=N'" & txtBox(1) & _
                                                "',ContractID=N'" & txtBox(2) & _
                                                "',[CreateDate]='" & Format(DTPicker1(0).Value, "yyyy-MM-dd") & _
                                                "',[Priority]='" & Combo2.Text & _
                                                "',[RequireDate]='" _
                                                & Format(DTPicker1(1).Value, "yyyy-MM-dd") & _
                                                "',[EarliesDate]='" & Format(DTPicker1(2).Value, "yyyy-MM-dd") & _
                                                "',[StartDate]='" & Format(DTPicker1(3).Value, "yyyy-MM-dd") & _
                                                "',[FinishedDate]='" & Format(DTPicker1(4).Value, "yyyy-MM-dd") & _
                                                "',no_rekomendasi='" & MyDDE.GetFieldByName("no_rekomendasi") & _
                     "' WHERE ([OrderID] = N'" & Label2 & "')"
    
    .PrepareDelete = " DELETE FROM [Manufacture Order] WHERE   ([OrderID] = N'" & Label2 & "') "
End With
End Sub

Private Sub GridLayout()
'DataGrid1(0).Height = 2225
'DataGrid1(0).Width = 7590
DataGrid1(1).Columns(0).width = 464.8819
DataGrid1(1).Columns(1).width = 1514.835
DataGrid1(1).Columns(2).width = 2500
DataGrid1(1).Columns(3).width = 1850
DataGrid1(1).Columns(4).width = 1850
DataGrid1(1).Columns(5).width = 1200
DataGrid1(1).Columns(6).width = 1400
DataGrid1(1).Columns(7).width = 1514.835
DataGrid1(2).Columns(0).width = 659.9055
DataGrid1(2).Columns(1).width = 1335.118
DataGrid1(2).Columns(2).width = 1964.976
DataGrid1(2).Columns(3).width = 2594.835
DataGrid1(2).Columns(4).width = 780.0945
DataGrid1(2).Columns(5).width = 780.0945
DataGrid1(2).Columns(6).width = 780.0945
DataGrid1(2).Columns(7).width = 1600.929
DataGrid1(2).Columns(8).width = 1514.835
DataGrid1(2).Columns(9).width = 975.1182
DataGrid1(2).Columns(10).width = 975.1182

With DataGrid1(2)
    .Columns(11).Visible = False
    .Columns(12).Visible = False
    .Columns(14).Visible = False
    'WIDTH
    .Columns(0).width = 700     'SeqNo
    .Columns(1).width = 1500    'StageID
    .Columns(2).width = 1500    'SeqStageID
    .Columns(3).width = 1500    'Kode Barang
    .Columns(4).width = 2500    'Nama Barang
    .Columns(5).width = 500    'UOM
    .Columns(6).width = 1000    'Quote Qty
    .Columns(7).width = .Columns(6).width
    .Columns(8).width = .Columns(6).width
    .Columns(9).width = .Columns(6).width
    .Columns(10).width = .Columns(6).width
    'ALIGNMENT
    .Columns(6).Alignment = dbgRight
    .Columns(7).Alignment = dbgRight
    .Columns(8).Alignment = dbgRight
    .Columns(9).Alignment = dbgRight
    .Columns(10).Alignment = dbgRight
    'NUMBER FORMAT
    .Columns(6).NumberFormat = QtyFormFloat
    .Columns(7).NumberFormat = QtyFormFloat
    .Columns(8).NumberFormat = QtyFormFloat
    .Columns(9).NumberFormat = QtyFormFloat
    .Columns(10).NumberFormat = QtyFormFloat
    Set .Columns(15).DataFormat = fmtBoolOpenClose
    'CAPTION
    .Columns(6).Caption = "Quote"
    .Columns(7).Caption = "Actual"
    .Columns(8).Caption = "Scrapped"
    .Columns(9).Caption = "WIP"
    .Columns(10).Caption = "Completed"
    .Columns(15).Caption = "Status"
    .Columns(16).Caption = "Keterangan"
End With
With DataGrid1(1)
    .Columns(1).Caption = "Work Center"
End With
End Sub

Private Sub OpenDetail(ByVal Param As String)
Dim mVarFrmDate As New StdDataFormat
Dim Rc As New DBQuick
Dim rsCek As New DBQuick
Dim stDate As Date
Dim nQtyOrder As Double
Dim nCountOfHour As Integer

nQtyOrder = IIf(IsNull(MyDDE.GetFieldByName("QTY Order")), 0, IIf(IsEmpty(MyDDE.GetFieldByName("QTY Order")), 0, MyDDE.GetFieldByName("QTY Order")))

Rc.DBOpen "SELECT [Order Output Detail].SeqNo, " & _
                 "[Order Output Detail].WCID AS StageID, " & _
                 "[Order Output Detail].OrderID," & _
                 "wcenter_header.Description AS Keterangan," & _
                 "[Order Output Detail].StartDate, " & _
                 "[Order Output Detail].EndDate, " & _
                 "[Order Output Detail].ResourcesID, " & _
                 "[Order Output Detail].Status, " & _
                 "[Order Output Detail].WareHouse, " & _
                 "[Order Output Detail].actual_time, " & _
                 "[Order Output Detail].overlap, " & _
                 "wcenter_header.cycle_time / 60 as unit_run, " & _
                 "wcenter_header.queue_time, " & _
                 "wcenter_header.setup_time, " & _
                 "wcenter_header.wait_time, " & _
                 "(wcenter_header.cycle_time / 60 / 60 *" & nQtyOrder & ") + ((wcenter_header.queue_time + wcenter_header.setup_time + wcenter_header.wait_time)/3600) as total_run_time, " & _
                 "wcenter_header.cycle_time / 60 / 60 *" & nQtyOrder & " as extended_run " & _
         "FROM [Order Output Detail] INNER JOIN wcenter_header ON [Order Output Detail].WCID = wcenter_header.WCID " & _
         "WHERE ([Order Output Detail].OrderID = N'" & Param & "') ORDER BY [Order Output Detail].SeqNo", CNN, lckLockBatch
'Debug.Print Rc.DBRecordset.Source
Set MyDDE.ChildRecordset = Rc.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(1).DataSource = MyDDE.ChildRecordset
Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
mVarFrmDate.Format = "dd-mmm-yy HH:MM:SS AM/PM"
Set DataGrid1(1).Columns(3).DataFormat = mVarFrmDate
Set DataGrid1(1).Columns(4).DataFormat = mVarFrmDate

With Rc.DBRecordset
   If .Recordcount > 0 Then
      .MoveFirst
      While Not .EOF
         stDate = GetEndDate(.Fields("StartDate"), .Fields("total_run_time"))
         .Fields("EndDate") = stDate
         SendDataToServer "update [order Output Detail] set endDate='" & Format(.Fields("EndDate"), "yyyy-MM-dd hh:mm:ss") & "' where seqNo=" & .Fields("SeqNo") & " and WCID='" & .Fields("StageID") & "' and OrderID='" & .Fields("OrderID") & "'"
         If Val(.Fields("Total_run_time")) >= Val(.Fields("actual_time")) Then
            .Fields("overlap") = False
         Else
            '*** cek holiday***'
            rsCek.DBOpen "select datefrom,dateto from [scheduling calendar detail] where datefrom between '" & Format(Rc.DBRecordset.Fields("startDate"), "yyyy-MM-dd") & "' and '" & Format(Rc.DBRecordset.Fields("EndDate"), "yyyy-MM-dd") & "'", CNN
            If rsCek.DBRecordset.Recordcount > 0 Then
               nCountOfHour = 0
               While Not rsCek.DBRecordset.EOF
                  If rsCek.DBRecordset.Fields(0) = rsCek.DBRecordset.Fields(1) Then
                     nCountOfHour = nCountOfHour + 24
                  Else
                     nCountOfHour = nCountOfHour + Val(SelisihHariJam(rsCek.DBRecordset.Fields(0), rsCek.DBRecordset.Fields(1), 2))
                  End If
                  rsCek.DBRecordset.MoveNext
               Wend
            End If
            If (Val(.Fields("Total_run_time")) + nCountOfHour) >= Val(.Fields("actual_time")) Then
               .Fields("overlap") = False
            Else
               .Fields("overlap") = True
            End If
         End If
         .MoveNext
         If Not .EOF Then
            .Fields("StartDate") = stDate
            SendDataToServer "update [order Output Detail] set startDate='" & Format(.Fields("startDate"), "yyyy-MM-dd hh:mm:ss") & "' where seqNo=" & .Fields("SeqNo") & " and WCID='" & .Fields("StageID") & "' and OrderID='" & .Fields("OrderID") & "'"
         End If
      Wend
   End If
End With

End Sub

Private Sub OpenPartner()
RcPart.DBOpen "SELECT PartnerID AS Pelanggan, CompanyName AS [Nama Perusahaan] FROM         PartnerDB WHERE     (PartnerType = N'CUSTOMER')", CNN, lckLockBatch
DataCombo1.ListField = "Nama Perusahaan"
Set DataCombo1.RowSource = RcPart.DBRecordset
End Sub

Private Function IndexAuto() As String
Dim Rc As New DBQuick
Dim TglSaiki As String
Dim Inom As Long
TglSaiki = Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2)
Rc.DBOpen "SELECT MAX(RIGHT(OrderID, 5)) AS MaxNom FROM [Manufacture Order]         [Manufacture Order] WHERE     (GETDATE() = { fn NOW() })", CNN, lckLockReadOnly
With Rc
     If .DBRecordset.Recordcount <> 0 Then
        Inom = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + 1
     Else
        Inom = 1
     End If
     Select Case Len(Trim(Str(Inom)))
            Case 0: IndexAuto = "OD/" & TglSaiki & "-" & Trim(Str(Inom))
            Case 1: IndexAuto = "OD/" & TglSaiki & "-" & "0000" & Trim(Str(Inom))
            Case 2: IndexAuto = "OD/" & TglSaiki & "-" & "000" & Trim(Str(Inom))
            Case 3: IndexAuto = "OD/" & TglSaiki & "-" & "00" & Trim(Str(Inom))
            Case 4: IndexAuto = "OD/" & TglSaiki & "-" & "0" & Trim(Str(Inom))
     End Select
End With
End Function

Private Sub OpenDetailPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1: RcPartner.DBOpen "SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], Merk, UOM FROM Inventory WHERE (Manufacture = 1) ORDER BY NoItem", CNN, lckLockReadOnly
       Case 2: RcPartner.DBOpen "SELECT StageID AS [Stage ID], Description AS Keterangan FROM [Manufacture Stage]", CNN, lckLockReadOnly
       Case 3: RcPartner.DBOpen "SELECT NoItem as [Kode Barang], ItemName as [Nama Barang] FROM Inventory WHERE (Manufacture = 1) and (status =N'ACTIVE') ORDER BY NoItem", CNN, lckLockReadOnly
       Case 4: RcPartner.DBOpen "SELECT WareHouse AS ID, [WareHouse Name] AS [Nama Gudang] FROM WareHouse ORDER BY WareHouse", CNN, lckLockReadOnly
       Case 5: RcPartner.DBOpen "SELECT LabRekomEkstraksi.SplNo as [Nomor Ekstraksi], keterangan as [Keterangan] From LabRekomEkstraksi where status=0", CNN, lckLockReadOnly
       Case 6: RcPartner.DBOpen "Select labrekomekstraksi.splno as [No Rekomendasi] from labrekomekstraksi where labrekomekstraksi.status=0 and labrekomekstraksi.approved_by is not null ", CNN, lckLockReadOnly
       Case 7: RcPartner.DBOpen "SELECT no_ekstraksi AS [Nomor Ekstraksi]  From labrekomekstraksi_no where status=0 and splno ='" & txtBox(5) & "'", CNN, lckLockReadOnly
       Case 8: RcPartner.DBOpen "SELECT [inventory tabel].sl_no, [inventory tabel].stockTmp, inventory.uom  from [inventory tabel] inner join inventory on inventory.NoItem = [inventory tabel].noItem where inventory.noItem='BB-BA-0001' and lockFIFO=0", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index:
          Case 1: mCall.FromTagActive = "Master Barang"
          Case 2: mCall.FromTagActive = "Master Stage"
          Case 3: mCall.FromTagActive = "BOM Listing"
          Case 4: mCall.FromTagActive = "Gudang"
          Case 5: mCall.FromTagActive = "No Ekstraksi"
          Case 6: mCall.FromTagActive = "No Rekomendasi"
          Case 7: mCall.FromTagActive = "Nomor Ekstraksi"
          Case 8: mCall.FromTagActive = "Daftar RL Batch"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
'Debug.Print RcPartner.DBRecordset.Source
   MessageBox "Data transaksi Belum Ada.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
'    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub MoveCtrl()
On Error GoTo Hell
If (Combo1(1).Text <> "RELEASED") Or (Combo1(0).Text <> "MAKE STOCK") Then
If MyDDE.ChildRecordset.Recordcount <> 0 And mAdd = True Then
'    TglStart.Enabled = mAdd
'    TglStart.Visible = mAdd
'    TglStart.SetFocus
   With DataGrid1(1)
'        TglStart.Height
        If .col < 1 Then .col = 3
        If .col = 3 Or .col = 4 Then
'           TglStart.Enabled = True
'           TglStart.Visible = True
'           TglStart.SetFocus
'           DataGrid1(1).Enabled = False
           TglStart.Move .Columns(.col).Left, (.RowTop(.row) + .RowHeight) - 250, .Columns(.col).width, .RowHeight + 60
           .AllowUpdate = False
'           TglStart.Enabled = False
'           TglStart.Visible = False
'           DataGrid1(1).Enabled = True
'           DataGrid1(1).SetFocus
'           TglStart.Enabled = True
'           TglStart.Visible = True
           TglStart.ZOrder (0)
           TglStart.Value = IIf(Not IsNull(DataGrid1(1).Columns(DataGrid1(1).col).Value), DataGrid1(1).Columns(DataGrid1(1).col).Value, Date)
           If TglStart.Visible = True And TglStart.Enabled = True Then TglStart.SetFocus
        Else
           TglStart.Enabled = False
           TglStart.Visible = False
        End If
   End With
End If
End If
Hell:
If Err.Number <> 0 Then
   TglStart.Enabled = False
   TglStart.Visible = False
End If
Err.Clear
End Sub

Private Sub OpenComponentDetail(ByVal Param As String)

'strSQL = "SELECT [Ord Comp Detail].SeqNo, [Ord Comp Detail].StageID AS [Work Center], " & _
'        " [Ord Comp Detail].SeqStageID AS Routing, [Ord Comp Detail].NoItem AS [Kode Barang], [Ord Comp Detail].[DESC] AS [Nama Barang], " & _
'        " [Ord Comp Detail].UOM, [Ord Comp Detail].[Quote Qty], [Ord Comp Detail].[Actual Qty], [Ord Comp Detail].wip_qty, " & _
'        " [Ord Comp Detail].scrap_qty, [Ord Comp Detail].PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], " & _
'        " [Ord Comp Detail].PurchaseID AS [No PO], [Ord Comp Detail].Phantom, [Ord Comp Detail].Complete " & _
'        ",[Ord Comp Detail].comment " & _
'        " FROM [Ord Comp Detail] INNER JOIN Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem " & _
'        " INNER JOIN PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID " & _
'        " WHERE     ([Ord Comp Detail].OrderID = N'" & Param & "') ORDER BY [Ord Comp Detail].SeqNo"
        
        
strSQL = "SELECT  [Ord Comp Detail].SeqNo, [Ord Comp Detail].StageID AS [Work Center], [Ord Comp Detail].SeqStageID AS Routing, " & _
                 "[Ord Comp Detail].NoItem AS [Kode Barang], Inventory.InternalName AS [Nama Barang], [Ord Comp Detail].UOM, " & _
                 "[Ord Comp Detail].[Quote Qty], [Ord Comp Detail].[Actual Qty], [Ord Comp Detail].wip_qty, [Ord Comp Detail].scrap_qty, " & _
                 "[Ord Comp Detail].PartnerID AS [Partner ID], PartnerDB.CompanyName AS [Nama Perusahaan], [Ord Comp Detail].PurchaseID AS [No PO], " & _
                 "[Ord Comp Detail].Phantom, [Ord Comp Detail].Complete, [Ord Comp Detail].comment, inventory_categories.description AS kategori, " & _
                 "[Inventory Group].[Group Name] AS kelompok, Inventory.LeadTimeDays " & _
         "FROM  [Ord Comp Detail] INNER JOIN " & _
                  "Inventory ON [Ord Comp Detail].NoItem = Inventory.NoItem INNER JOIN " & _
                  "PartnerDB ON Inventory.PartnerID = PartnerDB.PartnerID INNER JOIN " & _
                  "inventory_categories ON Inventory.categid = inventory_categories.categid AND " & _
                  "Inventory.NoGroup = inventory_categories.nogroup INNER JOIN " & _
                  "[Inventory Group] ON inventory_categories.nogroup = [Inventory Group].NoGroup " & _
         " WHERE  ([Ord Comp Detail].OrderID = N'" & Param & "') " & _
         "ORDER BY [Ord Comp Detail].SeqNo "
        
RcCompDetail.DBOpen strSQL, CNN, lckLockBatch
Set DataGrid1(2).DataSource = RcCompDetail.DBRecordset

End Sub

Private Sub LihatTanggal(ByVal Param As String, Optional ByVal TipicalShowTanggal As Boolean)
Dim rcTgl As New DBQuick
rcTgl.DBOpen "SELECT     MIN(StartDate) AS StartDate, MAX(EndDate) AS EndDate FROM         [Order Output Detail] WHERE     (OrderID = N'" & Param & "') GROUP BY SeqNo ORDER BY SeqNo", CNN, lckLockReadOnly
With rcTgl.DBRecordset
     If .Recordcount <> 0 Then
        If TipicalShowTanggal = False Then
        Else
           .MoveFirst
           DTPicker1(3).Value = IIf(Not IsNull(.Fields(0)), .Fields(0), Date)
           MyDDE.GetFieldByName("StartDate") = DTPicker1(3).Value
           .MoveLast
           DTPicker1(4).Value = IIf(Not IsNull(.Fields(1)), .Fields(1), Date)
           MyDDE.GetFieldByName("FinishedDate") = DTPicker1(4).Value
        End If
     End If
     .Close
End With
Set rcTgl = Nothing
End Sub

Private Sub TanggalStart()
Dim RcT As New DBQuick
Dim Avdata As Variant
Dim tglMulai, TglSelesai As Date
Dim I As Integer
Set RcT.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcT.DBRecordset
     If .Recordcount <> 0 Then
         tglMulai = IIf(Not IsNull(.Fields("StartDate")), .Fields("StartDate"), Date)
         TglSelesai = IIf(Not IsNull(.Fields("EndDate")), .Fields("EndDate"), Date)
         Avdata = .Getrows(.Recordcount, adBookmarkFirst)
         For I = 0 To UBound(Avdata, 2)
            ' 3 mulai 4 end
            If tglMulai <= CDate(Avdata(4, I)) Then DTPicker1(3).Value = tglMulai
            If TglSelesai <= CDate(Avdata(5, I)) Then DTPicker1(4).Value = CDate(Avdata(4, I))
         Next I
     End If
End With
Set Avdata = Nothing
RcT.CloseDB
Set RcT = Nothing
End Sub

Private Function LoadQty(ByVal NoItem As String) As Long
Dim Rcl As New DBQuick
Rcl.DBOpen "SELECT      SUM([Qty Received]) AS [Qty Received] FROM         backflush_line WHERE     (OrderID = N'" & Label2 & "') AND (NoItem = N'" & NoItem & "')  GROUP BY [Qty Received]", CNN, lckLockReadOnly
With Rcl.DBRecordset
     If .Recordcount <> 0 Then
        LoadQty = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        LoadQty = 0
     End If
End With
Rcl.CloseDB
Set Rcl = Nothing
End Function

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then ValidNum KeyAscii
End Sub

Private Function CekStock(ByVal NoItem As String) As Long
Dim Rcdata As New DBQuick
Rcdata.DBOpen "SELECT     QTYUsage FROM         [BOM Component Detail] WHERE     (COMPONENT = N'" & NoItem & "') GROUP BY QTYUsage", CNN, lckLockReadOnly
With Rcdata.DBRecordset
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 1
     End If
     Rcdata.CloseDB
End With
Set Rcdata = Nothing
End Function


Private Function GetEndDate(startDate As Date, addHours As Double) As Date
   Dim aSec As Integer
   Dim aMin As Integer
   Dim aHou As Integer
   Dim sTime, eTime As Date
   aSec = (addHours * 3600) Mod 60
   aMin = ((addHours * 3600) / 60) Mod 60
   aHou = addHours
'   Debug.Print Val(Second(startDate)) + aSec
   If Val(Second(startDate)) + aSec > 60 Then
      aMin = (Val(Second(startDate)) + aSec) / 60
      aSec = (aSec + Val(Second(startDate))) Mod 60
   Else
      aSec = aSec + Val(Second(startDate))
   End If
'   Debug.Print (Format(startDate, "m"))
'   Debug.Print Minute(startDate)
'   Debug.Print Hour(startDate)
'   Debug.Print Second(startDate)
   If Val(Minute(startDate)) + aMin > 60 Then
      aHou = (Val(Minute(startDate)) + aMin) / 60
      aMin = (aMin + Val(Minute(startDate))) Mod 60
   Else
      aMin = aMin + Val(Minute(startDate))
   End If
   
   If Val(Hour(startDate)) + aHou > 24 Then
      startDate = (Val(Hour(startDate)) + aHou) / 24
      aHou = (aHou + Val(Hour(startDate))) Mod 24
   Else
      aHou = (aHou + Val(Hour(startDate)))
   End If
 '  MsgBox Format(startDate, "dd/MM/yy") & " " & IIf(Len(aHou) = 1, "0" & aHou, aHou) & ":" & aMin & ":" & aSec
   If aSec = 60 Then
      aMin = aMin + 1
      aSec = 0
   End If
   
   If aMin = 60 Then
      aHou = aHou + 1
      aMin = 0
   End If
   GetEndDate = CDate(Format(startDate, "dd/MM/yy") & " " & IIf(Len(aHou) = 1, "0" & aHou, aHou) & ":" & aMin & ":" & aSec)
End Function


Private Sub LoadTreeBOM()
   Dim sWCID As String
   On Error GoTo xErr
   SemeruTree1.MenuTreeView.Nodes.Clear
   If RcCompDetail.DBRecordset.Recordcount > 0 Then
      With SemeruTree1
      
         txt(0).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("Kode Barang")), "", RcCompDetail.DBRecordset.Fields("Kode Barang"))
         txt(1).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("Nama Barang")), "", RcCompDetail.DBRecordset.Fields("Nama Barang"))
         txt(2).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("kategori")), "", RcCompDetail.DBRecordset.Fields("kategori"))
         txt(3).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("kelompok")), "", RcCompDetail.DBRecordset.Fields("kelompok"))
         txt(4).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("Quote Qty")), "", RcCompDetail.DBRecordset.Fields("Quote Qty"))
         txt(5).Text = IIf(IsNull(RcCompDetail.DBRecordset.Fields("LeadTimeDays")), "", RcCompDetail.DBRecordset.Fields("LeadTimeDays"))
         lblUOM.Caption = IIf(IsNull(RcCompDetail.DBRecordset.Fields("UOM")), "", RcCompDetail.DBRecordset.Fields("UOM"))

         
         Set .MenuTreeView.ImageList = MainMenu.ImageList1
         .BackColorTree = &H6D4016
         .NodeAdd , tvwChild, "Master", MyDDE.GetFieldByName("Nama Order"), "Master", , , True, , , True, , &HFCF1ED, &H6D4016
         sWCID = ""
         While Not RcCompDetail.DBRecordset.EOF
            If sWCID <> RcCompDetail.DBRecordset.Fields("Work Center") Then
               sWCID = RcCompDetail.DBRecordset.Fields("Work Center")
               .NodeAdd "Master", tvwChild, sWCID, sWCID, "biru", , , , , , True, , &HFCF1ED, &H6D4016
            End If
            .NodeAdd sWCID, tvwChild, RcCompDetail.DBRecordset.Fields("Kode Barang"), RcCompDetail.DBRecordset.Fields("Kode Barang"), "ijo", , , , , , True, , &HFCF1ED, &H6D4016
            RcCompDetail.DBRecordset.MoveNext
         Wend
      End With
   End If
Exit Sub
xErr:
   Err.Clear
End Sub
