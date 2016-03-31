VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FormWH 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Warehouse"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   1065
   ClientWidth     =   14730
   Icon            =   "FormWH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   5550
      TabIndex        =   93
      Top             =   105
      Visible         =   0   'False
      Width           =   4425
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   1005
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox PBFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EEDAC1&
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   3570
      ScaleHeight     =   1365
      ScaleWidth      =   4755
      TabIndex        =   98
      Top             =   6465
      Visible         =   0   'False
      Width           =   4785
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1830
         MaxLength       =   100
         TabIndex        =   22
         Top             =   510
         Width           =   2730
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3225
         TabIndex        =   24
         Top             =   900
         Width           =   1380
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1845
         TabIndex        =   23
         Top             =   900
         Width           =   1380
      End
      Begin VB.PictureBox PBHeader 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         FillColor       =   &H80000002&
         ForeColor       =   &H80000001&
         Height          =   405
         Left            =   -15
         ScaleHeight     =   405
         ScaleWidth      =   4770
         TabIndex        =   99
         Top             =   0
         Width           =   4770
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FIND PART"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   240
            Left            =   1860
            TabIndex        =   100
            Top             =   75
            Width           =   990
         End
      End
      Begin VB.Label LblFind 
         BackStyle       =   0  'Transparent
         Caption         =   "Criteria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         TabIndex        =   101
         Top             =   555
         Width           =   1620
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList IconList 
      Left            =   10995
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":209DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":2723C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":2DA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":34300
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":3AB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":413C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":47C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormWH.frx":4E488
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FRMMain 
      BackColor       =   &H00EAAF6F&
      Height          =   7215
      Left            =   75
      TabIndex        =   25
      Top             =   630
      Width           =   14685
      Begin VB.Frame FrLokasi 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Bin Content"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   3
         Left            =   11790
         TabIndex        =   66
         Top             =   1290
         Visible         =   0   'False
         Width           =   11310
         Begin VB.TextBox TxtContent 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            DataField       =   "Location_Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1035
            TabIndex        =   17
            Top             =   240
            Width           =   3090
         End
         Begin VB.TextBox TxtContent 
            BorderStyle     =   0  'None
            DataField       =   "Bin_Code"
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
            Left            =   1035
            TabIndex        =   19
            Top             =   960
            Width           =   3090
         End
         Begin VB.TextBox TxtContent 
            BorderStyle     =   0  'None
            DataField       =   "Description"
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
            Left            =   1035
            TabIndex        =   18
            Top             =   600
            Width           =   3090
         End
         Begin MSDataListLib.DataCombo DCFS 
            DataField       =   "Code"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   1410
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Description"
            BoundColumn     =   "Code"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCFS 
            DataField       =   "Code"
            Height          =   315
            Index           =   2
            Left            =   2235
            TabIndex        =   21
            Top             =   1410
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Description"
            BoundColumn     =   "Code"
            Text            =   ""
         End
         Begin VB.TextBox TxtContent 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            DataField       =   "itemname"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   7380
            TabIndex        =   68
            Top             =   330
            Width           =   3855
         End
         Begin VB.TextBox TxtContent 
            BorderStyle     =   0  'None
            DataField       =   "noitem"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   5340
            TabIndex        =   67
            Top             =   330
            Width           =   2000
         End
         Begin VB.TextBox TxtContent 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "STOCK"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00;(#.##0,00)"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   9015
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   1365
            Width           =   2220
         End
         Begin VB.TextBox TxtContent 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "SafetyStock"
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
            Height          =   285
            Index           =   9
            Left            =   9015
            TabIndex        =   73
            Top             =   1020
            Width           =   2220
         End
         Begin VB.TextBox TxtContent 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "ROP"
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
            Height          =   285
            Index           =   8
            Left            =   9015
            TabIndex        =   72
            Top             =   675
            Width           =   2220
         End
         Begin VB.TextBox TxtContent 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "Max_Qty"
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
            Height          =   285
            Index           =   7
            Left            =   5340
            TabIndex        =   71
            Top             =   1365
            Width           =   2000
         End
         Begin VB.TextBox TxtContent 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "Min_Qty"
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
            Height          =   285
            Index           =   6
            Left            =   5340
            TabIndex        =   70
            Top             =   1020
            Width           =   2000
         End
         Begin VB.TextBox TxtContent 
            BorderStyle     =   0  'None
            DataField       =   "UOM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   5340
            TabIndex        =   69
            Top             =   675
            Width           =   2000
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter Selection"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   26
            Left            =   2385
            TabIndex        =   105
            Top             =   1410
            Width           =   1425
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter Selection"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   25
            Left            =   270
            TabIndex        =   104
            Top             =   1410
            Width           =   1425
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
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
            Left            =   4365
            TabIndex        =   95
            Top             =   345
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lokasi"
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
            Index           =   15
            Left            =   240
            TabIndex        =   94
            Top             =   345
            Width           =   435
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock on Hand"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   24
            Left            =   7485
            TabIndex        =   92
            Top             =   1380
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Safety Stock"
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
            Index           =   20
            Left            =   7485
            TabIndex        =   81
            Top             =   1035
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Order Point"
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
            Index           =   19
            Left            =   7485
            TabIndex        =   80
            Top             =   690
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max.Qty"
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
            Index           =   18
            Left            =   4365
            TabIndex        =   79
            Top             =   1380
            Width           =   630
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min.Qty"
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
            Index           =   17
            Left            =   4365
            TabIndex        =   78
            Top             =   1035
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BIN Lokasi"
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
            Index           =   16
            Left            =   240
            TabIndex        =   77
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BIN Tipe"
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
            Left            =   240
            TabIndex        =   76
            Top             =   690
            Width           =   600
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
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
            Left            =   4365
            TabIndex        =   75
            Top             =   690
            Width           =   510
         End
      End
      Begin VB.Frame FrLokasi 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Bin Location "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   2
         Left            =   3315
         TabIndex        =   35
         Top             =   4440
         Visible         =   0   'False
         Width           =   11310
         Begin VB.TextBox TxtLoc 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            DataField       =   "Location_Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1335
            TabIndex        =   16
            Top             =   255
            Width           =   3375
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "View All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8655
            TabIndex        =   107
            Top             =   1410
            Width           =   1065
         End
         Begin MSDataListLib.DataCombo DCFS 
            DataField       =   "Code"
            Height          =   315
            Index           =   0
            Left            =   6660
            TabIndex        =   103
            Top             =   1410
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Description"
            BoundColumn     =   "Code"
            Text            =   ""
         End
         Begin VB.TextBox TxtLoc 
            BorderStyle     =   0  'None
            DataField       =   "Bin_Type"
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
            Index           =   7
            Left            =   6345
            TabIndex        =   97
            Top             =   1425
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.TextBox TxtLoc 
            BorderStyle     =   0  'None
            DataField       =   "Bin_Type"
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
            Index           =   6
            Left            =   6345
            TabIndex        =   96
            Top             =   1035
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.CommandButton CmdLookUp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4320
            Picture         =   "FormWH.frx":54CEA
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   653
            Width           =   345
         End
         Begin VB.TextBox TxtLoc 
            BorderStyle     =   0  'None
            DataField       =   "Code"
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
            Index           =   1
            Left            =   1335
            TabIndex        =   84
            Top             =   1035
            Width           =   3375
         End
         Begin VB.TextBox TxtLoc 
            BorderStyle     =   0  'None
            DataField       =   "Description"
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
            Index           =   2
            Left            =   1335
            TabIndex        =   85
            Top             =   1425
            Width           =   3375
         End
         Begin VB.TextBox TxtLoc 
            BorderStyle     =   0  'None
            DataField       =   "Bin_Type"
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
            Left            =   1335
            TabIndex        =   82
            Top             =   645
            Width           =   2985
         End
         Begin VB.TextBox TxtLoc 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "Max_Weight"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0;(0)"
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
            Index           =   5
            Left            =   6345
            TabIndex        =   87
            Top             =   645
            Width           =   3375
         End
         Begin VB.TextBox TxtLoc 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            DataField       =   "Bin_Ranking"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0;(0)"
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
            Index           =   4
            Left            =   6360
            TabIndex        =   86
            Top             =   255
            Width           =   3375
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter Selection"
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
            Index           =   12
            Left            =   5115
            TabIndex        =   102
            Top             =   1493
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max Weight"
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
            Index           =   5
            Left            =   5115
            TabIndex        =   59
            Top             =   713
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bin Ranking"
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
            Index           =   4
            Left            =   5115
            TabIndex        =   58
            Top             =   323
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe BIN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   57
            Top             =   698
            Width           =   600
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
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
            Left            =   240
            TabIndex        =   56
            Top             =   1493
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lokasi BIN"
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
            Index           =   1
            Left            =   240
            TabIndex        =   55
            Top             =   323
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
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
            Left            =   240
            TabIndex        =   54
            Top             =   1103
            Width           =   360
         End
      End
      Begin VB.Frame FrLokasi 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Bin Type "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   1
         Left            =   3570
         TabIndex        =   34
         Top             =   3615
         Visible         =   0   'False
         Width           =   11310
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "Receive"
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
            Index           =   4
            Left            =   6195
            TabIndex        =   111
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "Pick"
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
            Index           =   7
            Left            =   6195
            TabIndex        =   15
            Top             =   1365
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "Put_Away"
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
            Index           =   6
            Left            =   6195
            TabIndex        =   14
            Top             =   990
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "Ship"
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
            Index           =   5
            Left            =   6195
            TabIndex        =   13
            Top             =   615
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "bin_prefik"
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
            Left            =   1200
            TabIndex        =   12
            Top             =   1365
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "Description"
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
            Index           =   2
            Left            =   1200
            TabIndex        =   11
            Top             =   990
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "Location_Code"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Top             =   615
            Width           =   3375
         End
         Begin VB.TextBox TxtType 
            BorderStyle     =   0  'None
            DataField       =   "code"
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
            Left            =   1200
            TabIndex        =   9
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
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
            Index           =   6
            Left            =   240
            TabIndex        =   112
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BIN Prefix"
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
            Index           =   23
            Left            =   240
            TabIndex        =   90
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pick"
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
            Index           =   22
            Left            =   5010
            TabIndex        =   89
            Top             =   1365
            Width           =   270
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Put Away"
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
            Index           =   21
            Left            =   5010
            TabIndex        =   88
            Top             =   1020
            Width           =   690
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pengiriman"
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
            Index           =   10
            Left            =   5010
            TabIndex        =   63
            Top             =   675
            Width           =   780
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Penerimaan"
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
            Index           =   9
            Left            =   5010
            TabIndex        =   62
            Top             =   330
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
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
            Index           =   8
            Left            =   240
            TabIndex        =   61
            Top             =   1050
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lokasi"
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
            Index           =   7
            Left            =   240
            TabIndex        =   60
            Top             =   690
            Width           =   435
         End
      End
      Begin VB.Frame FrLokasi 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Warehouse Location "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Index           =   0
         Left            =   3240
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   11310
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Contact"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Height          =   315
            Index           =   5
            Left            =   6600
            MaxLength       =   25
            TabIndex        =   8
            Tag             =   "WH"
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "telpon"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Height          =   315
            Index           =   4
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "WH"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "kota"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Height          =   315
            Index           =   3
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "WH"
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "locations"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Height          =   315
            Index           =   2
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "WH"
            Top             =   1080
            Width           =   3585
         End
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "warehouse"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Height          =   315
            Index           =   0
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   1
            Tag             =   "WH"
            Top             =   360
            Width           =   3585
         End
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "warehouse name"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Height          =   315
            Index           =   1
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   2
            Tag             =   "WH"
            Top             =   720
            Width           =   3585
         End
         Begin VB.TextBox TxtWH 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Nama Kelompok Gudang"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00;(#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
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
            Index           =   6
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   4
            Tag             =   "WH"
            Top             =   1440
            Width           =   3225
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   4425
            Picture         =   "FormWH.frx":5B53C
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1448
            Width           =   330
         End
         Begin VB.Label lbltext 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Klasifikasi"
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
            Index           =   6
            Left            =   300
            TabIndex        =   108
            Top             =   1515
            Width           =   795
         End
         Begin VB.Label lbltext 
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
            Height          =   195
            Index           =   5
            Left            =   5370
            TabIndex        =   65
            Top             =   1170
            Width           =   1110
         End
         Begin VB.Label lbltext 
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
            Height          =   195
            Index           =   4
            Left            =   5370
            TabIndex        =   64
            Top             =   480
            Width           =   330
         End
         Begin VB.Label lbltext 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telepon"
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
            Left            =   5370
            TabIndex        =   33
            Top             =   825
            Width           =   570
         End
         Begin VB.Label lbltext 
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
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   32
            Top             =   1170
            Width           =   495
         End
         Begin VB.Label lbltext 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            Index           =   1
            Left            =   300
            TabIndex        =   31
            Top             =   830
            Width           =   405
         End
         Begin VB.Label lbltext 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
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
            Left            =   300
            TabIndex        =   30
            Top             =   480
            Width           =   360
         End
      End
      Begin MSComctlLib.TreeView TViewMenu 
         Height          =   6840
         Left            =   75
         TabIndex        =   26
         Top             =   225
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   12065
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
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
      Begin VB.Frame FrSetup 
         BackColor       =   &H00EAAF6F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5040
         Left            =   3240
         TabIndex        =   27
         Top             =   2040
         Width           =   11310
         Begin MSAdodcLib.Adodc DataTrans 
            Height          =   330
            Left            =   4980
            Top             =   2040
            Visible         =   0   'False
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid GridTrans 
            Height          =   4665
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   8229
            _Version        =   393216
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
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
               DataField       =   ""
               Caption         =   ""
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
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox PicBin 
         BackColor       =   &H00EAAF6F&
         Height          =   7425
         Left            =   2895
         ScaleHeight     =   7365
         ScaleWidth      =   12660
         TabIndex        =   36
         Top             =   225
         Width           =   12720
         Begin MSComctlLib.ListView LView 
            Height          =   3210
            Left            =   4230
            TabIndex        =   48
            Top             =   690
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   5662
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Bin Code"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Bin Type"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Location"
               Object.Width           =   1587
            EndProperty
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAAF6F&
            Caption         =   " Cari Data Barang Non BIN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   4260
            TabIndex        =   44
            Top             =   3945
            Width           =   3705
            Begin VB.CommandButton CmdFresh 
               Height          =   330
               Index           =   1
               Left            =   480
               Picture         =   "FormWH.frx":5B8C6
               Style           =   1  'Graphical
               TabIndex        =   110
               Tag             =   "True"
               ToolTipText     =   "Go"
               Top             =   600
               Width           =   345
            End
            Begin VB.CommandButton CmdFresh 
               Height          =   330
               Index           =   0
               Left            =   120
               Picture         =   "FormWH.frx":62118
               Style           =   1  'Graphical
               TabIndex        =   109
               Tag             =   "True"
               ToolTipText     =   "Refresh"
               Top             =   600
               Width           =   345
            End
            Begin VB.TextBox TxtCarik 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   870
               TabIndex        =   47
               Tag             =   "True"
               Top             =   608
               Width           =   2685
            End
            Begin VB.OptionButton OptSearch 
               BackColor       =   &H00EAAF6F&
               Caption         =   "&Kode"
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
               Left            =   1020
               TabIndex        =   46
               Tag             =   "True"
               Top             =   300
               Width           =   975
            End
            Begin VB.OptionButton OptSearch 
               BackColor       =   &H00EAAF6F&
               Caption         =   "&Nama"
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
               Index           =   1
               Left            =   240
               TabIndex        =   45
               Tag             =   "True"
               Top             =   300
               Value           =   -1  'True
               Width           =   780
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAAF6F&
            Caption         =   " Transfer Data"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   4245
            TabIndex        =   39
            Top             =   5055
            Width           =   3690
            Begin VB.CommandButton CmdPanah 
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Index           =   0
               Left            =   195
               TabIndex        =   43
               Tag             =   "True"
               ToolTipText     =   "Pilih 1 record"
               Top             =   330
               Width           =   825
            End
            Begin VB.CommandButton CmdPanah 
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Index           =   1
               Left            =   1020
               TabIndex        =   42
               Tag             =   "True"
               ToolTipText     =   "Pilih semua record"
               Top             =   330
               Width           =   825
            End
            Begin VB.CommandButton CmdPanah 
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Index           =   2
               Left            =   1845
               TabIndex        =   41
               Tag             =   "True"
               ToolTipText     =   "Pindah 1 record"
               Top             =   330
               Width           =   825
            End
            Begin VB.CommandButton CmdPanah 
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Index           =   3
               Left            =   2670
               TabIndex        =   40
               Tag             =   "True"
               ToolTipText     =   "Pindah semua record"
               Top             =   330
               Width           =   825
            End
         End
         Begin MSAdodcLib.Adodc DataProses2 
            Height          =   330
            Left            =   8055
            Top             =   5475
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc DataProses1 
            Height          =   330
            Left            =   122
            Top             =   5475
            Width           =   4000
            _ExtentX        =   7064
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid GridWizard 
            Height          =   5010
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Tag             =   "True"
            Top             =   300
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   8837
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
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
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
            EndProperty
         End
         Begin MSDataGridLib.DataGrid GridWizard 
            Height          =   5010
            Index           =   1
            Left            =   8055
            TabIndex        =   38
            Tag             =   "True"
            Top             =   300
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   8837
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
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
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
            EndProperty
         End
         Begin VB.Label LBLKateg 
            BackColor       =   &H00EAAF6F&
            Caption         =   "BIN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   4260
            TabIndex        =   106
            Top             =   375
            Width           =   3645
         End
         Begin VB.Label LBLKateg 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAAF6F&
            Caption         =   "List Barang Non BIN"
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
            Index           =   1
            Left            =   135
            TabIndex        =   53
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label LBLKateg 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAAF6F&
            Caption         =   "List Barang on BIN Lokasi : "
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
            Left            =   8070
            TabIndex        =   52
            Top             =   60
            Width           =   1950
         End
         Begin VB.Label LBLKateg 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAAF6F&
            Caption         =   "BIN Lokasi :"
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
            Left            =   4260
            TabIndex        =   51
            Top             =   60
            Width           =   840
         End
         Begin VB.Label LBLKateg 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAAF6F&
            Caption         =   "BIN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   10020
            TabIndex        =   50
            Top             =   60
            Width           =   285
         End
         Begin VB.Label LBLKateg 
            BackColor       =   &H00EAAF6F&
            Caption         =   "BIN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   5205
            TabIndex        =   49
            Top             =   60
            Width           =   2685
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SETUP WAREHOUSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   91
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "FormWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim sNodeKey As String
Dim strFilter, mCari As String
Dim VIn As Boolean
Dim SelectNode As Node
Dim strSQL As String
Dim vRole As String
Dim vSts As String
Dim rsWHouse As ADODB.Recordset
Dim rsBinLokasi As ADODB.Recordset
Dim rsBinContent As ADODB.Recordset
Dim rsBinType As ADODB.Recordset
Dim rsBIN As ADODB.Recordset
Dim rsBarang As ADODB.Recordset
    Dim rsFS As ADODB.Recordset
    Dim rsFS2 As ADODB.Recordset
    Dim rsFS3 As ADODB.Recordset
Dim myClass As New utility
Dim bkMark As Variant
Private mVarLastAccount, mVarGroupAccount       As String
Private Enum TransData
       adNew
       adEdit
       adCancel
       adDelete
       adSave
End Enum
Private RcPartner                       As New DBQuick
Private WithEvents mCall                As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mVarIndexStr  As String

Dim ChekAdd, chekwh, chekedit As Boolean
Dim chekbtype, chekAddBinType, chekeditbintype As Boolean
Dim chekbloc, chekaddbloc, chekeditbloc As Boolean
Dim chekbcont, chekeditbcont As Boolean
Dim mAdd As Boolean
Dim BinLocation As String
Dim rsPrint As ADODB.Recordset
Dim PrintPrev As New utility

Private Sub cmdLink_Click(Index As Integer)
OpenPartner 1
End Sub

Private Sub CmdLookUp_Click()
On Error GoTo 1
Screen.MousePointer = 11
Set FormLook.FormCaller = Me

Set FormLook.TextContainer = TxtLoc(6)
Set FormLook.TextContainer2 = TxtLoc(3)
Set FormLook.TextContainer3 = TxtLoc(7)
strSQL = "SELECT TOP 100 PERCENT Code AS Kode, Description AS [Tipe BIN], bin_prefik AS Prefix " & _
" From dbo.WHSE_BINTYPE WHERE (Location_Code = '" & TxtLoc(0).Text & "') ORDER BY Description"

'#KOLOM GRID CALLER TO BE INSERTED
FormLook.JudulForm = "Bin Type"
FormLook.SQLScript = strSQL
FormLook.ColRefNumber = 0
Load FormLook
FormLook.Show vbModal
TxtLoc(1).Text = AutoIndexAcc
Screen.MousePointer = 0
TxtLoc(2).SetFocus
Exit Sub
1:
MessageBox Err.Description, "formwh:cmdlookup_click " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub cmdOk_Click(Index As Integer)
On Error GoTo 1

If Index = 0 Then
    On Error GoTo FindErr
    Screen.MousePointer = vbHourglass
    
    Select Case sNodeKey
       Case "LOKWHSE"
            'LockObject TxtWH, TxtWH.Count, False
            With rsWHouse
                   If .Recordcount <> 0 And mCari <> "" And txtFind <> "" Then
                      .Find "[" & mCari & "] like '%" & txtFind & "%'", 0, adSearchForward, adBookmarkFirst
                      If .EOF Then
                         'messagebox "Field [" & mCari & "] With Criteria [" & txtFind & "] not found", vbExclamation, "Criteria"
                         MessageBox mCari & " = ' " & txtFind & " ' tidak ditemukan ..", vbInformation, "Find Part"
                         .MoveLast
                      Else
                         PBFind.Visible = False
                      End If
                   End If
            End With
       Case "BINTYPE"
              With rsBinType
                   If .Recordcount <> 0 And mCari <> "" And txtFind <> "" Then
                      .Find "[" & mCari & "] like '%" & txtFind & "%'", 0, adSearchForward, adBookmarkFirst
                      If .EOF Then
                         'messagebox "Field [" & mCari & "] With Criteria [" & txtFind & "] not found", vbExclamation, "Criteria"
                         MessageBox mCari & " = ' " & txtFind & " ' tidak ditemukan ..", vbInformation, "Find Part"
                         .MoveLast
                      Else
                         PBFind.Visible = False
                      End If
                   End If
              End With
       
       Case "BINLOKASI"
              With rsBinLokasi
                   If .Recordcount <> 0 And mCari <> "" And txtFind <> "" Then
                      .Find "[" & mCari & "] like '%" & txtFind & "%'", 0, adSearchForward, adBookmarkFirst
                      If .EOF Then
                         'messagebox "Field [" & mCari & "] With Criteria [" & txtFind & "] not found", vbExclamation, "Criteria"
                         MessageBox mCari & " = ' " & txtFind & " ' tidak ditemukan ..", vbInformation, "Find Part"
                         .MoveLast
                      Else
                         PBFind.Visible = False
                      End If
                   End If
              End With
       
       Case "BINCONT"
              With rsBinContent
                   If .Recordcount <> 0 And mCari <> "" And txtFind <> "" Then
                      .Find "[" & mCari & "] like '%" & txtFind & "%'", 0, adSearchForward, adBookmarkFirst
                      If .EOF Then
                         'messagebox "Field [" & mCari & "] With Criteria [" & txtFind & "] not found", vbExclamation, "Criteria"
                         MessageBox mCari & " = ' " & txtFind & " ' tidak ditemukan ..", vbInformation, "Find Part"
                         .MoveLast
                      Else
                         PBFind.Visible = False
                      End If
                   End If
              End With
    End Select
    PBFind.Visible = False
    'LockTombol adNew
    Screen.MousePointer = 0
    Exit Sub
    
Else
    PBFind.Visible = False
    Exit Sub
End If


'If Index = 0 Then
'    On Error GoTo FindErr
'    Screen.MousePointer = vbHourglass
'    If Left(SelectNode.Key, 1) = "C" Then
'        sNodeKey = SelectNode.Parent.Key
'        TimerON "Find Part.." & SelectNode.Parent.Text & " " & SelectNode.Text
'    Else
'        sNodeKey = SelectNode.Key
'        TimerON "Find Part.." & SelectNode.Text
'    End If
'
'    Select Case sNodeKey
'       Case "LOKWHSE"
'            'LockObject TxtWH, TxtWH.Count, False
'
'       Case "BINTYPE"

'       Case "BINLOKASI"

'       Case "BINCONT"
'            On Error Resume Next
'            If Len(txtFind.Text) <> 0 Then
'                'DataProses1.Recordset.Filter = strFilter & " like '" & TxtCarik.Text & "*'"
'                DataTrans.Recordset.Filter = "nama" & " like '" & txtFind.Text & "*'"
'            Else
               'cmdOK(0).Value = True
'            End If
'            Err.Clear
'       Case Else
'       PBFind.Visible = False
'        Exit Sub
'    End Select
'    'LockTombol adNew
'    Screen.MousePointer = 0
'    Exit Sub
    
'    Select Case CmbJenis.ListIndex
'        Case 0: CariB CmbFilter.ListIndex, Trim(TKunci.Text)
'        Case 1: CariTS CmbFilter.ListIndex, Trim(TKunci.Text)
'        Case 2: CariAS CmbFilter.ListIndex, Trim(TKunci.Text)
'    End Select
'Else
'    PBFind.Visible = False
'    Exit Sub
'End If

FindErr:
    Screen.MousePointer = 0
Exit Sub
1:
MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
End Sub

Private Sub cmdView_Click()
DCFS(0).Text = ""
If TViewMenu.SelectedItem Is Nothing Then
   FrSetup.Visible = False
   Exit Sub
End If
PicBin.Visible = False
Select Case TViewMenu.SelectedItem.Key
    Case "BINLOKASI"
         Set DataTrans.Recordset = rsBinLokasi
         FrLokasi(2).Move 3225, 240, FrLokasi(2).width, FrLokasi(2).Height
         FrLokasi(2).Visible = True
         FrLokasi(2).ZOrder (0)
    Case Else
        If Not TViewMenu.SelectedItem.Parent Is Nothing Then
            Select Case TViewMenu.SelectedItem.Parent.Key
                Case "BINLOKASI"
                    Set rsBinLokasi = New ADODB.Recordset
                    OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
                    Set DataTrans.Recordset = rsBinLokasi
                End Select
        End If
        
End Select
Set SelectNode = TViewMenu.SelectedItem
GridLayout
GridTrans.Refresh
End Sub

Private Sub Command1_Click()



End Sub

Private Sub DataTrans_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'DataProses1.Caption = "Record  " & Format(pRecordset.AbsolutePosition, RecForm) & " of " & Format(pRecordset.RecordCount, RecForm)
'    DataTrans.Caption = "Record  " & Format(pRecordset.AbsolutePosition, RecForm) & " of " & Format(pRecordset.RecordCount, RecForm)

End Sub

Private Sub DCFS_Change(Index As Integer)
Dim I As Long
Dim strFilter
        If TViewMenu.SelectedItem Is Nothing Then
           FrSetup.Visible = False
           Exit Sub
        End If
        PicBin.Visible = False
        Select Case TViewMenu.SelectedItem.Key
            Case "LOKWHSE"
                Set DataTrans.Recordset = rsWHouse
                GridTrans.Refresh
                FrLokasi(0).Move 3225, 240, FrLokasi(0).width, FrLokasi(0).Height
                FrLokasi(0).Visible = True
                FrLokasi(0).ZOrder (0)
                LockObject TxtWH, TxtWH.Count, True
                For I = 0 To TxtWH.Count - 1
                    Set TxtWH(I).DataSource = rsWHouse
                Next
            Case "BINTYPE"
                Set DataTrans.Recordset = rsBinType
                FrLokasi(1).Move 3225, 240, FrLokasi(1).width, FrLokasi(1).Height
                FrLokasi(1).Visible = True
                FrLokasi(1).ZOrder (0)
            Case "BINLOKASI"
                 Set DataTrans.Recordset = rsBinLokasi
                 FrLokasi(2).Move 3225, 240, FrLokasi(2).width, FrLokasi(2).Height
                 FrLokasi(2).Visible = True
                 FrLokasi(2).ZOrder (0)
            Case "BINCONT"
                Set DataTrans.Recordset = rsBinContent
                FrLokasi(3).Move 3225, 240, FrLokasi(3).width, FrLokasi(3).Height
                FrLokasi(3).Visible = True
                FrLokasi(3).ZOrder (0)
        
            Case Else
        '        Debug.Print Node.Text
                If Not TViewMenu.SelectedItem.Parent Is Nothing Then
                    Select Case TViewMenu.SelectedItem.Parent.Key
                        Case "BINTYPE"
                            Set rsBinType = New ADODB.Recordset
                            OpenTable rsBinType, CNN, "Select * From WHSE_BINtype WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
        '                    Set GridTrans.DataSource = rsBinType
                            Set DataTrans.Recordset = rsBinType
                            FrLokasi(1).Move 3225, 240, FrLokasi(1).width, FrLokasi(1).Height
                            FrLokasi(1).Visible = True
                            FrLokasi(1).ZOrder (0)
                            LockObject TxtType, TxtType.Count, True
                            For I = 0 To TxtType.Count - 1
                                Set TxtType(I).DataSource = rsBinType
                            Next
                        Case "BINLOKASI"
                            Set rsBinLokasi = New ADODB.Recordset
                            strFilter = "Select * From V_BINLOCATION WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "') AND (Bin_Type = '" & DCFS(Index).Text & "')"
                            OpenTable rsBinLokasi, CNN, strFilter
        '                   Debug.Print rsBinLokasi.RecordCount
                            Set DataTrans.Recordset = rsBinLokasi
                            LockObject TxtLoc, TxtLoc.Count, True
                            FrLokasi(2).Move 3225, 240, FrLokasi(2).width, FrLokasi(2).Height
                            FrLokasi(2).Visible = True
                            FrLokasi(2).ZOrder (0)
                            For I = 0 To TxtLoc.Count - 1
                                Set TxtLoc(I).DataSource = rsBinLokasi
                            Next
                        Case "BINCONT"
                            Set rsBinContent = New ADODB.Recordset
                            Select Case Index
                                Case Is = 1:
                                    If Len(DCFS(Index + 1).BoundText) <> 0 Then
                                        strFilter = "SELECT * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "') AND (Bin_Code = '" & DCFS(Index + 1).BoundText & "') AND (Description = '" & DCFS(Index).Text & "')"
                                    Else
                                        strFilter = "SELECT * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "') AND (Description = '" & DCFS(Index).Text & "')"
                                    End If
                                Case Is = 2
                                    If Len(DCFS(Index - 1).Text) <> 0 Then
                                        strFilter = "SELECT * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "') AND (Bin_Code = '" & DCFS(Index).BoundText & "') AND (Description = '" & DCFS(Index - 1).Text & "')"
                                    Else
                                        strFilter = "SELECT * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "') AND (Bin_Code = '" & DCFS(Index).BoundText & "')"
                                    End If
                                End Select
                            'OpenTable rsBinContent, cnn, "Select * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
                            OpenTable rsBinContent, CNN, strFilter
                            Set DataTrans.Recordset = rsBinContent
                            FrLokasi(3).Move 3225, 240, FrLokasi(3).width, FrLokasi(3).Height
                            FrLokasi(3).Visible = True
                            FrLokasi(3).ZOrder (0)
                            LockObject TxtContent, TxtContent.Count, True
                            For I = 0 To TxtContent.Count - 1
                                Set TxtContent(I).DataSource = rsBinContent
                            Next
                        Case "BINCONTENTRY"
                            'BUKA MASTER BIN
                            myClass.Gelas True
                            Set rsBIN = New ADODB.Recordset
                            Debug.Print DCFS(1).BoundText
                            Select Case DCFS(1).Tag
                                Case "Bin Type":
                                        strSQL = "SELECT WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, WHSE_BIN.Max_Weight" & _
                                                " FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code" & _
                                                " WHERE (WHSE_BINTYPE.Description = '" & DCFS(1).Text & "')" & _
                                                " ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"
                                Case "Location":
                                        strSQL = "SELECT WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, WHSE_BIN.Max_Weight" & _
                                                " FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code" & _
                                                " WHERE (WHSE_BIN.Location_Code = '" & DCFS(1).Text & "')" & _
                                                " ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"
                            End Select
                            
                            rsBIN.CursorLocation = adUseClient
                            rsBIN.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
                            Debug.Print rsBIN.Recordcount
                            With rsBIN
                                LView.ListItems.Clear
                                Do While Not .EOF
                                    With LView.ListItems.Add(, , .Fields("Code").Value)
                                       .SubItems(1) = rsBIN.Fields("Description").Value
                                       .SubItems(2) = rsBIN.Fields("Location_Code").Value
                                    End With
                                    .MoveNext
                                Loop
                            End With
                            PicBin.Visible = True
                            PicBin.ZOrder (0)
                        
                            OptSearch_Click (1)
                            If rsBIN.Recordcount <> 0 Then
                                LView.ListItems(1).Selected = True
                                LView_ItemClick LView.ListItems(1)
                                LView.SetFocus
                                For I = 0 To 3
                                    CmdPanah(I).Enabled = True
                                Next
                            Else
                                LBLKateg(3).Caption = ""
                                LBLKateg(4).Caption = ""
                                Set GridWizard(0).DataSource = rsBIN
                                Set GridWizard(1).DataSource = rsBIN
                                DataProses1.Caption = ""
                                DataProses2.Caption = ""
                                For I = 0 To 3
                                    CmdPanah(I).Enabled = False
                                Next
                            End If
                            myClass.Gelas False
                    End Select
                End If
        End Select
        
        GridLayout
        GridTrans.Refresh
    
End Sub

Private Sub Form_Load()
Dim j As Byte
myClass.Gelas True

HiasFormManTell PicBin, Me

Set mCall = New frmCaller

MyDDE.SetPermissions = aksess.MayDo("WareHouse")

OpenTable rsWHouse, CNN, "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
OpenTable rsBinType, CNN, "Select code,location_code,description,receive,ship,put_away,pick,bin_prefik,timestamp From WHSE_BINtype "
OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION "
OpenTable rsBinContent, CNN, "Select * From V_BINCONTENT "
OpenTable rsBarang, CNN, "Select * From V_BINBARANG "


BuildTree
'GridLayout
'***** WIZARD INITIALIZE *****
'BUKA MASTER PROSES
Set rsBIN = New ADODB.Recordset
strSQL = " SELECT  WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, " & _
" WHSE_BIN.Max_Weight FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code " & _
" GROUP BY WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, WHSE_BIN.Max_Weight, " & _
" WHSE_BIN.Code ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"

'Debug.Print strSQL
rsBIN.CursorLocation = adUseClient
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
With rsBIN
   Do While Not .EOF
      With LView.ListItems.Add(, , .Fields("Code").Value)
         .SubItems(1) = rsBIN.Fields("Description").Value
         .SubItems(2) = rsBIN.Fields("Location_Code").Value
      End With
      .MoveNext
   Loop
End With
''strSQL = "SELECT KODE, NAMA, TIDAK_TERPAKAI, LOKASI_NEW From P_PRODUCT " & _
'" WHERE  (LOKASI_NEW IS NULL) AND (TIDAK_TERPAKAI = '1') ORDER BY KODE"
'strSQL = "SELECT P_PRODUCT.KODE, P_PRODUCT.NAMA, P_PRODUCT.KODE_SATUAN_KECIL AS UOM FROM WHSE_BINCONTENT RIGHT OUTER JOIN " & _
'" P_PRODUCT ON WHSE_BINCONTENT.Item_No = P_PRODUCT.KODE " & _
'" WHERE (P_PRODUCT.TIDAK_TERPAKAI = '1') AND (WHSE_BINCONTENT.Item_No IS NULL) ORDER BY P_PRODUCT.KODE"


strSQL = "SELECT Inventory.NoItem, Inventory.ItemName, Inventory.UOM AS UOM FROM WHSE_BINCONTENT RIGHT OUTER JOIN " & _
" Inventory ON WHSE_BINCONTENT.NoItem = Inventory.noItem " & _
" WHERE (WHSE_BINCONTENT.NoItem IS NULL) ORDER BY Inventory.noitem"

Set rsBIN = New ADODB.Recordset
rsBIN.CursorLocation = adUseClient
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
Set GridWizard(0).DataSource = rsBIN
Set DataProses1.Recordset = rsBIN

'BUKA GRID WITH BIN SELECTED
''strSQL = "SELECT KODE, NAMA, TIDAK_TERPAKAI, LOKASI_NEW From P_PRODUCT " & _
'" WHERE  (LOKASI_NEW IS NULL) AND (TIDAK_TERPAKAI = '1') ORDER BY KODE"
'strSQL = "SELECT     TOP 100 PERCENT dbo.WHSE_BINCONTENT.Item_No AS KODE, dbo.P_PRODUCT.NAMA, dbo.WHSE_BINCONTENT.Location_Code, " & _
'" dbo.WHSE_BINCONTENT.UOM FROM dbo.WHSE_BINCONTENT INNER JOIN dbo.P_PRODUCT ON dbo.WHSE_BINCONTENT.Item_No = dbo.P_PRODUCT.KODE " & _
'" WHERE (dbo.P_PRODUCT.TIDAK_TERPAKAI = '1') ORDER BY dbo.WHSE_BINCONTENT.Item_No"

strSQL = "SELECT     TOP 100 PERCENT dbo.WHSE_BINCONTENT.NoItem AS KODE, dbo.inventory.ItemName, dbo.WHSE_BINCONTENT.Location_Code, " & _
" dbo.WHSE_BINCONTENT.UOM FROM dbo.WHSE_BINCONTENT INNER JOIN dbo.inventory ON dbo.WHSE_BINCONTENT.NoItem = dbo.inventory.noitem " & _
" ORDER BY dbo.WHSE_BINCONTENT.NoItem"


'digunakan set backcolor di labelkateg
For j = 0 To 5
    LBLKateg(j).BackColor = &HEAAF6F
Next j

Set rsBIN = New ADODB.Recordset
rsBIN.CursorLocation = adUseClient
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
Set GridWizard(1).DataSource = rsBIN
Set DataProses2.Recordset = rsBIN

TViewMenu.Nodes(2).Selected = True
TviewMenu_NodeClick TViewMenu.SelectedItem
myClass.Gelas False
End Sub

Private Sub DataProses1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
DataProses1.Caption = "Record  " & Format(pRecordset.AbsolutePosition, RecForm) & " of " & Format(pRecordset.Recordcount, RecForm)
End Sub

Private Sub DataProses2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
DataProses2.Caption = "Record  " & Format(pRecordset.AbsolutePosition, RecForm) & " of " & Format(pRecordset.Recordcount, RecForm)
End Sub

Private Sub GridLayout()
If Left(SelectNode.Key, 1) = "C" Then
    sNodeKey = SelectNode.Parent.Key
Else
    sNodeKey = SelectNode.Key
End If
Select Case sNodeKey
   Case "LOKWHSE"
      With GridTrans
         .Columns(6).Visible = False
         .Columns(7).Visible = False
         .Columns(8).Visible = False
        ' .Columns(9).Visible = False
         .Columns(1).width = 1500
         .Columns(2).width = 1450
         .Columns(3).width = 1350
         .Columns(4).width = 1080
         .Columns(5).width = 1350
       '  .HoldFields
      End With
   Case "BINTYPE"
      With GridTrans
         '.Columns(0).Visible = False
         .Columns(7).Visible = False
         .Columns(8).Visible = False
         .Columns(1).width = 900
         .Columns(2).width = 1450
         .Columns(3).width = 2000
         .Columns(4).width = 1080
         .Columns(5).width = 1350
         .Columns(4).Alignment = dbgRight
         .Columns(5).Alignment = dbgRight
         .Columns(6).Alignment = dbgRight
         .Columns(7).Alignment = dbgRight
         .Columns(4).NumberFormat = QtyForm
         .Columns(5).NumberFormat = QtyForm
         .Columns(6).NumberFormat = QtyForm
         .Columns(7).NumberFormat = QtyForm
         
'         .HoldFields
      End With
   Case "BINLOKASI"
      With GridTrans
         .Columns(0).Visible = True
         .Columns(1).width = 900
         .Columns(2).width = 1450
         .Columns(3).width = 1350
         .Columns(4).width = 1080
         .Columns(5).width = 1350

'         .HoldFields
      End With
   Case "BINCONT"
      With GridTrans
         .Columns(0).Visible = False
        ' .Columns(8).Visible = False
         .Columns(1).width = 1000
         .Columns(2).width = 1000
         .Columns(3).width = 1350
         .Columns(4).width = 2150   'NAMABARANG
         .Columns(5).width = 600
         .Columns(6).width = 1000
         .Columns(7).width = 1000
         .Columns(8).width = 1000
         .Columns(9).width = 1000
         .Columns(10).width = 1200
         .Columns(6).Alignment = dbgRight
         .Columns(7).Alignment = dbgRight
         .Columns(8).Alignment = dbgRight
         .Columns(9).Alignment = dbgRight
         .Columns(10).Alignment = dbgRight
         '.Columns(11).Alignment = dbgRight
'         .HoldFields
      End With
   Case Else
      Exit Sub
End Select
With GridWizard(0)
    .Columns(0).width = 1100
    .Columns(1).width = 2350
'    .Columns(2).Visible = False
'    .Columns(3).Visible = False
    .HoldFields
End With
With GridWizard(1)
    .Columns(0).width = 1100
    .Columns(1).width = 2350
'    .Columns(2).Visible = False
'    .Columns(3).Visible = False
    .HoldFields
End With
End Sub

Private Sub ShowTreeview()
    TviewMenu_NodeClick TViewMenu.SelectedItem
End Sub

Private Function GetMainPurchase()
On Error GoTo 10
Dim VTb As ADODB.Recordset
    OpenTable VTb, CNN, "Select * From S_Org Where Main_Purchase = 'YES'"
    If VTb.EOF = False Then
       GetMainPurchase = VTb![Kode]
    Else
       GetMainPurchase = ""
    End If
    VTb.Close
    Set VTb = Nothing
Exit Function
10:
MessageBox Err.Description, "formwh:getmainpurchase " & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub BuildTree()
On Error GoTo 4
Dim VTbl As ADODB.Recordset
Dim vNode As Node
Dim sWHouse, sWHName As String

    TViewMenu.Nodes.Clear
    Set vNode = TViewMenu.Nodes.Add(, , "R", "SETUP WAREHOUSE")
    vNode.Expanded = True
    vNode.Bold = True
    Set vNode = TViewMenu.Nodes.Add("R", tvwChild, "LOKWHSE", "Warehouse Location")
    vNode.Bold = True
    Set vNode = TViewMenu.Nodes.Add("R", tvwChild, "BINTYPE", "BIN Type")
    vNode.Bold = True
    Set vNode = TViewMenu.Nodes.Add("R", tvwChild, "BINLOKASI", "BIN Location")
    vNode.Bold = True
    Set vNode = TViewMenu.Nodes.Add("R", tvwChild, "BINCONT", "BIN Content")
    vNode.Bold = True
    Set vNode = TViewMenu.Nodes.Add("R", tvwChild, "BINCONTENTRY", "BIN Content Entry")
    vNode.Bold = True
    
'    VWhereOrg = CekOrg(VUser![KODE_ORG])
'    Select Case sUserGroup
'        Case "ADMIN"
'            strSQL = "Select * From V_WAREHOUSE "
'        Case "ADMINFS"
'            strSQL = "SELECT * From V_WAREHOUSE WHERE (Code IN ('IGT', 'PU'))"
'        Case "ADMMTC"
'            strSQL = "Select * From V_WAREHOUSE Where CODE = '" & VUser.Fields("KODE_ORG").Value & "'"
'        Case Else
''            strSQL = "Select * From V_WAREHOUSE Where CODE = '" & VUser.Fields("KODE_ORG").Value & "'"
'    End Select

   ' VWhereOrg = CekOrg(VUser![KODE_ORG])
   ' Select Case sUserGroup
   Select Case "ADMIN"
        Case "ADMIN"
            strSQL = "Select * From WAREHOUSE "
'        Case "ADMINFS"
'            strSQL = "SELECT * From WAREHOUSE WHERE (warehouse IN ('FG Gdg 1', 'FIX Gdg 3'))"
'        Case "ADMMTC"
'            strSQL = "Select * From WAREHOUSE Where warehouse = '" & VUser.Fields("warehouse").Value & "'"
'        Case Else
'            strSQL = "Select * From WAREHOUSE Where warehouse = " & VUser.Fields("warehouse").Value & "'"
    End Select

    
   ' VNode.Expanded = True
    OpenTable vTransaction, CNN, strSQL
    With vTransaction
        While .EOF = False
'            sWHouse = .Fields("code").Value
'            sWHName = .Fields("name").Value & " (" & .Fields("code").Value & ")"
            sWHouse = .Fields("warehouse").Value
            sWHName = .Fields("warehouse name").Value & " (" & .Fields("warehouse").Value & ")"
            
            Set vNode = TViewMenu.Nodes.Add("BINTYPE", tvwChild, "CBINTYPE-" & sWHouse, sWHName)
            vNode.Tag = sWHouse
            Set vNode = TViewMenu.Nodes.Add("BINLOKASI", tvwChild, "CBINLOKASI-" & sWHouse, sWHName)
            vNode.Tag = sWHouse
            Set vNode = TViewMenu.Nodes.Add("BINCONT", tvwChild, "CBINCONT-" & sWHouse, sWHName)
            vNode.Tag = sWHouse
            Set vNode = TViewMenu.Nodes.Add("BINCONTENTRY", tvwChild, "CBINCONTENTRY-" & sWHouse, sWHName)
            vNode.Tag = sWHouse
            .MoveNext
            Set vNode = Nothing
        Wend
    End With
Exit Sub
4:
MessageBox Err.Description, "formwh:buildtree" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridTrans_HeadClick(ByVal ColIndex As Integer)
Select Case sNodeKey
    Case "LOKWHSE"
    Case "BINTYPE"
        rsBinType.Sort = GridTrans.Columns(ColIndex).Caption
    Case "BINLOKASI"
        rsBinLokasi.Sort = GridTrans.Columns(ColIndex).Caption
    Case "BINCONT"
        rsBinContent.Sort = GridTrans.Columns(ColIndex).Caption
    Case Else
        Exit Sub
End Select
End Sub

Private Sub GridTrans_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If GridTrans.col < 0 Then GridTrans.col = 0
mCari = GridTrans.Columns(GridTrans.col).Caption
LblFind = "Cari : " & mCari
End Sub

Private Sub GridWizard_HeadClick(Index As Integer, ByVal ColIndex As Integer)
If Index = 0 Then
'   If ColIndex = 0 Then
'      DataProses1.Recordset.Sort = "NoReg"
'   Else
      DataProses1.Recordset.Sort = GridWizard(0).Columns(ColIndex).Caption
'      Debug.Print GridWizard(0).Columns(ColIndex).Caption
'   End If
Else
'   If ColIndex = 1 Then
'      DataProses2.Recordset.Sort = "NoReg"
'   Else
      DataProses2.Recordset.Sort = GridWizard(1).Columns(ColIndex).Caption
'   End If
End If
End Sub

Private Sub LView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'DCFS(1).BoundText = ""
'Select Case ColumnHeader
'    Case Is = "Bin Type":
'        'Debug.Print LView.ColumnHeaders(2).Text
'        DCFS(1).BoundColumn = "": DCFS(1).DataField = "": DCFS(1).ListField = ""
'        'OpenTable rsFS, cnn, "SELECT     WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, " & _
'                    " WHSE_BIN.Max_Weight FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code " & _
'                    " WHERE (WHSE_BIN.Location_Code = '" & TViewMenu.SelectedItem.Tag & "') ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"
'        OpenTable rsFS, cnn, "SELECT Code, Location_Code,Description AS Deskripsi FROM WHSE_BINTYPE"
'        Set DCFS(1).RowSource = rsFS.DataSource
'
'        DCFS(1).BoundColumn = "Deskripsi": DCFS(1).DataField = "Deskripsi": DCFS(1).ListField = "Deskripsi"
'        lblLabels(25).Move 4245, 330, lblLabels(25).Width, lblLabels(25).Height: lblLabels(25).Visible = True: lblLabels(25).ZOrder 0
'        DCFS(1).Move 5790, 330, DCFS(1).Width, DCFS(1).Height: DCFS(1).Visible = True: DCFS(1).ZOrder 0
'        DCFS(1).Tag = ColumnHeader
'    Case Is = "Location":
'        'Debug.Print LView.ColumnHeaders(2).Text
'        DCFS(1).BoundColumn = "": DCFS(1).DataField = "": DCFS(1).ListField = ""
'        OpenTable rsFS, cnn, "SELECT Location_Code From WHSE_BINTYPE GROUP BY Location_Code"
'        Set DCFS(1).RowSource = rsFS.DataSource
'
'        DCFS(1).BoundColumn = "Location_Code": DCFS(1).DataField = "Location_Code": DCFS(1).ListField = "Location_Code"
'        lblLabels(25).Move 4245, 330, lblLabels(25).Width, lblLabels(25).Height: lblLabels(25).Visible = True: lblLabels(25).ZOrder 0
'        DCFS(1).Move 5790, 330, DCFS(1).Width, DCFS(1).Height: DCFS(1).Visible = True: DCFS(1).ZOrder 0
'        DCFS(1).Tag = ColumnHeader
'End Select
End Sub

Private Sub LView_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Debug.Print Item.Text
FilterProc Item.Text, Item.SubItems(2)
LBLKateg(3).Caption = Item & " : " & Item.SubItems(1)
LBLKateg(4).Caption = Item
LBLKateg(5).Caption = Item.SubItems(1) & " : " & BinLocation

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
MyDDE.GetFieldByName("Kode Kelompok Gudang") = mCall.GetFieldByName(0)
MyDDE.GetFieldByName("Nama Kelompok Gudang") = mCall.GetFieldByName(1)
TxtWH(6).Text = mCall.GetFieldByName(1)
If ChekAdd = True Then MyDDE.GetFieldByName("NoAccount") = MyAutoIndex
End Sub

Private Sub MEdit_Click(Index As Integer)
    Select Case Index
    Case 0: RunMenuAdd
    Case 1: RunMenuEdit
'    Case 2: RunMenuDelete
'    Case 7: FORMHistory.Show vbModal
'    Case 9: RunMenuNotApproved
    End Select
End Sub

Private Sub MFile_Click(Index As Integer)
    Select Case Index
    Case 0
    Case 1: RunMenuRefresh
    Case 3: RunMenuPrint
    Case 5: RunMenuExit
    End Select
End Sub

Private Sub MValid_Click(Index As Integer)
'    RunMenuValid
End Sub

Private Sub EditDetilKodeBin()
On Error GoTo 7
Dim rsBIN As New ADODB.Recordset
Dim KodeBin, KodeBin2 As String
Dim Dash As Long

'strSQL = " SELECT * From WHSE_BIN WHERE (Code LIKE '10%')"
strSQL = "Select * From WHSE_BIN"
rsBIN.CursorLocation = adUseClient
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockOptimistic, adCmdText
Screen.MousePointer = vbHourglass
rsBinLokasi.MoveFirst
With rsBIN
    Do While Not .EOF
'        Debug.Print .Fields("code").Value
        Dash = InStr(1, .Fields("code").Value, "-")
        KodeBin2 = Trim(Mid$(.Fields("code").Value, Dash + 1, Len(.Fields("code").Value)))
        KodeBin = Left(.Fields("code").Value, Dash)
        Select Case Len(KodeBin2)
            Case 1
                KodeBin = KodeBin + "00" + KodeBin2
            Case 2
                KodeBin = KodeBin + "0" + KodeBin2
            Case 3
                KodeBin = KodeBin + KodeBin2
        End Select
        .Fields("KodeBin").Value = KodeBin
        rsBinLokasi.MoveNext
        .MoveNext
    Loop
End With
Screen.MousePointer = 0
rsBinLokasi.Requery
Exit Sub
7:
MessageBox Err.Description, "frmwh:editdetilkodebin" & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub EditHeaderKodeBin()
On Error GoTo 8
Dim rsBIN As New ADODB.Recordset
Dim KodeBin, KodeBin2 As String
Dim Dash As Long

'strSQL = " SELECT * From WHSE_BIN WHERE (Code LIKE 'FLEXIBEL HOSE%')"
strSQL = "Select * From WHSE_BIN"
rsBIN.CursorLocation = adUseClient
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockOptimistic, adCmdText
Screen.MousePointer = vbHourglass
rsBinLokasi.MoveFirst
With rsBIN
    Do While Not .EOF
'        Debug.Print .Fields("code").Value
        Dash = InStr(1, .Fields("KodeBin").Value, "-")
        KodeBin2 = Trim(Mid$(.Fields("KodeBin").Value, Dash, Len(.Fields("KodeBin").Value)))
        KodeBin = Left(.Fields("KodeBin").Value, Dash - 1)
        If Val(KodeBin) <> 0 Then
            Select Case Len(KodeBin)
                Case 1
                    KodeBin = "0" + KodeBin + KodeBin2
                Case 2
                    KodeBin = KodeBin + KodeBin2
            End Select
            .Fields("KodeBin").Value = KodeBin
        End If
'        Debug.Print Val(KodeBin)
        rsBinLokasi.MoveNext
        .MoveNext
    Loop
End With
Screen.MousePointer = 0
rsBinLokasi.Requery
Exit Sub
8:
MessageBox Err.Description, "formwh:editheaderkodebin " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub MTools_Click(Index As Integer)
Select Case Index
'    Case 0: RunMenuFind
'    Case 2: RunMenuRefresh
'   ' Case 3: FormProduct.Show
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim Node As MSComctlLib.Node
On Error GoTo xErr
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If chekwh = True Then
                For I = 0 To TxtWH.Count - 1
                   TxtWH(I).Text = ""
                Next
                LockObject TxtWH, TxtWH.Count, False
                Set GridTrans.DataSource = rsWHouse
                TxtWH(0).SetFocus
                ChekAdd = True
                chekAddBinType = False
                chekaddbloc = False
                chekedit = False
                cmdLink(1).Enabled = True
             ElseIf chekbtype = True Then
                For I = 2 To TxtType.Count - 1
                   TxtType(I).Text = ""
                Next
                LockObject TxtType, TxtType.Count, False
                TxtType(0).Text = CounterKode: TxtType(0).Locked = True
                TxtType(1).Locked = True
                Set GridTrans.DataSource = rsBinType
                GridTrans.Columns(7).Visible = False
                GridTrans.Columns(8).Visible = False
                TxtType(2).SetFocus
                chekAddBinType = True
                chekeditbintype = False
                ChekAdd = False
                chekaddbloc = False
                chekeditbloc = False
                Inisialisasi (SelectNode.Tag)
              ElseIf chekbloc = True Then
                For I = 1 To TxtLoc.Count - 1
                   TxtLoc(I).Text = ""
                Next
                LockObject TxtLoc, TxtLoc.Count, False
                TxtLoc(0).Locked = True: TxtLoc(0).BackColor = &HC0FFFF
                Set GridTrans.DataSource = rsBinLokasi
                GridTrans.Columns(4).Visible = False
                TxtLoc(3).SetFocus
                chekaddbloc = True
                chekeditbloc = False
                ChekAdd = False
                chekAddBinType = False
                chekeditbintype = False
                InitBinLokasi (SelectNode.Tag)
             End If
             TViewMenu.Enabled = False
       Case tmbEdit:
            If chekwh = True Then
               LockObject TxtWH, TxtWH.Count, False
               chekedit = True
               ChekAdd = False
               chekeditbintype = False
               chekeditbloc = False
               cmdLink(1).Enabled = True
            ElseIf chekbtype = True Then
               LockObject TxtType, TxtType.Count, False
               TxtType(0).BackColor = &HC0FFFF: TxtType(1).BackColor = &HC0FFFF
               TxtType(0).Locked = True: TxtType(1).Locked = True:
               chekeditbintype = True
               chekAddBinType = False
               chekedit = False
               chekeditbloc = False
               
            ElseIf chekbloc = True Then
               LockObject TxtLoc, TxtLoc.Count, False
               TxtLoc(3).Locked = True: TxtLoc(1).Locked = True: TxtLoc(0).Locked = True
               chekeditbloc = True
               chekaddbloc = False
               chekedit = False
               chekeditbintype = False
            ElseIf chekeditbcont = True Then
                For I = 1 To TxtContent.Count - 1
                   TxtContent(I).BackColor = &HC0FFFF
                Next
                LockObject TxtContent, TxtContent.Count, False
                TxtContent(0).Locked = True: TxtContent(1).Locked = True: TxtContent(2).Locked = True: TxtContent(3).Locked = True: TxtContent(4).Locked = True: TxtContent(10).Locked = True
        
               
            End If
            TViewMenu.Enabled = False
       Case tmbDelete:
            If chekwh = True Then
                For I = 0 To TxtWH.Count - 1
                  TxtWH(I).Text = ""
                Next
                LockObject TxtWH, TxtWH.Count, True
               
                OpenTable rsWHouse, CNN, "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
                rsWHouse.Requery
                BuildTree  'digunakan untuk refresh TViewList
            ElseIf chekbtype = True Then
                 For I = 0 To TxtType.Count - 1
                  TxtType(I).Text = ""
                Next
                LockObject TxtType, TxtType.Count, True
                OpenTable rsBinType, CNN, "Select code,location_code,description,receive,ship,put_away,pick,bin_prefik,timestamp From WHSE_BINtype "
                rsBinType.Requery
                GridTrans.Columns(7).Visible = False
                GridTrans.Columns(8).Visible = False
            ElseIf chekbloc = True Then
                For I = 0 To TxtLoc.Count - 1
                  TxtLoc(I).Text = ""
                Next
                LockObject TxtLoc, TxtLoc.Count, True
                OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION"
                rsBinLokasi.Requery
                GridTrans.Columns(4).Visible = False
            End If
            

       Case tmbSave:
            If ChekAdd = True Then   'simpan ke warehouse
                MyDDE.SendDataToServer (" INSERT INTO WareHouse (WareHouse, [WareHouse Name], Locations,kota,telpon,contact,GroupAccount,NoAccount) " & _
                         " VALUES ('" & TxtWH(0).Text & "', '" & TxtWH(1).Text & "', N'" & TxtWH(2).Text & "',N'" & TxtWH(3).Text & "',N'" & TxtWH(4).Text & "',N'" & TxtWH(5).Text & "','" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "','" & MyDDE.GetFieldByName("NoAccount") & "')")
                MyDDE.SetPermissions = UserOk
                LockObject TxtWH, TxtWH.Count, True
                'OpenTable rsWHouse, CNN, "Select warehouse,[warehouse name],locations,kota,telpon,contact From warehouse Order By warehouse"
                OpenTable rsWHouse, CNN, "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
                rsWHouse.Requery
                ChekAdd = False
                cmdLink(1).Enabled = False
                BuildTree  'digunakan untuk refresh TViewList
            ElseIf chekAddBinType = True Then  'simpan ke whse_bintype
                MyDDE.SendDataToServer (" INSERT INTO WHSE_BINTYPE (code,location_code, description, bin_prefik,receive,ship,put_away,pick) " & _
                    " VALUES ('" & TxtType(0).Text & "', '" & TxtType(1).Text & "', N'" & TxtType(2).Text & "',N'" & TxtType(3).Text & "'," & TxtType(4).Text & "," & TxtType(5).Text & "," & TxtType(6).Text & "," & TxtType(7) & ")")
                MyDDE.SetPermissions = UserOk
                LockObject TxtType, TxtType.Count, True
                OpenTable rsBinType, CNN, "Select code,location_code,description,receive,ship,put_away,pick,bin_prefik,timestamp From WHSE_BINtype "
                rsBinType.Requery
                GridTrans.Columns(7).Visible = False
                GridTrans.Columns(8).Visible = False
                chekAddBinType = False
               
            ElseIf chekaddbloc = True Then  'simpan whse_bin   query V_binlocation
                MyDDE.SendDataToServer (" INSERT INTO WHSE_BIN (code,location_code, description, bin_type_code, bin_ranking, max_weight) " & _
                    " VALUES ('" & TxtLoc(1).Text & "', '" & TxtLoc(0).Text & "', N'" & TxtLoc(2).Text & "',N'" & TxtLoc(6).Text & "'," & TxtLoc(4).Text & "," & TxtLoc(5).Text & ")")
                MyDDE.SetPermissions = UserOk
                LockObject TxtLoc, TxtLoc.Count, True
                OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION "
                rsBinLokasi.Requery
                GridTrans.Columns(4).Visible = False
                chekaddbloc = False
                
            ElseIf chekedit = True Then '
                 MyDDE.SendDataToServer (" Update WareHouse set WareHouse = '" & TxtWH(0).Text & "' ," & _
                                    "[WareHouse Name]='" & TxtWH(1) & "', Locations = '" & TxtWH(2) & "',kota='" & TxtWH(3) & "',telpon='" & TxtWH(4) & "',contact='" & TxtWH(5) & "',GroupAccount='" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "',NoAccount='" & MyDDE.GetFieldByName("NoAccount") & "' where warehouse = '" & TxtWH(0).Text & "'")
                 MyDDE.SetPermissions = UserOk
                 LockObject TxtWH, TxtWH.Count, True
                 OpenTable rsWHouse, CNN, "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
                 rsWHouse.Requery
                 cmdLink(1).Enabled = False
                 BuildTree  'digunakan untuk refresh TViewList
             ElseIf chekeditbintype = True Then 'edit whse_BINTYPE
                 MyDDE.SendDataToServer (" Update WHSE_BINTYPE set code = '" & TxtType(0).Text & "' ," & _
                                    "location_code='" & TxtType(1) & "', description = '" & TxtType(2) & "',bin_prefik='" & TxtType(3) & "',receive=" & FQty(TxtType(4)) & ",ship=" & FQty(TxtType(5)) & ",put_away=" & FQty(TxtType(6)) & ",pick=" & FQty(TxtType(7)) & " where code = '" & TxtType(0).Text & "'")
                 MyDDE.SetPermissions = UserOk
                 LockObject TxtType, TxtType.Count, True
                 OpenTable rsBinType, CNN, "Select code,location_code,description,receive,ship,put_away,pick,bin_prefik,timestamp From WHSE_BINtype"
                 rsBinType.Requery
                 GridTrans.Columns(7).Visible = False
                 GridTrans.Columns(8).Visible = False
              ElseIf chekeditbloc = True Then   'edit whse_BIN  query BIN LOCATION
                 MyDDE.SendDataToServer (" Update WHSE_BIN set code = '" & TxtLoc(1).Text & "' ," & _
                                    "location_code='" & TxtLoc(0) & "', description = '" & TxtLoc(2) & "',bin_ranking=" & FQty(TxtLoc(4)) & ",max_weight=" & FQty(TxtLoc(5)) & " where code = '" & TxtLoc(1).Text & "'")
                 MyDDE.SetPermissions = UserOk
                 LockObject TxtType, TxtType.Count, True
                 OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION"
                 rsBinLokasi.Requery
                 GridTrans.Columns(4).Visible = False
                 
             ElseIf chekeditbcont = True Then   'edit bin content
'                 strSQL = " Update WHSE_BINCONTENT SET Location_Code = '" & TxtContent(0) & "', Bin_Code = '" & TxtContent(1) & "', NoItem = '" & TxtContent(3) & "', UOM = '" & TxtContent(5) & "', Min_QTY = " & FQty(TxtContent(6)) & ", Max_QTY = " & FQty(TxtContent(7)) & ", ROP = " & FQty(TxtContent(8)) & ", SafetyStock = " & FQty(TxtContent(9)) & " WHERE (IDX = '" & TakeKode(SelectNode.Tag) & "')"
                 MyDDE.SendDataToServer (" Update WHSE_BINCONTENT SET Location_Code = '" & TxtContent(0) & "', Bin_Code = '" & TxtContent(1) & "', noitem = '" & TxtContent(3) & "', UOM = '" & TxtContent(5) & "', Min_QTY = " & FQty(TxtContent(6)) & ", Max_QTY = " & FQty(TxtContent(7)) & ", ROP = " & FQty(TxtContent(8)) & ", SafetyStock = " & FQty(TxtContent(9)) & " WHERE IDX = '" & TakeKode(SelectNode.Tag) & "'")
                 MyDDE.SetPermissions = UserOk
                  LockObject TxtContent, TxtContent.Count, True
                 OpenTable rsBinContent, CNN, "Select * From V_BINCONTENT"
                 rsBinContent.Requery
                 GridTrans.Columns(0).Visible = False
             
             
            End If
            TViewMenu.Enabled = True
       Case tmbCancel:
            If ChekAdd = True Or chekedit = True Then
               LockObject TxtWH, TxtWH.Count, True
               OpenTable rsWHouse, CNN, "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
               rsWHouse.Requery
               Set GridTrans.DataSource = rsWHouse '
               cmdLink(1).Enabled = False
            ElseIf chekAddBinType = True Or chekeditbintype = True Then
                LockObject TxtType, TxtType.Count, True
                OpenTable rsBinType, CNN, "Select code,location_code,description,receive,ship,put_away,pick,bin_prefik,timestamp From WHSE_BINtype"
                rsBinType.Requery
                Set GridTrans.DataSource = rsBinType
                GridTrans.Columns(7).Visible = False
                GridTrans.Columns(8).Visible = False
            ElseIf chekaddbloc = True Or chekeditbloc = True Then
                LockObject TxtLoc, TxtLoc.Count, True
                OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION"
                rsBinLokasi.Requery
                Set GridTrans.DataSource = rsBinLokasi
                GridTrans.Columns(4).Visible = False
            ElseIf chekeditbcont = True Then
                LockObject TxtContent, TxtContent.Count, True
                OpenTable rsBinContent, CNN, "Select * From V_BINCONTENT"
                rsBinContent.Requery
                Set GridTrans.DataSource = rsBinContent
                'GridTrans.Columns(4).Visible = False
            End If
            TViewMenu.Enabled = True
       Case tmbNextRecord:
            If chekwh = True Then
                next_WH
            ElseIf chekbtype = True Then
                next_BinType
            ElseIf chekbloc = True Then
                next_BinLocation
            End If
       Case tmbPreviousRecord:
             If chekwh = True Then
                next_WH
            ElseIf chekbtype = True Then
                next_BinType
            ElseIf chekbloc = True Then
                next_BinLocation
            End If
       Case tmbBottomRecord:
             If chekwh = True Then
                next_WH
            ElseIf chekbtype = True Then
                next_BinType
            ElseIf chekbloc = True Then
                next_BinLocation
            End If
       Case tmbTopRecord:
             If chekwh = True Then
                next_WH
            ElseIf chekbtype = True Then
                next_BinType
            ElseIf chekbloc = True Then
                next_BinLocation
            End If
       Case Else:
     
End Select
GridLayout
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'On Error Resume Next
'PrepareQuery
'Err.Clear
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If chekwh = True Then
               OpenDBWH
             End If
       Case tmbEdit:
'            LockObject TxtWH, TxtWH.Count, False
       Case tmbDelete:
            If chekwh = True Then
                MyDDE.PrepareDelete = " DELETE FROM WareHouse WHERE   (WareHouse = N'" & TxtWH(0) & "') "
                Set GridTrans.DataSource = rsWHouse
                For I = 0 To TxtWH.Count - 1
                    Set TxtWH(I).DataSource = rsWHouse
                Next
            ElseIf chekbtype = True Then
                MyDDE.PrepareDelete = " DELETE FROM WHSE_BINtype WHERE   (code = N'" & TxtType(0) & "') "
                Set GridTrans.DataSource = rsBinType
                For I = 0 To TxtType.Count - 1
                    Set TxtType(I).DataSource = rsBinType
                Next
            ElseIf chekbloc = True Then
                MyDDE.PrepareDelete = " DELETE FROM WHSE_BIN WHERE   (code = N'" & TxtLoc(1) & "') "
                For I = 0 To TxtLoc.Count - 1
                    'Set TxtLoc(i).DataSource = MyDDE.ActiveRecordset
                    Set TxtLoc(I).DataSource = rsBinLokasi
                Next
            End If
            TViewMenu.Enabled = True
       Case tmbSave:
End Select
End Sub

Private Sub OptSearch_Click(Index As Integer)
If Index = 0 Then
   strFilter = "NoItem"
Else
   strFilter = "ItemName"
End If
TxtCarik.SetFocus
TxtCarik.Text = ""
End Sub

Private Sub PBHeader_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   ' MovingPicBox PBFind
End Sub

Private Sub SemeruOleDC1_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)

End Sub



'Private Sub RunFilter()
'    FormFilter.SQLScript = "SELECT * From Tools_Filter WHERE (FORM_NAME = N'" & mVarFormTag & "')"
'    FormFilter.ObjTableName = mVarObjTable
'    FormFilter.Show vbModal
'    If Trim(FormFilter.FilterTag) <> "" Then
'        Screen.ActiveForm.SQLSource = "SELECT * FROM " & mVarObjTable & FormFilter.FilterTag
'        Screen.ActiveForm.LoadData
'    End If
'End Sub
Private Sub RunMenuCancel()
LockTombol adCancel
TViewMenu.Enabled = True
Select Case UCase(SelectNode.Parent)
    Case "BIN TYPE"
        ProsesTipe adCancel
    Case "BIN LOCATION"
        ProsesLokasi adCancel
    Case Else
End Select

Select Case sNodeKey
   Case "LOKWHSE"
        LockObject TxtWH, TxtWH.Count, True
        GridTrans.Refresh
   Case "BINTYPE"
        LockObject TxtType, TxtType.Count, True
   Case "BINLOKASI"
        LockObject TxtLoc, TxtLoc.Count, True
   Case "BINCONT"
        LockObject TxtContent, TxtContent.Count, True
   Case Else
      Exit Sub
End Select

End Sub

Private Sub RunMenuHapus()
Select Case UCase(SelectNode.Parent)
    Case "BIN TYPE"
        ProsesTipe adDelete
    Case "BIN LOCATION"
        ProsesLokasi adDelete
    Case Else
End Select
LockTombol adDelete
End Sub

Private Sub RunMenuAdd()
On Error GoTo AddErr
Screen.MousePointer = vbHourglass
If Left(SelectNode.Key, 1) = "C" Then
    sNodeKey = SelectNode.Parent.Key
    TimerON "Editing.." & SelectNode.Parent.Text & " " & SelectNode.Text
Else
    sNodeKey = SelectNode.Key
    TimerON "Editing.." & SelectNode.Text
End If

Select Case sNodeKey
   Case "LOKWHSE"
        LockObject TxtWH, TxtWH.Count, False
   Case "BINTYPE"
        For I = 0 To TxtType.Count - 1
            Set TxtType(I).DataSource = Nothing
        Next
'        ProsesTipe adNew
'        If UCase(VGUserID) = "ADMIN" Or UCase(VGUserID) = "SA" Then
'            LockClearObject TxtType, TxtType.Count, 0, False
'        Else
            LockClearObject TxtType, TxtType.Count, 2, False
'        End If
        Inisialisasi (SelectNode.Tag)
   Case "BINLOKASI"
        For I = 0 To TxtLoc.Count - 1
            Set TxtLoc(I).DataSource = Nothing
        Next
'       ' If UCase(VGUserID) = "ADMIN" Or UCase(VGUserID) = "SA" Then
'            LockClearObject TxtLoc, TxtLoc.Count, 0, False
'        Else
            LockClearObject TxtLoc, TxtLoc.Count, 4, False
            TxtLoc(2).Locked = False: TxtLoc(2).BackColor = &HFFFFFF: TxtLoc(2).Text = ""
'        End If
        InitBinLokasi (SelectNode.Tag)
   Case "BINCONT"
        Dim oNode As Node
        Dim sNodeEntry As String
'        Debug.Print TViewMenu.SelectedItem.Tag
        sNodeEntry = "CBINCONTENTRY-" & SelectNode.Tag
        Text1 = sNodeEntry
        Set oNode = TViewMenu.Nodes(sNodeEntry)
        TViewMenu.Nodes(sNodeEntry).Selected = True
        Set TViewMenu.SelectedItem = oNode
        TviewMenu_NodeClick oNode
   Case Else
      Exit Sub
End Select
LockTombol adNew
TViewMenu.Enabled = False
Screen.MousePointer = 0
Exit Sub

AddErr:
    Screen.MousePointer = 0
    MessageBox Err.Description, vbExclamation, App.ProductName
    
End Sub

Private Sub RunMenuSave()
On Error GoTo 21
LockTombol adSave
TViewMenu.Enabled = True
If MessageBox("Simpan data " & SelectNode.Parent.Text & " ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
    Select Case sNodeKey
       Case "LOKWHSE"
            LockObject TxtWH, TxtWH.Count, True
       Case "BINTYPE"
            LockObject TxtType, TxtType.Count, True
            If TxtType(0) = "" Or TxtType(1) = "" Or TxtType(2) = "" Or TxtType(3) = "" Then Exit Sub
'           ' If Tbar.Tag = "ADD" Then
'                ProsesTipe adNew
'            ElseIf Tbar.Tag = "EDIT" Then
'                ProsesTipe adEdit
'            End If
       Case "BINLOKASI"
            LockObject TxtLoc, TxtLoc.Count, True
'            If Tbar.Tag = "ADD" Then
'                ProsesLokasi adNew
'            ElseIf Tbar.Tag = "EDIT" Then
'                ProsesLokasi adEdit
'            End If
'       Case "BINCONT"
'            LockObject TxtContent, TxtContent.Count, True
'            If Tbar.Tag = "ADD" Then
'                ProsesISI adNew
'            ElseIf Tbar.Tag = "EDIT" Then
'                ProsesISI adEdit
'            End If
       Case Else
          Exit Sub
    End Select
Else
    Select Case sNodeKey
       Case "LOKWHSE"
            LockObject TxtWH, TxtWH.Count, True
       Case "BINTYPE"
            LockObject TxtType, TxtType.Count, True
       Case "BINLOKASI"
            LockObject TxtLoc, TxtLoc.Count, True
       Case "BINCONT"
            LockObject TxtContent, TxtContent.Count, True
       Case Else
          Exit Sub
    End Select
End If
Exit Sub
21:
MessageBox Err.Description, "formwh:runmenusave " & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub RunMenuEdit()
'Dim sNodeText As String
'TimerON "Editing.." & SelectNode.Text
If Left(SelectNode.Key, 1) = "C" Then
    sNodeKey = SelectNode.Parent.Key
    TimerON "Editing.." & SelectNode.Parent.Text & " " & SelectNode.Text
Else
    sNodeKey = SelectNode.Key
    TimerON "Editing.." & SelectNode.Text
End If

Select Case sNodeKey
   Case "LOKWHSE"
        LockObject TxtWH, TxtWH.Count, False
        TxtWH(0).BackColor = &HC0FFFF
        TxtWH(0).Locked = True
   Case "BINTYPE"
        LockObject TxtType, TxtType.Count, False
        TxtType(0).BackColor = &HC0FFFF: TxtType(1).BackColor = &HC0FFFF
        TxtType(0).Locked = True: TxtType(1).Locked = True: 'TxtType(2).Locked = True : TxtType(3).Locked = True
   Case "BINLOKASI"
        LockObject TxtLoc, TxtLoc.Count, False
        TxtLoc(0).Locked = True: TxtLoc(1).Locked = True: TxtLoc(3).Locked = True
   Case "BINCONT"
        LockObject TxtContent, TxtContent.Count, False
        TxtContent(0).Locked = True: TxtContent(1).Locked = True: TxtContent(2).Locked = True: TxtContent(3).Locked = True: TxtContent(4).Locked = True: TxtContent(10).Locked = True
        TombolSiap True
   Case Else
      Exit Sub
End Select
LockTombol adEdit
TViewMenu.Enabled = False
End Sub
Private Sub LockTombol(TagTombol As TransData)
Dim I As Integer

'1  New
'2  Edit
'3  Cancel
'4  Save
'5  Delete
'7  Refresh
'8  Find
'10 Exit

'Select Case TagTombol
'    Case TransData.adNew
'        Tbar.Buttons(1).Enabled = False
'        Tbar.Buttons(2).Enabled = False
'        Tbar.Buttons(3).Enabled = True
'        Tbar.Buttons(4).Enabled = True
'        Tbar.Buttons(5).Enabled = False
'        Tbar.Buttons(7).Enabled = False
'        Tbar.Buttons(8).Enabled = False
'    Case TransData.adEdit
'        Tbar.Buttons(1).Enabled = False
'        Tbar.Buttons(2).Enabled = False
'        Tbar.Buttons(3).Enabled = True
'        Tbar.Buttons(4).Enabled = True
'        Tbar.Buttons(5).Enabled = False
'        Tbar.Buttons(7).Enabled = False
'        Tbar.Buttons(8).Enabled = False
'        GridTrans.Enabled = False
'    Case TransData.adCancel
'        For I = 1 To Tbar.Buttons.Count
'            Tbar.Buttons(I).Enabled = True
'        Next
'        GridTrans.Enabled = True
'    Case TransData.adDelete
'    Case TransData.adSave
'        For I = 1 To Tbar.Buttons.Count
'            Tbar.Buttons(I).Enabled = True
'        Next
'        GridTrans.Enabled = True
'End Select

End Sub

Private Sub RunMenuPrint()
'Dim WCon As String
'Dim I As Long
'
'   On Error GoTo ErrView
'   CRpt.Reset
'   CRpt.DiscardSavedData = True
'
'If TViewMenu.SelectedItem Is Nothing Then
'   FrSetup.Visible = False
'   Exit Sub
'End If
''    Set PrintPrev = New Utility
'
'Select Case TViewMenu.SelectedItem.key
'    Case "LOKWHSE"
''        Set DataTrans.Recordset = rsWHouse
''        strSQL = "Select * From V_WAREHOUSE Order By CODE"
''        PrintPrev.ReportFileName = "Laporan_Warehouse.rpt"
'        CRpt.SQLQuery = "Select * From V_WAREHOUSE Order By CODE"
'        'CRpt.SQLQuery = "SELECT * From R_BIN_CONTENT"  'strSQL = "SELECT * FROM " & (View) & (Seleksi)
'        CRpt.ReportFileName = App.Path & "\" & "Laporan_Warehouse.rpt"
'        'CRpt.ReportFileName = App.Path & "\" & "Bin_Location_85.rpt"
'    Case "BINTYPE"
''        Set DataTrans.Recordset = rsBinType
''        strSQL = "Select * From WHSE_BINTYPE "
''        PrintPrev.ReportFileName = "Laporan_Bin_Type.rpt"
'        CRpt.SQLQuery = "Select * From WHSE_BINTYPE"
'        CRpt.ReportFileName = App.Path & "\" & "Laporan_Bin_Type.rpt"
'    Case "BINLOKASI"
''        Set DataTrans.Recordset = rsBinLokasi
''        strSQL = "Select * From V_BINLOCATION "
''        PrintPrev.ReportFileName = "Laporan_Bin_Location.rpt"
'        CRpt.SQLQuery = "Select * From V_BINLOCATION "
'        CRpt.ReportFileName = App.Path & "\" & "Laporan_Bin_Location.rpt"
'    Case "BINCONT"
''        Set DataTrans.Recordset = rsBinContent
''        strSQL = "Select * From V_BINCONTENT "
''        PrintPrev.ReportFileName = "Laporan_Bin_Content.rpt"
'        CRpt.SQLQuery = "Select * From V_BINCONTENT ORDER BY Description"
'        CRpt.ReportFileName = App.Path & "\" & "Laporan_Bin_Content.rpt"
'    Case Else
'        If Not TViewMenu.SelectedItem.Parent Is Nothing Then
'            Select Case TViewMenu.SelectedItem.Parent.key
'                Case "BINTYPE"
'                    strSQL = "Select * From WHSE_BINTYPE WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
'                    CRpt.SQLQuery = strSQL
'                    Set DataTrans.Recordset = rsBinType
''                    PrintPrev.ReportFileName = "Laporan_Bin_Type.rpt"
'                    CRpt.ReportFileName = ""
'
''                Case "BINLOKASI"
''                    strSQL = "Select * From V_BINLOCATION WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
''                    Set DataTrans.Recordset = rsBinLokasi
''                    PrintPrev.ReportFileName = "Laporan_Bin_Location.rpt"
'
''                Case "BINCONT"
''                    strSQL = "Select * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
''                    Set DataTrans.Recordset = rsBinContent
''                    PrintPrev.ReportFileName = "Laporan_Bin_Content_Stock.rpt"
'
''                Case "BINCONTENTRY"
''                    Exit Sub
'                    'BUKA MASTER BIN
''                    strSQL = "SELECT WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, WHSE_BIN.Max_Weight " & _
''                            " FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code " & _
''                            " WHERE (WHSE_BIN.Location_Code = '" & TViewMenu.SelectedItem.Tag & "') ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"
''                Case Else
''                    strSQL = "SELECT * From R_BIN_CONTENT"
''                    PrintPrev.ReportFileName = "Bin_Location.rpt"
''
'            End Select
'        End If
'End Select
   
'   'WCon = "DSN=" & VGDataSource & ";DSQ=" & VGFPath & ";UID=" & VGUserID & ";PWD=" & VGPass & ""
'   CRpt.WindowTitle = TViewMenu.SelectedItem.Text   'UCase(StatusBar1.Panels(3).Text)
'  ' VGConnection = WCon
'
'   With CRpt
'
'     .Connect = VGConnection
''     .ReportFileName = VPathReport & "\" & VReport.Fields("REPORT_FILE").Value
''      .ReportFileName = VReport![Nama_File_Report]
'     .DiscardSavedData = False
'
'      .Destination = crptToWindow
'      .WindowState = crptMaximized
'      .WindowShowPrintSetupBtn = True
'      .WindowShowRefreshBtn = True
'      .WindowShowSearchBtn = True
'      .WindowShowGroupTree = True
'      .ReportTitle = TViewMenu.SelectedItem.Text   'UCase(LBLSubs(1).Caption)
'      .WindowAllowDrillDown = True
'   End With
'
'   CRpt.Action = 1
'
'   Exit Sub
'ErrView:
'    messagebox Err.Description, vbInformation, "Report Viewer"
'    Err.Clear

End Sub

'Private Sub RunMenuPrint()
'Dim I As Long

'If TViewMenu.SelectedItem Is Nothing Then
'   FrSetup.Visible = False
'   Exit Sub
'End If
'    Set PrintPrev = New Utility

'Select Case TViewMenu.SelectedItem.Key
'    Case "LOKWHSE"
'        Set DataTrans.Recordset = rsWHouse
'        strSQL = "Select * From V_WAREHOUSE Order By CODE"
'        PrintPrev.ReportFileName = "Laporan_Warehouse.rpt"

'    Case "BINTYPE"
'        Set DataTrans.Recordset = rsBinType
'        strSQL = "Select * From WHSE_BINTYPE "
'        PrintPrev.ReportFileName = "Laporan_Bin_Type.rpt"
        
'    Case "BINLOKASI"
'        Set DataTrans.Recordset = rsBinLokasi
'        strSQL = "Select * From V_BINLOCATION "
'        PrintPrev.ReportFileName = "Laporan_Bin_Location.rpt"

'    Case "BINCONT"
'        Set DataTrans.Recordset = rsBinContent
'        strSQL = "Select * From V_BINCONTENT "
'        PrintPrev.ReportFileName = "Laporan_Bin_Content.rpt"
        
'    Case Else
'        If Not TViewMenu.SelectedItem.Parent Is Nothing Then
'            Select Case TViewMenu.SelectedItem.Parent.Key
'                Case "BINTYPE"
'                    strSQL = "Select * From WHSE_BINTYPE WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
'                    Set DataTrans.Recordset = rsBinType
'                    PrintPrev.ReportFileName = "Laporan_Bin_Type.rpt"
                
'                Case "BINLOKASI"
'                    strSQL = "Select * From V_BINLOCATION WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
'                    Set DataTrans.Recordset = rsBinLokasi
'                    PrintPrev.ReportFileName = "Laporan_Bin_Location.rpt"
                
'                Case "BINCONT"
'                    strSQL = "Select * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TViewMenu.SelectedItem.Tag & "')"
'                    Set DataTrans.Recordset = rsBinContent
'                    PrintPrev.ReportFileName = "Laporan_Bin_Content_Stock.rpt"
                
'                Case "BINCONTENTRY"
'                    Exit Sub
                    'BUKA MASTER BIN
'                    strSQL = "SELECT WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, WHSE_BIN.Max_Weight " & _
'                            " FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code " & _
'                            " WHERE (WHSE_BIN.Location_Code = '" & TViewMenu.SelectedItem.Tag & "') ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"
'                Case Else
'                    strSQL = "SELECT * From R_BIN_CONTENT"
'                    PrintPrev.ReportFileName = "Bin_Location.rpt"
'
'            End Select
'        End If
'End Select
'Set SelectNode = TViewMenu.SelectedItem
'
'    strSQL = "SELECT * From R_BIN_CONTENT" 'strSQL = "SELECT * FROM " & (View) & (Seleksi)
'    PrintPrev.ReportQuery = strSQL
'    PrintPrev.ReportFileName = "Bin_Location_new2.rpt"
'
'    PrintPrev.ReportLocation = App.Path  '"C:\ENKEI"
    'PrintPrev.ReportTitle = "Judul"
'
'    If Trim(PrintPrev.GetReportFileName) <> "" And Trim(PrintPrev.GetReportQuery) <> "" Then
'        PrintPrev.CallReportViewer
'    End If
'
'End Sub

Private Sub Preview()
On Error GoTo Hell
'Dim RcTes As New DBQuick
Dim rsPrint As ADODB.Recordset

Set rsPrint = New Recordset
  If CekGridEmpty = False Then
    'strSQL = " SELECT * FROM [" & rsReport.Fields("ViewObject").Value & "]" & ScanFilter2
    strSQL = "SELECT * From R_BIN_CONTENT"
    'strSQL = strSQL & mVarTmp
  Else
'    strSQL = " SELECT * FROM [" & rsReport.Fields("ViewObject").Value & "]"
    strSQL = "SELECT * From R_BIN_CONTENT"
  End If
  
    rsPrint.CursorLocation = adUseClient
    rsPrint.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
'  RcTes.DBOpen strSQL, CNN ', lckLockBatch, lckLockSync
'  ReportPos = PathRPT
    Set PrintPrev = New utility
  If rsPrint.Recordcount <> 0 Then
    'myPart.CallReportView strSQL, rsReport.Fields("FileNameReport").Value, ReportPos, rsReport.Fields("Description").Value 'App.Path & "\" & "Report"
'    PrintPrev.CallReportView strSQL, "Copy of Bin_Location_new.rpt", "C:\ENKEI\", "Judul"
    PrintPrev.ReportQuery = strSQL
    PrintPrev.ReportFileName = "Copy of Bin_Location_new.rpt"
    PrintPrev.ReportLocation = "C:\ENKEI"
    PrintPrev.ReportTitle = "Judul"
    PrintPrev.CallReportViewer
'    CRpt.Reset
'    CRpt.DiscardSavedData = True
'
'    CRpt.WindowTitle = UCase(StatusBar1.Panels(3).Text)
'    With CRpt
'        .Connect = cnn
'        Debug.Print cnn
'        .SQLQuery = strSQL

'        .ReportFileName = "C:\ENKEI\Copy of Bin_Location_new.rpt"  'VPathReport & "\" & VReport.Fields("REPORT_FILE").Value
        '      .ReportFileName = VReport![Nama_File_Report]
'        CRpt.DiscardSavedData = False
'        .Destination = crptToWindow
'        .WindowState = crptMaximized
'        .WindowShowPrintSetupBtn = True
'        .WindowShowRefreshBtn = True
'        .WindowShowSearchBtn = True
'        .WindowShowGroupTree = True
'        .ReportTitle = "Judul" 'UCase(LBLSubs(1).Caption)
'        .WindowAllowDrillDown = True
'    End With
'    CRpt.Action = 0
  Else
'    messagebox "Laporan Belum Ada Datanya. Harap Diperiksa Filter Kriterianya", "Peringatan", msgOkOnly
  End If

  rsPrint.Close
  Exit Sub
Hell:


  MessageBox Err.Description, vbExclamation, App.ProductName
  Err.Clear
End Sub

Private Function CekGridEmpty() As Boolean
On Error GoTo 5
'  Dim nCount As Integer
'  For nCount = 1 To ListFilter.ListItems.Count
'    If ListFilter.ListItems(nCount).Checked = True Then cekListKosong = False
'  Next
Dim idx As Integer
For idx = 0 To 4
    If Not IsNull(GridTrans.Columns(idx).Value) Then CekGridEmpty = True
Next idx
Exit Function
5:
MessageBox Err.Description, "frmwh:cekgridempty" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub RunMenuFind()
'    PBFind.Tag = "Find"
       PBFind.Left = (Me.ScaleWidth / 2) - (PBFind.width / 2)
       PBFind.Top = (Me.Height / 2) - (PBFind.Height / 2)
       PBFind.Visible = True
       txtFind.SetFocus
End Sub

Private Sub RunMenuRefresh()
'OpenTable rsWHouse, cnn, "Select * From V_WAREHOUSE Order By CODE"
'OpenTable rsBinLokasi, cnn, "Select * From V_BINLOCATION "
'"SELECT * From WHSE_BIN WHERE (Code LIKE 'FLEXIBEL HOSE%') ORDER BY KODEBIN" '
'OpenTable rsBinContent, cnn, "Select * From V_BINCONTENT "
'OpenTable rsBinType, cnn, "Select * From WHSE_BINtype "
'OpenTable rsBinContent, cnn, "Select * From V_BINCONTENT "
'Select Case SelectNode.Key
'    Case "LOKWHSE"
'        OpenTable rsWHouse, cnn, "Select * From V_WAREHOUSE Order By CODE"
'        Set DataTrans.Recordset = rsWHouse
'    Case "BINTYPE"
'        OpenTable rsBinType, cnn, "Select * From WHSE_BINtype "
'        Set DataTrans.Recordset = rsBinType
'    Case "BINLOKASI"
'        OpenTable rsBinLokasi, cnn, "Select * From V_BINLOCATION "
'        Set DataTrans.Recordset = rsBinLokasi
'    Case "BINCONT"
'        OpenTable rsBinContent, cnn, "Select * From V_BINCONTENT "
'        Set DataTrans.Recordset = rsBinContent
'End Select
    myClass.Gelas True
'    DataTrans.Recordset.Requery
    myClass.Gelas False
    TombolSiap False
End Sub

Private Sub RunMenuExit()

myClass.CloseRcset rsWHouse
myClass.CloseRcset rsBinLokasi
myClass.CloseRcset rsBinContent
myClass.CloseRcset rsBinType
myClass.CloseRcset rsBIN
myClass.CloseRcset rsBarang
CNN.Close
Set CNN = Nothing
'End
Unload Me
End Sub

Private Sub TviewMenu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then RunMenuExit
End Sub

Private Sub TviewMenu_NodeClick(ByVal Node As MSComctlLib.Node)
'Dim vLst As ListItem
'Dim VTb As ADODB.Recordset
Dim I As Long

'Text1 = UCase(Node.Key & "-" & Node.Tag)
If Left(Node.Key, 1) <> "C" Then
'    Label1(0).Font = "Arial"
'    Label1(0).FontBold = True
'    Label1(0).FontSize = 14
'    Label1(0).Caption = Node.Text
    Label1(1).Font = "Arial"
    Label1(1).FontBold = True
    Label1(1).FontSize = 14
    Label1(1).ForeColor = &HFF0000
    Label1(1).Caption = Node.Text
Else
'    Label1(0).Font = "Arial"
'    Label1(0).FontBold = True
'    Label1(0).FontSize = 14
'    Label1(0).Caption = Node.Parent & " - " & Node.Text
    Label1(1).Font = "Arial"
    Label1(1).FontBold = True
    Label1(1).FontSize = 14
    Label1(1).ForeColor = &HFF0000
    Label1(1).Caption = Node.Parent & " - " & Node.Text
End If
If Node Is Nothing Then
   FrSetup.Visible = False
   Exit Sub
End If
PicBin.Visible = False
Select Case UCase(Node.Key)
    Case "LOKWHSE"
        chekbloc = False
        chekbtype = False
        chekeditbcont = False
        chekwh = True
        OpenDBWH
        OpenTable rsWHouse, CNN, "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
        Set DataTrans.Recordset = rsWHouse
        Set GridTrans.DataSource = rsWHouse
        GridTrans.Refresh
        FrLokasi(0).Move 3225, 240, FrLokasi(0).width, FrLokasi(0).Height
        FrLokasi(0).Visible = True
        FrLokasi(0).ZOrder (0)
        LockObject TxtWH, TxtWH.Count, True
        For I = 0 To TxtWH.Count - 1
            Set TxtWH(I).DataSource = rsWHouse
        Next
        MyDDE.SetPermissions = UserOk
    Case "BINTYPE"
        chekbloc = False
        chekwh = False
        chekeditbcont = False
        chekbtype = True
        OpenDBTYPE
        Set DataTrans.Recordset = rsBinType
        Set GridTrans.DataSource = rsBinType
        GridTrans.Refresh
        FrLokasi(1).Move 3225, 240, FrLokasi(1).width, FrLokasi(1).Height
        FrLokasi(1).Visible = True
        FrLokasi(1).ZOrder (0)
        MyDDE.SetPermissions = UserOk
    Case "BINLOKASI"
         chekwh = False
         chekbtype = False
         chekeditbcont = False
         chekbloc = True
         OpenDBINLOC
         Set DataTrans.Recordset = rsBinLokasi
         Set GridTrans.DataSource = rsBinLokasi
         GridTrans.Refresh
         GridTrans.Columns(4).Visible = False
         FrLokasi(2).Move 3225, 240, FrLokasi(2).width, FrLokasi(2).Height
         FrLokasi(2).Visible = True
         FrLokasi(2).ZOrder (0)
         MyDDE.SetPermissions = UserOk
    Case "BINCONT"
         chekwh = False
         chekbtype = False
         chekbloc = False
         chekeditbloc = False
         chekeditbcont = True
         OpenDBINCONT
         Set DataTrans.Recordset = rsBinContent
         Set GridTrans.DataSource = rsBinContent
         GridTrans.Refresh
         FrLokasi(3).Move 3225, 240, FrLokasi(3).width, FrLokasi(3).Height
         FrLokasi(3).Visible = True
         FrLokasi(3).ZOrder (0)
         MyDDE.SetPermissions = UserAddnewDenied
        
    Case Else
  
'        Debug.Print Node.Text
        If Not Node.Parent Is Nothing Then
            Select Case Node.Parent.Key
                Case "BINTYPE"
                      chekbloc = False
                      chekwh = False
                      chekeditbcont = False
                      chekbtype = True
                    Set rsBinType = New ADODB.Recordset
                    OpenTable rsBinType, CNN, "Select * From WHSE_BINtype WHERE (Location_Code = '" & Node.Tag & "')"
                   ' OpenTable rsBinType, cnn, "Select * From WHSE_BINtype WHERE (Location_Code = 'ASML')"
                  '  Set GridTrans.DataSource = rsBinType
                    If rsBinType.Recordcount > 0 Then
                        Set DataTrans.Recordset = rsBinType
                        Set GridTrans.DataSource = rsBinType
                        GridTrans.Refresh
                        FrLokasi(1).Move 3225, 240, FrLokasi(1).width, FrLokasi(1).Height
                        FrLokasi(1).Visible = True
                        FrLokasi(1).ZOrder (0)
                        LockObject TxtType, TxtType.Count, True
                        For I = 0 To TxtType.Count - 1
                            Set TxtType(I).DataSource = rsBinType
                        Next
                    Else  ' digunakan apabila record kosong agar bisa di input
                        Set DataTrans.Recordset = rsBinType
                        Set GridTrans.DataSource = rsBinType
                        GridTrans.Refresh
                        FrLokasi(1).Move 3225, 240, FrLokasi(1).width, FrLokasi(1).Height
                        FrLokasi(1).Visible = True
                        FrLokasi(1).ZOrder (0)
                        LockObject TxtType, TxtType.Count, True
                        For I = 0 To TxtType.Count - 1
                             TxtType(I).Text = ""
                        Next
                    End If
                Case "BINLOKASI"
                    chekwh = False
                    chekbtype = False
                    chekeditbcont = False
                    chekbloc = True
                    Set rsBinLokasi = New ADODB.Recordset
                    OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION WHERE (Location_Code = '" & Node.Tag & "')"
                    If rsBinLokasi.Recordcount > 0 Then
                        Set DataTrans.Recordset = rsBinLokasi
                        Set GridTrans.DataSource = rsBinLokasi
                        GridTrans.Refresh
                        GridTrans.Columns(4).Visible = False
                        LockObject TxtLoc, TxtLoc.Count, True
                            'RcGroup.DBOpen "SELECT     NoAccount AS [Id Group], AccountName AS [Aktiva Group] FROM         GlAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'Aktiva Tetap Kantor' OR                      Type = N'Aktiva Tetap Produksi' OR                      Type = N'Aktiva Tetap Tak Berwujud') ORDER BY NoAccount", CNN, lckLockReadOnly
    '                        OpenTable rsFS, cnn, "SELECT WHSE_BINTYPE.Location_Code, WHSE_BINTYPE.Description AS Bin_Type" & _
                                                    " FROM WHSE_BINTYPE INNER JOIN" & _
                                                    " WHSE_BIN ON WHSE_BINTYPE.Code = WHSE_BIN.Bin_Type_Code AND" & _
                                                    " WHSE_BINTYPE.Location_Code = WHSE_BIN.Location_Code" & _
                                                    " GROUP BY WHSE_BINTYPE.Location_Code, WHSE_BINTYPE.Description"
                            OpenTable rsFS, CNN, "SELECT Code, Location_Code, Description, Receive, Ship, Put_Away, Pick, bin_prefik" & _
                                                    " From dbo.WHSE_BINTYPE WHERE (Location_Code = '" & Node.Tag & "')"
                            Set DCFS(0).RowSource = rsFS.DataSource
                        FrLokasi(2).Move 3225, 240, FrLokasi(2).width, FrLokasi(2).Height
                        FrLokasi(2).Visible = True
                        FrLokasi(2).ZOrder (0)
                        For I = 0 To TxtLoc.Count - 1
                            Set TxtLoc(I).DataSource = rsBinLokasi
                        Next
                     Else
                        Set DataTrans.Recordset = rsBinLokasi
                        Set GridTrans.DataSource = rsBinLokasi
                        GridTrans.Refresh
                        GridTrans.Columns(4).Visible = False
                        LockObject TxtLoc, TxtLoc.Count, True
                        FrLokasi(2).Move 3225, 240, FrLokasi(2).width, FrLokasi(2).Height
                        FrLokasi(2).Visible = True
                        FrLokasi(2).ZOrder (0)
                        For I = 0 To TxtLoc.Count - 1
                            TxtLoc(I).Text = ""
                        Next
                     End If
                     
                Case "BINCONT"
                    chekwh = False
                    chekbtype = False
                    chekbloc = False
                    chekeditbloc = False
                    chekeditbcont = True
                    Set rsBinContent = New ADODB.Recordset
                    OpenTable rsBinContent, CNN, "Select * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & Node.Tag & "')"
'                    Debug.Print rsBinContent.Source
                    Set DataTrans.Recordset = rsBinContent
                    Set GridTrans.DataSource = rsBinContent
                    GridTrans.Refresh
                    GridTrans.Columns(11).Visible = False
                            DCFS(1).BoundColumn = "": DCFS(1).DataField = "": DCFS(1).ListField = ""
                            'OpenTable rsFS2, cnn, "SELECT     Code, Description From WHSE_BINTYPE WHERE     (Location_Code = '" & Node.Tag & "')"
                            OpenTable rsFS2, CNN, "SELECT Code, Location_Code, Description, Receive, Ship, Put_Away, Pick, bin_prefik" & _
                                                    " From dbo.WHSE_BINTYPE WHERE (Location_Code = '" & Node.Tag & "')"
                            DCFS(1).BoundColumn = "Code": DCFS(1).DataField = "Code": DCFS(1).ListField = "Description"
                            Set DCFS(1).RowSource = rsFS2.DataSource
                            'lblLabels(25).Move 270, 1410, lblLabels(25).Width, lblLabels(25).Height: lblLabels(25).Visible = True: lblLabels(25).ZOrder 0
                            DCFS(1).Move 240, 1410, DCFS(1).width, DCFS(1).Height: DCFS(1).Visible = True: DCFS(1).ZOrder 0
                            
                                    DCFS(2).BoundColumn = "": DCFS(2).DataField = "": DCFS(2).ListField = ""
                                    'OpenTable rsFS3, cnn, "SELECT * From WHSE_BINTYPE WHERE (Location_Code = '" & Node.Tag & "')"
                                    OpenTable rsFS3, CNN, "SELECT Code, Description, Bin_Type, Bin_Ranking, Max_Weight" & _
                                                            " From V_BINLOCATION WHERE (Location_Code = '" & Node.Tag & "')"
                                    Set DCFS(2).RowSource = rsFS3.DataSource
                                    DCFS(2).BoundColumn = "Code": DCFS(2).DataField = "Code": DCFS(2).ListField = "Description"
                                    'lblLabels(26).Move 2385, 1410, lblLabels(26).Width, lblLabels(26).Height: lblLabels(26).Visible = True: lblLabels(26).ZOrder 0
                                    DCFS(2).Move 2355, 1410, DCFS(2).width, DCFS(2).Height: DCFS(2).Visible = True: DCFS(2).ZOrder 0

                    FrLokasi(3).Move 3225, 240, FrLokasi(3).width, FrLokasi(3).Height
                    FrLokasi(3).Visible = True
                    FrLokasi(3).ZOrder (0)
                    LockObject TxtContent, TxtContent.Count, True
                    For I = 0 To TxtContent.Count - 1
                        Set TxtContent(I).DataSource = rsBinContent
                    Next
                   
                    MyDDE.SetPermissions = UserAddnewDenied
                Case "BINCONTENTRY"
                    'BUKA MASTER BIN
                    myClass.Gelas True
                  
                    Set rsBIN = New ADODB.Recordset
                    strSQL = "SELECT     WHSE_BIN.Code, WHSE_BINTYPE.Description, WHSE_BIN.Location_Code, WHSE_BIN.Bin_Type_Code, WHSE_BIN.Bin_Ranking, " & _
                    " WHSE_BIN.Max_Weight FROM WHSE_BIN INNER JOIN WHSE_BINTYPE ON WHSE_BIN.Bin_Type_Code = WHSE_BINTYPE.Code " & _
                    " WHERE (WHSE_BIN.Location_Code = '" & Node.Tag & "') ORDER BY WHSE_BINTYPE.Description, WHSE_BIN.Code"
                    
'                    Debug.Print strSQL
                    rsBIN.CursorLocation = adUseClient
                    rsBIN.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
                    With rsBIN
                        LView.ListItems.Clear
                        Do While Not .EOF
                            With LView.ListItems.Add(, , .Fields("Code").Value)
                               .SubItems(1) = rsBIN.Fields("Description").Value
                               .SubItems(2) = rsBIN.Fields("Location_Code").Value
                            End With
                            .MoveNext
                        Loop
                    End With
                    PicBin.Visible = True
                    PicBin.ZOrder (0)
                
                    OptSearch_Click (1)
                    If rsBIN.Recordcount <> 0 Then
                        LView.ListItems(1).Selected = True
                        LView_ItemClick LView.ListItems(1)
                        LView.SetFocus
                        For I = 0 To 3
                            CmdPanah(I).Enabled = True
                        Next
                    Else
                        LBLKateg(3).Caption = ""
                        LBLKateg(4).Caption = ""
                        Set GridWizard(0).DataSource = rsBIN
                        Set GridWizard(1).DataSource = rsBIN
                        DataProses1.Caption = ""
                        DataProses2.Caption = ""
                        For I = 0 To 3
                            CmdPanah(I).Enabled = False
                        Next
                    End If
                    myClass.Gelas False
                    MyDDE.SetPermissions = UserEditAddnewDenied
            End Select
        End If
End Select
Set SelectNode = Node
GridLayout
GridTrans.Refresh
Exit Sub
1:
MessageBox Err.Description, "formwh:tviewmenu_nodeclick " & Err.Number, msgOkOnly, msgExclamation
End Sub

'Sub SetMenuFile(VFind As Boolean, VRefresh As Boolean, VPrint As Boolean, VExit As Boolean)
'    MFile(1).Enabled = VRefresh And Mid(vRole, 6, 1) = "1"
'    MFile(3).Enabled = VPrint And Mid(vRole, 7, 1) = "1"
'    MFile(5).Enabled = VExit And Mid(vRole, 8, 1) = "1"
'
'    Toolbar1.Buttons(6).Enabled = VRefresh And Mid(vRole, 6, 1) = "1"
'    Toolbar1.Buttons(8).Enabled = VPrint And Mid(vRole, 7, 1) = "1"
'    Toolbar1.Buttons(11).Enabled = VExit And Mid(vRole, 8, 1) = "1"
'End Sub
'
'Sub SetMenuEdit(VAdd, VEdit, VDelete, VSave)
'
'    MEdit(0).Enabled = VAdd And Mid(vRole, 1, 1) = "1"
'    MEdit(1).Enabled = VEdit And Mid(vRole, 2, 1) = "1"
'    MEdit(2).Enabled = VDelete And Mid(vRole, 3, 1) = "1"
'
'    If IsMissing(VHist) Then
'       MEdit(7).Enabled = False
'       Toolbar1.Buttons(9).Enabled = False
'    Else
'       MEdit(7).Enabled = VHist
'       Toolbar1.Buttons(9).Enabled = True
'    End If
'
'    If IsMissing(VNotAppr) Then
'       MEdit(9).Enabled = False
'    Else
'       MEdit(9).Enabled = VNotAppr And Mid(vRole, 5, 1) = "1"
'    End If
'
'    Toolbar1.Buttons(1).Enabled = VAdd And Mid(vRole, 1, 1) = "1"
'    Toolbar1.Buttons(2).Enabled = VEdit And Mid(vRole, 2, 1) = "1"
'    Toolbar1.Buttons(3).Enabled = VDelete And Mid(vRole, 3, 1) = "1"
'End Sub

Private Sub FilterProc(sKateg As String, sWH As String)
On Error GoTo 9
Dim rsBIN As New ADODB.Recordset
Dim rsBinLoc As New ADODB.Recordset

'FILTER GRID SEBELAH KIRI
'SEMUA DATA KECUALI BARANG TANPA LOKASI
'BASED ON SELECTED BIN LOKASI ON LIST VIEW



' dbo.WHSE_BINTYPE.location_code, dbo.WHSE_BINCONTENT.Bin_Code, dbo.WHSE_BINTYPE.Description, dbo.WHSE_BINCONTENT.NoItem,
'dbo.Inventory.ItemName, dbo.WHSE_BINCONTENT.UOM, dbo.WHSE_BINCONTENT.Min_Qty, dbo.WHSE_BINCONTENT.Max_Qty,
'dbo.WHSE_BINCONTENT.Qty_per_UOM , dbo.WHSE_BINCONTENT.ROP, dbo.WHSE_BINCONTENT.SafetyStock





'strSQL = "SELECT P_PRODUCT.KODE, P_PRODUCT.NAMA, P_PRODUCT.KODE_SATUAN_KECIL AS UOM FROM WHSE_BINCONTENT RIGHT OUTER JOIN " & _
'" P_PRODUCT ON WHSE_BINCONTENT.Item_No = P_PRODUCT.KODE " & _
'" WHERE (P_PRODUCT.TIDAK_TERPAKAI = '1') AND (WHSE_BINCONTENT.Item_No IS NULL) ORDER BY P_PRODUCT.KODE"


strSQL = "SELECT Inventory.noItem, Inventory.ItemName, Inventory.UOM AS UOM FROM WHSE_BINCONTENT RIGHT OUTER JOIN " & _
" Inventory ON WHSE_BINCONTENT.NoItem = Inventory.noItem " & _
" WHERE (WHSE_BINCONTENT.NoItem IS NULL) ORDER BY Inventory.noitem"



rsBIN.CursorLocation = adUseClient
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockReadOnly, adCmdText
Set GridWizard(0).DataSource = rsBIN
Set DataProses1.Recordset = rsBIN
'GridWizard(0).ReBind
DataProses1.Caption = "Record  " & Format(DataProses1.Recordset.AbsolutePosition, RecForm) & " of " & Format(DataProses1.Recordset.Recordcount, RecForm)

'FILTER GRID SEBELAH KANAN
'strSQL = " SELECT WHSE_BINCONTENT.Item_No AS KODE, P_PRODUCT.NAMA, WHSE_BINCONTENT.Location_Code,WHSE_BINCONTENT.UOM " & _
'        " FROM WHSE_BINCONTENT INNER JOIN P_PRODUCT ON WHSE_BINCONTENT.Item_No = P_PRODUCT.KODE " & _
'        " WHERE (WHSE_BINCONTENT.Bin_Code ='" & sKateg & "') AND (P_PRODUCT.TIDAK_TERPAKAI = '1') " & _
'        " AND (WHSE_BINCONTENT.Location_Code = '" & sWH & "')"
        
        
strSQL = " SELECT WHSE_BINCONTENT.NOItem AS KODE, inventory.itemNAMe, WHSE_BINCONTENT.Location_Code,WHSE_BINCONTENT.UOM " & _
        " FROM WHSE_BINCONTENT INNER JOIN Inventory ON WHSE_BINCONTENT.Noitem = inventory.noitem " & _
        " WHERE (WHSE_BINCONTENT.Bin_Code ='" & sKateg & "') " & _
        " AND (WHSE_BINCONTENT.Location_Code = '" & sWH & "')"
        
Set rsBIN = New ADODB.Recordset
rsBIN.CursorLocation = adUseClient
'Debug.Print strSQL
rsBIN.Open strSQL, CNN, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set GridWizard(1).DataSource = rsBIN
Set DataProses2.Recordset = rsBIN
'GridWizard(1).ReBind
Set rsBIN = Nothing
DataProses2.Caption = "Record  " & Format(DataProses2.Recordset.AbsolutePosition, RecForm) & " of " & Format(DataProses2.Recordset.Recordcount, RecForm)


strSQL = "Select * From V_BINLOCATION WHERE (Code ='" & sKateg & "') AND (Location_Code = '" & sWH & "')"
Set rsBinLoc = New ADODB.Recordset
rsBinLoc.CursorLocation = adUseClient
rsBinLoc.Open strSQL, CNN, adOpenKeyset, adLockBatchOptimistic, adCmdText
BinLocation = IIf(Not IsNull(rsBinLoc.Fields("Description")), rsBinLoc.Fields("Description"), "")
rsBinLoc.Close
Set rsBinLoc = Nothing
Exit Sub
9:
MessageBox Err.Description, "formwh:filterproc " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub TxtCarik_Change()
On Error GoTo 1
If Len(TxtCarik.Text) <> 0 Then
   DataProses1.Recordset.Filter = strFilter & " like '" & TxtCarik.Text & "*'"
Else
   CmdFresh(0).Value = True
End If
Err.Clear
1:
MessageBox Err.Description, "formwh:txtcarik_change " & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub CmdFresh_Click(Index As Integer)
On Error GoTo 1
Select Case Index
    Case 0
        DataProses1.Recordset.Filter = adFilterNone
        TxtCarik.Text = ""
        DataProses1.Recordset.Requery
        TxtCarik.SetFocus
    Case Else
        If Len(TxtCarik.Text) <> 0 Then DataProses1.Recordset.Filter = strFilter & " like '" & TxtCarik.Text & "%'"
        TxtCarik.SetFocus
End Select
Exit Sub
1:
MessageBox Err.Description, "formwh:cmdfresh_click " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub CmdPanah_Click(Index As Integer)
'On Error GoTo PanahErr
On Error Resume Next
Dim sKode, sUOM As String

Screen.MousePointer = vbHourglass
Select Case Index
    Case 0
'        Debug.Print GridWizard(0).Columns(0).Value
        If DataProses1.Recordset.EOF Then GoTo Keluar
       ' strSQL = " Update P_PRODUCT SET LOKASI_NEW = '" & LView.SelectedItem.Text & "' WHERE (KODE = '" & GridWizard(0).Columns(0).Value & "')"
        sKode = GridWizard(0).Columns(0).Text
        If Len(GridWizard(0).Columns(2).Text) = 0 Then
            MessageBox "Kode Barang ' " & sKode & " 'Satuan kosong ", vbInformation, "Entry Satuan"
            sUOM = UCase(InputBox("Satuan : ", "Entry Satuan"))
        Else
            sUOM = GridWizard(0).Columns(2).Text
        End If
      
        strSQL = "INSERT INTO WHSE_BINCONTENT (NoItem, Location_Code, Bin_Code, UOM)" & _
        " Values ('" & sKode & "','" & LView.SelectedItem.SubItems(2) & _
        "','" & LView.SelectedItem.Text & "','" & IIf(IsNull(sUOM), "PCS", sUOM) & "')"
         
        myClass.SendKoman strSQL
    Case 1
        If DataProses1.Recordset.EOF Then GoTo Keluar
       ' If MessageBox("Pilih BIN Lokasi  " & LView.SelectedItem.Text & "  untuk semua Barang " & UCase(TxtCarik.Text) & " ?", vbYesNo + vbQuestion, "Transfer Data") = vbYes Then
        If MessageBox("Pilih BIN Lokasi  '" & LView.SelectedItem.Text & "'  untuk semua Barang " & UCase(TxtCarik.Text) & " ?", "Transfer Data", msgYesNo, msgQuestion) = 1 Then
            Do While Not DataProses1.Recordset.EOF
                sKode = GridWizard(0).Columns(0).Text
                If Len(GridWizard(0).Columns(2).Text) = 0 Then
                    MessageBox "Kode Barang ' " & sKode & " ' Satuan kosong ", vbExclamation, "Entry Satuan"
                    sUOM = UCase(InputBox("Satuan : ", "Entry Satuan"))
                Else
                    sUOM = GridWizard(0).Columns(2).Text
                End If
                strSQL = "INSERT INTO WHSE_BINCONTENT (NoItem, Location_Code, Bin_Code, UOM)" & _
                " Values ('" & sKode & "','" & LView.SelectedItem.SubItems(2) & _
                "','" & LView.SelectedItem.Text & "','" & sUOM & "')"
                
                myClass.SendKoman strSQL
                DataProses1.Recordset.MoveNext
            Loop
        End If
    Case 2
        If DataProses2.Recordset.EOF Then GoTo Keluar
        strSQL = "DELETE FROM WHSE_BINCONTENT WHERE (NoItem = '" & GridWizard(1).Columns(0).Value & "') " & _
        " AND (Bin_Code = '" & LView.SelectedItem.Text & "') AND (Location_Code = '" & LView.SelectedItem.SubItems(2) & "')"
        myClass.SendKoman strSQL
    Case 3
        If DataProses2.Recordset.EOF Then GoTo Keluar
        If MessageBox("Hapus semua BIN Lokasi barang  ' " & LView.SelectedItem.Text & " ' ?", "Transfer Data", msgYesNo, msgQuestion) = 1 Then
           strSQL = "DELETE FROM WHSE_BINCONTENT WHERE (Bin_Code = '" & LView.SelectedItem.Text & "') AND (Location_Code = '" & LView.SelectedItem.SubItems(2) & "')"
           myClass.SendKoman strSQL
        End If
    Case Else: Exit Sub
End Select

DataProses1.Recordset.Requery
DataProses2.Recordset.Requery

Keluar:
   Screen.MousePointer = 0
   GridWizard(0).ReBind
   GridWizard(1).ReBind
   Exit Sub

'PanahErr:
'   Screen.MousePointer = 0
'   messagebox Err.Description, vbCritical
End Sub

Private Sub ProsesTipe(Index As TransData)
On Error GoTo Keluar
Dim bkMark As Variant
Dim iAsk As Integer

'If Tbar.Tag <> "ADD" Then bkMark = rsBinType.Bookmark
Select Case Index
    Case TransData.adNew
    Case TransData.adEdit
        bkMark = rsBinType.Bookmark
End Select
Select Case Index
    Case TransData.adNew    'NEW
        strSQL = "INSERT INTO WHSE_BINtype ( Code, Location_Code, Description, Receive, Ship, Put_Away, Pick, bin_prefik) " & _
                "VALUES ('" & TxtType(0) & "', '" & TxtType(1) & "', '" & TxtType(2) & "', " & FQty(TxtType(4)) & ", " & FQty(TxtType(5)) & ", " & FQty(TxtType(6)) & ", " & FQty(TxtType(7)) & ", '" & TxtType(3) & "' ) "
'        myClass.SendKoman strSQL
    Case TransData.adEdit     'EDIT
        strSQL = " Update WHSE_BINtype SET Location_Code = '" & TxtType(1) & "', Description=  '" & TxtType(2) & "', Receive = " & FQty(TxtType(4)) & ", Ship = " & FQty(TxtType(5)) & ", Put_Away = " & FQty(TxtType(6)) & ", Pick = " & FQty(TxtType(7)) & ", bin_prefik='" & TxtType(3) & "' " & _
        " WHERE (Code = '" & TxtType(0) & "')"
'        myClass.SendKoman strSQL
    Case TransData.adDelete     'DELETE
'        strSQL = "DELETE FROM WHSE_BINCONTENT WHERE (Item_No = '" & GridWizard(1).Columns(0).Value & "') " & _
        " AND (Bin_Code = '" & LView.SelectedItem.Text & "') AND (Location_Code = '" & LView.SelectedItem.SubItems(2) & "')"
        
        strSQL = "DELETE FROM WHSE_BINtype WHERE (Code = '" & TxtType(0) & "') "
        
        iAsk = MessageBox("' " & TxtType(2).Text & " ' akan dihapus..?", vbQuestion + vbYesNo, "Konfirmasi")
        If iAsk = vbYes Then
'            myClass.SendKoman strSQL
        Else
            GoSub Mencolot
        End If
    Case TransData.adCancel
        GoSub Mencolot
    Case Else
        Exit Sub
End Select

If myClass.SendCommandToServer(strSQL) Then
    If Left(SelectNode.Key, 1) = "C" Then
        sNodeKey = SelectNode.Parent.Key
        TimerON "Saving " & SelectNode.Parent.Text & " " & SelectNode.Text & " Successfully..."
    Else
        sNodeKey = SelectNode.Key
        TimerON "Saving " & SelectNode.Parent.Text & " " & SelectNode.Text & " Successfully..."
    End If
Else
    MessageBox "Saving data error..", vbCritical, App.ProductName
    Exit Sub
End If

Mencolot:
Set rsBinType = New ADODB.Recordset
OpenTable rsBinType, CNN, "Select * From WHSE_BINtype WHERE (Location_Code = '" & TxtType(1).Text & "')"
Set DataTrans.Recordset = rsBinType

For I = 0 To TxtType.Count - 1
    Set TxtType(I).DataSource = rsBinType
Next
GridLayout
Select Case Index
    Case TransData.adNew
    Case TransData.adEdit
        rsBinType.Bookmark = bkMark
    Case TransData.adDelete
        If iAsk = vbYes Then If rsBinType.Recordcount <> 0 Then rsBinType.MoveLast
End Select

GridTrans.SetFocus
Exit Sub
Keluar:
    MessageBox Err.Description, vbCritical, App.ProductName
End Sub
'LOKASI WAREHOUSE
Private Sub ProsesWHouse(Index As TransData)
On Error GoTo Keluar

'Dim TipeKode As String
'Dim bkMark As Variant

'Screen.MousePointer = vbHourglass
bkMark = rsWHouse.Bookmark
Select Case Index
    Case TransData.adEdit
'        TipeKode = BinTipeKode(TxtLoc(3))
        strSQL = " Update WHSE_LOCATION SET Address = '" & TxtWH(2) & "', City=  '" & TxtWH(3) & "', Phone = '" & TxtWH(4).Text & "', Contact = '" & TxtWH(5).Text & "' WHERE (Code = '" & TxtWH(0).Text & "')"
        myClass.SendKoman strSQL
    
    Case TransData.adCancel
        GoSub Mencolot
    Case Else: Exit Sub
End Select

Mencolot:
Set rsWHouse = New ADODB.Recordset
OpenTable rsWHouse, CNN, "Select * From V_WAREHOUSE Order By CODE"
Set DataTrans.Recordset = rsWHouse
For I = 0 To TxtWH.Count - 1
    Set TxtWH(I).DataSource = rsWHouse
Next
GridLayout

Select Case Index
    Case TransData.adEdit
        rsWHouse.Bookmark = bkMark
    Case Else
End Select

Exit Sub

Keluar:

End Sub
'BIN LOKASI
Private Sub ProsesLokasi(Index As TransData)
On Error GoTo Keluar

Dim TipeKode As String

'Screen.MousePointer = vbHourglass
'bkMark = rsBinLokasi.Bookmark
Select Case Index
    Case TransData.adNew
    Case TransData.adEdit
        bkMark = rsBinLokasi.Bookmark
End Select
Select Case Index
    Case TransData.adNew
        strSQL = "INSERT INTO WHSE_BIN ( Code, Location_Code, Description, Bin_Type_Code, Bin_Ranking, Max_Weight) " & _
                "VALUES ('" & TxtLoc(1).Text & "', '" & TxtLoc(0).Text & "', '" & TxtLoc(2).Text & "', '" & TxtLoc(6).Text & "', " & FQty(TxtLoc(4).Text) & ", " & FQty(TxtLoc(5).Text) & ")"
'        Debug.Print strSQL
        myClass.SendKoman strSQL
    
    Case TransData.adEdit
        TipeKode = BinTipeKode(TxtLoc(3))
        strSQL = " Update WHSE_BIN SET Location_Code = '" & TxtLoc(0) & "', Description=  '" & TxtLoc(2) & "', Bin_Type_Code = '" & TipeKode & "', Bin_Ranking = " & FQty(TxtLoc(4)) & ", Max_Weight = " & FQty(TxtLoc(5)) & " WHERE (Code = '" & TxtLoc(1) & "')"
'        Debug.Print strSQL
        myClass.SendKoman strSQL
    
    Case TransData.adDelete
        
        strSQL = "DELETE FROM WHSE_BIN WHERE (Code = '" & TxtLoc(1) & "') "
        If MessageBox("Bin Lokasi : ' " & TxtLoc(1).Text & " - " & TxtLoc(2).Text & " ' akan dihapus..?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            myClass.SendKoman strSQL
        End If
    Case TransData.adCancel
        GoSub Mencolot
    Case Else: Exit Sub
End Select

Mencolot:
Set rsBinLokasi = New ADODB.Recordset
OpenTable rsBinLokasi, CNN, "Select * From V_BINLOCATION WHERE (Location_Code = '" & SelectNode.Tag & "')"
Set DataTrans.Recordset = rsBinLokasi
For I = 0 To TxtLoc.Count - 1
    Set TxtLoc(I).DataSource = rsBinLokasi
Next
GridLayout

Select Case Index
    Case TransData.adNew
    Case TransData.adEdit
        rsBinLokasi.Bookmark = bkMark
    Case TransData.adDelete
End Select

Exit Sub

Keluar:

End Sub

Private Sub ProsesISI(Index As TransData)
On Error GoTo 19
Dim strSQL As String
On Error GoTo Keluar

'Screen.MousePointer = vbHourglass
'bkMark = rsBinLokasi.Bookmark
Select Case Index
    Case TransData.adNew
    Case TransData.adEdit
        bkMark = rsBinContent.Bookmark
End Select
Select Case Index
    Case TransData.adNew
        
        'paketbarang
        strSQL = "INSERT INTO WHSE_BINCONTENT (Location_Code, Bin_Code, Item_No, UOM, Min_QTY, Max_QTY, ROP, SafetyStock) " & _
                "VALUES ('" & TxtContent(0) & "', '" & TxtContent(1) & "', '" & TxtContent(3) & "', '" & TxtContent(5) & "', " & FQty(TxtContent(6)) & ", " & FQty(TxtContent(7)) & ", " & FQty(TxtContent(8)) & ", " & FQty(TxtContent(9)) & ")"
'        Debug.Print strSQL
        myClass.SendKoman strSQL
        
    Case TransData.adEdit
        'SendDataToServer ("UPDATE [Planned Order] Set [Convert] = " & BoolToInt(RcPlan.DBRecordset.Fields("Convert")) WHERE (ID = '" & RcPlan.DBRecordset.Fields("ID") & "')")
        
        strSQL = " Update WHSE_BINCONTENT SET Location_Code = '" & TxtContent(0) & "', Bin_Code = '" & TxtContent(1) & "', Item_No = '" & TxtContent(3) & "', UOM = '" & TxtContent(5) & "', Min_QTY = " & FQty(TxtContent(6)) & ", Max_QTY = " & FQty(TxtContent(7)) & ", ROP = " & FQty(TxtContent(8)) & ", SafetyStock = " & FQty(TxtContent(9)) & " WHERE (IDX = '" & TakeKode(SelectNode.Tag) & "')"
'        Debug.Print strSQL
        myClass.SendKoman strSQL
    
'    Case 2
'
''        strSQL = "DELETE FROM WHSE_BINCONTENT WHERE (Item_No = '" & GridWizard(1).Columns(0).Value & "') " & _
'        " AND (Bin_Code = '" & LView.SelectedItem.Text & "') AND (Location_Code = '" & LView.SelectedItem.SubItems(2) & "')"
'
'        strSQL = "DELETE FROM WHSE_BINCONTENT WHERE (IDX = '" & TakeKode(SelectNode.Tag) & "')"
''        messagebox Err.Number & vbCrLf & Err.Description, vbInformation, "Message"
'        If messagebox("Yakin akan dihapus..?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
'            myClass.SendKoman strSQL
'        End If
    
    Case Else: Exit Sub
End Select

'TombolSiap False
Set rsBinContent = New ADODB.Recordset
OpenTable rsBinContent, CNN, "Select * From V_BINCONTENT_STOCK WHERE (Location_Code = '" & TxtContent(0).Text & "')"
'                    Debug.Print rsBinContent.Source
Set DataTrans.Recordset = rsBinContent
For I = 0 To TxtContent.Count - 1
    Set TxtContent(I).DataSource = rsBinContent
Next
GridLayout
Select Case Index
    Case TransData.adNew
    Case TransData.adEdit
        rsBinContent.Bookmark = bkMark
    Case TransData.adDelete
End Select
GridTrans.SetFocus
Exit Sub


Keluar:
Exit Sub
19:
MessageBox Err.Description, "formwh:prosesisi " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub LockClearObject(TxtObj As Object, TxtCount As Long, vStart As Integer, bStatus As Boolean)
On Error GoTo 12
Dim I As Long
For I = vStart To TxtCount - 1
    TxtObj(I).Locked = bStatus
    If bStatus = True Then
        TxtObj(I).BackColor = &HC0FFFF
    Else
        TxtObj(I).BackColor = &HFFFFFF
        TxtObj(I).Text = ""
    End If
Next
Exit Sub
12:
MessageBox Err.Description, "formwh:lockclearobject" & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub LockObject(TxtObj As Object, TxtCount As Long, bStatus As Boolean)
On Error GoTo 13
Dim I As Long
For I = 0 To TxtCount - 1
    TxtObj(I).Locked = bStatus
    If bStatus = True Then
        TxtObj(I).BackColor = &HC0FFFF
    Else
        TxtObj(I).BackColor = &HFFFFFF
    End If
Next
Exit Sub
13:
MessageBox Err.Description, "formwh:lockobject " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub TimerON(sMsg As String)
'FormSession.sMessage = sMsg
'FormSession.Show vbModal
End Sub

Function TakeKode(inKODE As String) As String
Dim Index As String
Dim dataTMP As ADODB.Recordset
'OpenTable dataTMP, cnn, "Select * From V_TMP WHERE (Location_Code = '" & Node.Tag & "')"
OpenTable dataTMP, CNN, "Select * From V_TMP WHERE (Location_Code = '" & inKODE & "')"

dataTMP.AbsolutePosition = GridTrans.row + 1

Index = dataTMP.Fields("IDX")
TakeKode = Index
End Function

Private Sub TombolSiap(status As Boolean)
'    CmdLok.Visible = status
'    CmdBrg.Visible = status
End Sub

'Private Sub TampilBinType()
'Dim rcSET As ADODB.Recordset
'
''OpenTable rcSET, cnn, "Select Code,Location_Code,Description From WHSE_BINTYPE Order By Code"
'OpenTable rcSET, cnn, "Select * From WHSE_BINTYPE Order By Code"
'Set DataGrid1.DataSource = rcSET
'FrmShow.Tag = "WHSE_BINTYPE"
'FrmShow.Move 8310, 465, FrmShow.Width, FrmShow.Height
'FrmShow.Visible = True
'FrmShow.ZOrder 0
'
'End Sub

Function BinTipeKode(inKODE As String)
Dim rcSet As ADODB.Recordset
On Error GoTo 3
OpenTable rcSet, CNN, "Select Code, Description From WHSE_BINTYPE Where Description = '" & inKODE & "' Order By Code"
BinTipeKode = IIf(Not IsNull(rcSet.Fields(0).Value), rcSet.Fields(0).Value, "BT")
Exit Function
3:
MessageBox Err.Description, "formwh:bintipekode" & Err.Number, msgOkOnly, msgExclamation
End Function
Private Sub InitBinLokasi(ByVal strNode As String)
TxtLoc(0).Text = strNode: TxtLoc(0).Locked = True: TxtLoc(0).BackColor = &HC0FFFF
'TxtLoc(1).Text = strNode: TxtLoc(1).Locked = True
TxtLoc(4).Text = 0
TxtLoc(5).Text = 0
TxtLoc(3).SetFocus
End Sub
Private Sub Inisialisasi(ByVal strNode As String)
TxtType(0).Text = CounterKode: TxtType(0).Locked = False
TxtType(1).Text = strNode: TxtType(1).Locked = True
TxtType(4).Text = 0
TxtType(5).Text = 0
TxtType(6).Text = 0
TxtType(7).Text = 0
TxtType(2).SetFocus
End Sub

Function CounterKode() As String
On Error GoTo 6
Dim Kode As String
Dim rcSet As ADODB.Recordset
    OpenTable rcSet, CNN, "Select max(cast(substring(code,3,10)as int)) From WHSE_BINTYPE"
    If rcSet.Recordcount <> 0 Then
        If Not IsNull(rcSet.Fields(0)) Then
           ' Dim Teks: Teks = Right(rcSet.Fields(0), Len(rcSet.Fields(0)) - 2)
           ' Kode = "BT" & Val(Right(rcSet.Fields(0), Len(rcSet.Fields(0)) - 2)) + 1
            Kode = "BT" & Val(rcSet.Fields(0)) + 1
        Else
            Kode = "BT1"
        End If
    Else
        Kode = "BT1"
    End If
    CounterKode = Kode
Exit Function
6:
MessageBox Err.Description, "frmwh:counterkode" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function AutoIndexAcc() As String
On Error GoTo 2
Dim Rckode As ADODB.Recordset
Dim mVarTotalDigit
'OpenTable Rckode, CNN, "SELECT MAX(Code) AS MaxKode FROM WHSE_BIN Where Code like '" & TxtLoc(7).Text & "[^1]%'"
OpenTable Rckode, CNN, "SELECT MAX(Code) AS MaxKode FROM WHSE_BIN Where Code like '" & TxtLoc(7).Text & "%'"
With Rckode
     If .Recordcount <> 0 Then
        If Not IsNull(.Fields(0)) Then
            mVarTotalDigit = Val(Right(Rckode.Fields(0).Value, Len(Rckode.Fields(0).Value) - (Len(TxtLoc(7).Text) + 1))) + 1
        Else
           ' mVarTotalDigit = "001"
             mVarTotalDigit = "1"
        End If
     Else
        'mVarTotalDigit = "001"
         mVarTotalDigit = "1"
     End If
End With
'AutoIndexAcc = Left(mVarGroupAccount, 6) & mVarTotalDigit & KirimNull(2)
'AutoIndexAcc = TxtLoc(7) & "-" & mVarTotalDigit
AutoIndexAcc = TxtLoc(7) & "-" & "00" & mVarTotalDigit
Exit Function
2:
MessageBox Err.Description, "formwh:autoindexacc" & Err.Number, msgOkOnly, msgExclamation
End Function

'Function CounterLok(inKODE As String) As String
''On Error GoTo Keluar
'Dim KodeLok As String
'Dim BANYAK As Integer
'Dim rcSET As ADODB.Recordset
'                        'Select * From WHSE_BIN Where Code like '03[^1]%'
'    BANYAK = POS(inKODE)
''    OpenTable rcSET, cnn, "Select * From WHSE_BIN Where Code like  '" & Right(inKODE, Len(inKODE) - Len("BT")) & "[^1]%' Order By Code"
'
'    OpenTable rcSET, cnn, "Select max((code) * From WHSE_BIN Where Code like  '" & grid & " Order By Code"
'
''    If rcSET.RecordCount = 0 Then KodeLok = Right(inKODE, Len(inKODE) - Len("BT")) & "001": Exit Function
'    If rcSET.RecordCount = 0 Then KodeLok = Right(DataGrid1.Columns(2).Value, Len(DataGrid1.Columns(2).Value) - Len("RAK ")) & "-" & "001": Exit Function
'    rcSET.MoveLast
'
''    KodeLok = JadiAkhir(rcSET.Fields(1), Len(Right(inKODE, Len(inKODE) - Len("BT"))) + 1)
'    BANYAK = POS(rcSET.Fields(1))
'    KodeLok = JadiAkhir(rcSET.Fields(1), Len(rcSET.Fields(1)) - BANYAK)
'
'    CounterLok = KodeLok
'    Exit Function
'': Keluar
'
'End Function

Private Function JadiAkhir(inKODE As String, jmlHEADER As Integer) As String
'    JadiAkhir = Left(inKODE, jmlHEADER) + AkhirKode(Mid(inKODE, jmlHEADER + 1, 100))
    JadiAkhir = Left(inKODE, jmlHEADER) + AkhirKode(Right(inKODE, Len(inKODE) - jmlHEADER))
End Function

Private Function AkhirKode(inKODE As String) As String
On Error GoTo 1
Dim I As Long
    AkhirKode = Val(inKODE) + 1
    For I = 1 To Len(inKODE)
        If Len(AkhirKode) = I Then AkhirKode = Digit(Len(inKODE) - I) & AkhirKode
    Next I
Exit Function
1:
MessageBox Err.Description, "formwh:akhirkode" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function Digit(jml As Long) As String
Dim I As Long
For I = 1 To jml
    Digit = Digit + "0"
Next I
End Function

Private Function pos(Str As String) As Integer
On Error GoTo 18
Dim I As Integer
For I = 1 To Len(Str)
    If Mid(Str, I, 1) = "-" Then
        pos = I
        Exit Function
    Else
        pos = Len(Str)
    End If
Next I
Exit Function
18:
MessageBox Err.Description, "formwh:pos " & Err.Number, msgOkOnly, msgExclamation
End Function

Private Function TakeValue(Value As String) As String
    
End Function

Private Sub TxtContent_GotFocus(Index As Integer)
TxtContent(Index).SelStart = 0
TxtContent(Index).SelLength = Len(TxtContent(Index))
End Sub

Private Sub TxtType_GotFocus(Index As Integer)
TxtType(Index).SelStart = 0
TxtType(Index).SelLength = Len(TxtType(Index))
End Sub

Private Sub TxtWH_GotFocus(Index As Integer)
TxtWH(Index).SelStart = 0
TxtWH(Index).SelLength = Len(TxtWH(Index))
End Sub

Private Sub OpenDBWH()
On Error GoTo 17
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FormWH
    .BindFormTAG = "WH"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT WareHouse.WareHouse, WareHouse.[WareHouse Name], WareHouse.Locations, warehouse.kota, warehouse.telpon, warehouse.contact, WareHouse.NoAccount, WareHouse.GroupAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM  WareHouse INNER JOIN  GLAccount ON WareHouse.GroupAccount = GLAccount.NoAccount"
End With
Exit Sub
17:
MessageBox Err.Description, "formwh:opendbwh " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDBTYPE()
On Error GoTo 16
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FormWH
    Set .ActiveConnection = CNN
    '.PrepareQuery = "Select * From V_BINLOCATION"
    .PrepareQuery = "select code,location_code,description,receive,ship,put_away,pick,bin_prefik,timestamp from WHSE_BINtype"
End With
Exit Sub
16:
MessageBox Err.Description, "formwh:opendbtype " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDBINLOC()
On Error GoTo 15
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FormWH
    Set .ActiveConnection = CNN
    '.PrepareQuery = "Select * From WHSE_BINtype "
    .PrepareQuery = "Select * From V_BINLOCATION"
End With
Exit Sub
15:
MessageBox Err.Description, "formwh:opendbinloc" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDBINCONT()
On Error GoTo 14
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FormWH
    Set .ActiveConnection = CNN
    '.PrepareQuery = "Select * From WHSE_BINtype "
    .PrepareQuery = "Select * From V_BINCONTENT"
End With
Exit Sub
14:
MessageBox Err.Description, "formwh:open dbincount " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub PrepareQuery()
On Error GoTo 20
With MyDDE
    ' .PrepareAppend = " INSERT INTO WareHouse (WareHouse, [WareHouse Name], Locations) " & _
                      " VALUES ('" & TxtWH(0).Text & "', '" & TxtWH(1).Text & "', N'" & TxtWH(2).Text & "')"
                      
    ' .PrepareUpdate = " UPDATE WareHouse Set NoAccount='" & MyDDE.GetFieldByName("NoAccount") & "',GroupAccount ='" & MyDDE.GetFieldByName("Kode Kelompok Gudang") & "',[WareHouse Name] = N'" & ValidString(txtBox(1)) & "', Locations=N'" & ValidString(txtBox(2)) & "'  WHERE     (WareHouse = N'" & ValidString(txtBox(0)) & "')"
                     
     '.PrepareDelete = " DELETE FROM WareHouse WHERE   (WareHouse = N'" & txtBox(0) & "') "
End With
Exit Sub
20:
MessageBox Err.Description, "formwh:preparequery " & Err.Number, msgOkOnly, msgExclamation
End Sub


Private Sub next_WH()
On Error GoTo 13
Dim I As Integer
Set GridTrans.DataSource = MyDDE.ActiveRecordset
For I = 0 To TxtWH.Count - 1
     Set TxtWH(I).DataSource = MyDDE.ActiveRecordset
Next
Exit Sub
13:
MessageBox Err.Description, "formwh:nextwh " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub next_BinType()
On Error GoTo 12
Set GridTrans.DataSource = MyDDE.ActiveRecordset
For I = 0 To TxtType.Count - 1
    Set TxtType(I).DataSource = MyDDE.ActiveRecordset
Next
GridTrans.Columns(7).Visible = False
GridTrans.Columns(8).Visible = False
Exit Sub
12:
MessageBox Err.Description, "formwh:next_bintype " & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub next_BinLocation()
On Error GoTo 19

Set GridTrans.DataSource = MyDDE.ActiveRecordset
For I = 1 To TxtLoc.Count - 1
    Set TxtLoc(I).DataSource = MyDDE.ActiveRecordset
Next
GridTrans.Columns(4).Visible = False
Exit Sub
19:
MessageBox Err.Description, "formwh:lnext_binlocation " & Err.Number, msgOkOnly, msgExclamation
End Sub



Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1:
            RcPartner.DBOpen "SELECT GLAccount.NoAccount AS [Kode Kelompok Gudang], GLAccount.AccountName AS [Nama Kelompok Gudang] FROM         GLAccount INNER JOIN                       AccType ON GLAccount.Type = AccType.Tipe WHERE     (AccType.ID = 37) AND (GLAccount.[Group] = N'list Account')", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1:
            mCall.FromTagActive = "Nama Kelompok Gudang"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Kelompok Gudang Belum Ada.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Public Property Get MyAutoIndex() As String
       mVarIndexStr = AutoIndexWHAcc(MyDDE.GetFieldByName("Kode Kelompok Gudang"))
       MyAutoIndex = mVarIndexStr
End Property

Private Function AutoIndexWHAcc(ByVal GroupAcc As String) As String
Dim Rckode As New DBQuick
Dim mVarTotalDigit As Long
Rckode.DBOpen "SELECT     MAX(RIGHT(NoAccount, 2)) AS MaxNom FROM         GLAccount WHERE     (GroupAccount = N'" & GroupAcc & "') AND ([Group] = N'Detail List Account')", CNN, lckLockReadOnly
With Rckode.DBRecordset
     If .Recordcount <> 0 Then
        mVarTotalDigit = IIf(Not IsNull(.Fields(0)), .Fields(0), 10) + 1
     Else
        mVarTotalDigit = 10
     End If
End With
AutoIndexWHAcc = Left(GroupAcc, Len(GroupAcc) - 2) & mVarTotalDigit
End Function
