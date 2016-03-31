VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{63532264-0E3B-4975-8ED5-42900FAB471A}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmOtorisasi 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12390
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Tag             =   "Otorisasi User"
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
      Height          =   7860
      Left            =   45
      ScaleHeight     =   7830
      ScaleWidth      =   12225
      TabIndex        =   2
      Top             =   0
      Width           =   12255
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
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
         Height          =   7290
         Left            =   75
         ScaleHeight     =   7260
         ScaleWidth      =   12000
         TabIndex        =   3
         Top             =   90
         Width           =   12030
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   810
            Top             =   825
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   7159830
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmOtorisasi.frx":0000
                  Key             =   "Orang"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmOtorisasi.frx":0BD4
                  Key             =   "person1"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmOtorisasi.frx":14B0
                  Key             =   "person2"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmOtorisasi.frx":1D8C
                  Key             =   "TOP"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmOtorisasi.frx":2BE0
                  Key             =   "Dept"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   7050
            Left            =   60
            TabIndex        =   0
            Top             =   105
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   12435
            _Version        =   393217
            LabelEdit       =   1
            Style           =   3
            ImageList       =   "ImageList1"
            BorderStyle     =   1
            Appearance      =   1
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   7095
            Left            =   3525
            TabIndex        =   4
            Top             =   105
            Width           =   8430
            _ExtentX        =   14870
            _ExtentY        =   12515
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   15380335
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Akses Form"
            TabPicture(0)   =   "FrmOtorisasi.frx":36AC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Line1(2)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Line1(1)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Line1(0)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label1(0)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label1(1)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label1(2)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "GridFormList"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "TreeOto"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txtBox(0)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtBox(2)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Combo1"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "MSHFlexGrid1"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).ControlCount=   12
            TabCaption(1)   =   "Akses Laporan"
            TabPicture(1)   =   "FrmOtorisasi.frx":36C8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "DataGrid1(1)"
            Tab(1).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
               Height          =   5880
               Left            =   4185
               TabIndex        =   14
               Top             =   1110
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   10372
               _Version        =   393216
               BorderStyle     =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.ComboBox Combo1 
               DataField       =   "User Name"
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
               Left            =   3705
               TabIndex        =   7
               Tag             =   "ASM"
               Text            =   "Combo1"
               Top             =   345
               Width           =   2925
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
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
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   1215
               MaxLength       =   100
               PasswordChar    =   "*"
               TabIndex        =   6
               Tag             =   "ASM"
               Top             =   705
               Width           =   5415
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "User ID"
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
               Left            =   1215
               MaxLength       =   15
               TabIndex        =   5
               Tag             =   "ASM"
               Top             =   360
               Width           =   1440
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   3450
               Index           =   1
               Left            =   -74910
               TabIndex        =   8
               Top             =   390
               Width           =   7140
               _ExtentX        =   12594
               _ExtentY        =   6085
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   "Nama Report"
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
                  DataField       =   "Report"
                  Caption         =   "Laporan"
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
                     ColumnWidth     =   4875.024
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1560.189
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.TreeView TreeOto 
               Height          =   5895
               Left            =   75
               TabIndex        =   9
               Top             =   1110
               Width           =   4080
               _ExtentX        =   7197
               _ExtentY        =   10398
               _Version        =   393217
               LabelEdit       =   1
               Style           =   7
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               Appearance      =   0
            End
            Begin MSDataGridLib.DataGrid GridFormList 
               Height          =   3480
               Left            =   3285
               TabIndex        =   10
               Top             =   1425
               Visible         =   0   'False
               Width           =   3420
               _ExtentX        =   6033
               _ExtentY        =   6138
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BorderStyle     =   0
               Enabled         =   0   'False
               ColumnHeaders   =   -1  'True
               HeadLines       =   1
               RowHeight       =   15
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
               Caption         =   "Otorisasi Menu"
               ColumnCount     =   1
               BeginProperty Column00 
                  DataField       =   "Form List"
                  Caption         =   "Form List"
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
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     Button          =   -1  'True
                     ColumnWidth     =   3165.166
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
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
               Height          =   210
               Index           =   2
               Left            =   120
               TabIndex        =   13
               Top             =   765
               Width           =   765
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "User Name"
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
               Height          =   210
               Index           =   1
               Left            =   2760
               TabIndex        =   12
               Top             =   420
               Width           =   885
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "User ID"
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
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   420
               Width           =   600
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   120
               X2              =   1455
               Y1              =   675
               Y2              =   675
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   2715
               X2              =   3750
               Y1              =   675
               Y2              =   675
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   120
               X2              =   1455
               Y1              =   1020
               Y2              =   1020
            End
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmOtorisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
