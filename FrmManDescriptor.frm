VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmManDescriptor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacture Desc"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmManDescriptor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Tag             =   "Master Outsourced "
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   10095
      TabIndex        =   13
      Top             =   0
      Width           =   10095
      Begin TabDlg.SSTab SSTab1 
         Height          =   5460
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   9631
         _Version        =   393216
         Style           =   1
         Tab             =   1
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
         TabCaption(0)   =   "List Job-Stock"
         TabPicture(0)   =   "FrmManDescriptor.frx":6852
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Picture4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "FrmManDescriptor.frx":686E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Picture3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Additional"
         TabPicture(2)   =   "FrmManDescriptor.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture5"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            ForeColor       =   &H80000008&
            Height          =   4920
            Left            =   120
            ScaleHeight     =   4890
            ScaleWidth      =   9570
            TabIndex        =   26
            Top             =   390
            Width           =   9600
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "UOM"
               DataSource      =   "Adodc1"
               Height          =   330
               Index           =   2
               Left            =   1275
               MaxLength       =   15
               TabIndex        =   5
               Tag             =   "Partner"
               Top             =   765
               Width           =   3045
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Catatan"
               Height          =   630
               Index           =   5
               Left            =   1275
               MaxLength       =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Tag             =   "Partner"
               Top             =   1890
               Width           =   8190
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "UnitCost"
               DataSource      =   "Adodc1"
               Height          =   330
               Index           =   4
               Left            =   5580
               MaxLength       =   15
               TabIndex        =   10
               Tag             =   "Partner"
               Top             =   1140
               Width           =   1785
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "UnitPrice"
               DataSource      =   "Adodc1"
               Height          =   330
               Index           =   3
               Left            =   2535
               MaxLength       =   15
               TabIndex        =   7
               Tag             =   "Partner"
               Top             =   1125
               Width           =   1785
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "DescID"
               DataSource      =   "Adodc1"
               Height          =   330
               Index           =   0
               Left            =   1275
               MaxLength       =   15
               TabIndex        =   2
               Tag             =   "Partner"
               Top             =   45
               Width           =   1995
            End
            Begin VB.TextBox txtBox 
               Appearance      =   0  'Flat
               DataField       =   "Keterangan"
               DataSource      =   "Adodc1"
               Height          =   330
               Index           =   1
               Left            =   4515
               MaxLength       =   50
               TabIndex        =   3
               Tag             =   "Partner"
               Top             =   30
               Width           =   4950
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Ta&x"
               DataField       =   "TAX"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   90
               TabIndex        =   6
               Tag             =   "Partner"
               Top             =   1140
               Width           =   1395
            End
            Begin VB.TextBox txtBox 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "CurrID"
               DataSource      =   "Adodc1"
               Height          =   330
               Index           =   6
               Left            =   5580
               MaxLength       =   15
               TabIndex        =   11
               Tag             =   "Partner"
               Top             =   1507
               Width           =   1785
            End
            Begin VB.CommandButton cmdLink 
               Enabled         =   0   'False
               Height          =   315
               Left            =   7365
               Picture         =   "FrmManDescriptor.frx":68A6
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1515
               Width           =   330
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               DataField       =   "TypeID"
               Height          =   330
               Index           =   0
               Left            =   1275
               TabIndex        =   4
               Tag             =   "Partner"
               Top             =   405
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "Description"
               BoundColumn     =   "TypeID"
               Text            =   "DataCombo1"
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   2250
               Index           =   0
               Left            =   75
               TabIndex        =   27
               Top             =   2565
               Width           =   9390
               _ExtentX        =   16563
               _ExtentY        =   3969
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               BackColor       =   16577005
               ForeColor       =   7159830
               HeadLines       =   2
               RowHeight       =   16
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
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "Tipe Cost"
                  Caption         =   "Tipe Cost"
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
               BeginProperty Column01 
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
               BeginProperty Column02 
                  DataField       =   "Cost"
                  Caption         =   "Cost"
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
                  DataField       =   "Percentage"
                  Caption         =   "Percentage"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0;(0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
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
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
               EndProperty
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               DataField       =   "PartnerID"
               Height          =   330
               Index           =   1
               Left            =   1275
               TabIndex        =   8
               Tag             =   "Partner"
               Top             =   1515
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "Nama Perusahaan"
               BoundColumn     =   "PartnerID"
               Text            =   "DataCombo1"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
               ForeColor       =   &H80000005&
               Height          =   210
               Index           =   5
               Left            =   105
               TabIndex        =   36
               Top             =   2265
               Width           =   405
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000001&
               Index           =   5
               X1              =   1530
               X2              =   105
               Y1              =   2505
               Y2              =   2505
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier"
               ForeColor       =   &H80000005&
               Height          =   210
               Index           =   4
               Left            =   105
               TabIndex        =   35
               Top             =   1575
               Width           =   645
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000001&
               Index           =   4
               X1              =   4305
               X2              =   105
               Y1              =   1830
               Y2              =   1830
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   3210
               X2              =   105
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit"
               ForeColor       =   &H80000005&
               Height          =   210
               Index           =   3
               Left            =   105
               TabIndex        =   34
               Top             =   825
               Width           =   330
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000001&
               Index           =   1
               X1              =   7545
               X2              =   3345
               Y1              =   345
               Y2              =   345
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   3210
               X2              =   105
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipe ID"
               ForeColor       =   &H80000005&
               Height          =   210
               Index           =   0
               Left            =   105
               TabIndex        =   33
               Top             =   105
               Width           =   600
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan"
               ForeColor       =   &H80000005&
               Height          =   210
               Index           =   1
               Left            =   3390
               TabIndex        =   32
               Top             =   90
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
               ForeColor       =   &H80000005&
               Height          =   210
               Index           =   2
               Left            =   105
               TabIndex        =   31
               Top             =   465
               Width           =   420
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000001&
               Index           =   2
               X1              =   4305
               X2              =   105
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   6
               Left            =   1635
               TabIndex        =   30
               Top             =   1185
               Width           =   675
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   3795
               X2              =   1635
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Line Line1 
               Index           =   7
               X1              =   6765
               X2              =   4605
               Y1              =   1455
               Y2              =   1455
            End
            Begin VB.Line Line1 
               Index           =   8
               X1              =   1470
               X2              =   75
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Line Line1 
               Index           =   13
               X1              =   6750
               X2              =   4590
               Y1              =   1822
               Y2              =   1822
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   11
               Left            =   4605
               TabIndex        =   29
               Top             =   1575
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Cost"
               ForeColor       =   &H80000005&
               Height          =   195
               Index           =   12
               Left            =   4605
               TabIndex        =   28
               Top             =   1208
               Width           =   660
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            ForeColor       =   &H80000008&
            Height          =   4920
            Left            =   -74880
            ScaleHeight     =   4890
            ScaleWidth      =   9570
            TabIndex        =   24
            Top             =   390
            Width           =   9600
            Begin MSComctlLib.ListView ListView1 
               Height          =   4815
               Left            =   45
               TabIndex        =   25
               Top             =   45
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   8493
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15380335
               Appearance      =   0
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Descriptor ID"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Keterangan"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Descriptor Type"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            ForeColor       =   &H80000008&
            Height          =   4920
            Left            =   -74910
            ScaleHeight     =   4890
            ScaleWidth      =   9570
            TabIndex        =   14
            Top             =   390
            Width           =   9600
            Begin VB.Frame Frame1 
               BackColor       =   &H00EAAF6F&
               Caption         =   "QTY Purchased"
               Height          =   3060
               Left            =   120
               TabIndex        =   15
               Top             =   135
               Width           =   4785
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   0
                  Left            =   2055
                  TabIndex        =   19
                  Text            =   "0"
                  Top             =   300
                  Width           =   2490
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   2
                  Left            =   2055
                  TabIndex        =   18
                  Text            =   "0"
                  Top             =   990
                  Width           =   2490
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   3
                  Left            =   2055
                  TabIndex        =   17
                  Top             =   2400
                  Width           =   2490
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   1
                  Left            =   2055
                  TabIndex        =   16
                  Text            =   "0"
                  Top             =   645
                  Width           =   2490
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last 30 Days"
                  ForeColor       =   &H80000005&
                  Height          =   210
                  Index           =   7
                  Left            =   270
                  TabIndex        =   23
                  Top             =   360
                  Width           =   1035
               End
               Begin VB.Line Line1 
                  Index           =   9
                  X1              =   3375
                  X2              =   270
                  Y1              =   615
                  Y2              =   615
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last 90 Days"
                  ForeColor       =   &H80000005&
                  Height          =   210
                  Index           =   8
                  Left            =   270
                  TabIndex        =   22
                  Top             =   705
                  Width           =   1035
               End
               Begin VB.Line Line1 
                  Index           =   10
                  X1              =   3375
                  X2              =   270
                  Y1              =   960
                  Y2              =   960
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last 365 Days"
                  ForeColor       =   &H80000005&
                  Height          =   210
                  Index           =   9
                  Left            =   255
                  TabIndex        =   21
                  Top             =   1050
                  Width           =   1140
               End
               Begin VB.Line Line1 
                  Index           =   11
                  X1              =   3360
                  X2              =   255
                  Y1              =   1305
                  Y2              =   1305
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Warranty"
                  ForeColor       =   &H80000005&
                  Height          =   210
                  Index           =   10
                  Left            =   255
                  TabIndex        =   20
                  Top             =   2460
                  Width           =   750
               End
               Begin VB.Line Line1 
                  Index           =   12
                  X1              =   3360
                  X2              =   255
                  Y1              =   2715
                  Y2              =   2715
               End
            End
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmManDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDes As New DBQuick
Private RcPart As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner As New DBQuick
Private mAdd As Boolean

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_Click()
OpenPartner 2
End Sub

Private Sub DataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataGrid1_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DataGrid1_LostFocus(Index As Integer)
If mAdd = True Then MyDDE.SetFocus
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If mAdd = True Then
   Select Case DataGrid1(Index).col
          Case 1, 2, 3: DataGrid1(Index).AllowUpdate = True
          Case Else: DataGrid1(Index).AllowUpdate = False
   End Select
   DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
Else
   DataGrid1(Index).AllowUpdate = False
   DataGrid1(Index).MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
Set mCall = New frmCaller
SSTab1.Tab = 0
OpenDesc
OpenPart
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmManDescriptor
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     DescID, Description AS Keterangan, TypeID, UOM, UnitPrice, UnitCost, PartnerID, TAX, Note AS Catatan,CurrID FROM         [Descriptor Header] ORDER BY DescID"
End With
ListView1.ColumnHeaders(1).width = ((ListView1.width / 1.8) / 2) '+ 184
ListView1.ColumnHeaders(2).width = (ListView1.width / 2)  '+ 184
ListView1.ColumnHeaders(3).width = (ListView1.width / 2)  '+ 184
OpenHeader
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmManDescriptor = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmManDescriptor = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mCall = Nothing
Set FrmManDescriptor = Nothing
End Sub

Private Sub ListView1_DblClick()
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   SSTab1.Tab = 1
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   MyDDE.FindStringData "DescID='" & Item.Text & "'"
End If
End Sub

Private Sub ListView1_LostFocus()
If SSTab1.Tab = 0 Then MyDDE.SetFocus
End Sub

Private Sub mCall_BeforeUnload()
Select Case mCall.Caption
       Case "COST ELEMENT"
            If FindOwnRecordset(MyDDE.ChildRecordset, "[Tipe Cost] = '" & MyDDE.ChildRecordset.Fields("Tipe Cost") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Tipe Cost") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If Not IsNull(MyDDE.ChildRecordset.Fields("Tipe Cost")) = True Then
                  If MyDDE.ChildRecordset.Fields("Tipe Cost") = "" Then
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  End If
               End If
            End If
End Select
mAdd = DataCombo1(0).Enabled
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "COST ELEMENT"
            With MyDDE.ChildRecordset
                 .Fields("Tipe Cost") = mCall.GetFieldByName(0)
                 .Fields("Keterangan") = mCall.GetFieldByName(1)
                 .Fields("Cost") = 0
                 .Fields("Percentage") = 0
            End With
       Case "Master Currency": MyDDE.GetFieldByName("Currid") = mCall.GetFieldByName(0)
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
            txtBox(0).SetFocus
            With MyDDE
                 .GetFieldByName("Keterangan") = "-"
                 .GetFieldByName("UOM") = "PCS"
                 .GetFieldByName("UnitPrice") = 0
                 .GetFieldByName("UnitCost") = 0
                 .GetFieldByName("Catatan") = "-"
            End With
       Case tmbDetail:
            OpenPartner 1
       Case tmbEdit:
            txtBox(0).Enabled = False
            mAdd = True
            txtBox(1).SetFocus
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
               mAdd = True
            Else
               mAdd = False
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
                  With MyDDE.ChildRecordset
                       If .Recordcount <> 0 Then
                          If SendDataToServer("DELETE FROM [Descriptor Detail Costing] WHERE     (DescID = N'" & txtBox(0) & "')") = True Then
                            .MoveFirst
                            Do
                              If .EOF = True Then Exit Do
                              SendDataToServer ("INSERT INTO [Descriptor Detail Costing]" & _
                                                " (DescID, [Cost Element Type], Cost, Percentage)" & _
                                                " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Tipe Cost") & "', " & .Fields("Cost") & ", " & .Fields("Percentage") & ")")
                              .MoveNext
                            Loop
                            .MoveLast
                          End If
                       End If
                  End With
                  mAdd = False
            End If
       Case tmbPrint:
            CallRPTReport "Descriptor.rpt", "SELECT  * From [Descriptor] Where [Desc ID]=N'" & txtBox(0) & "'"
       Case Else: 'mVarDataDc = False
End Select
mAdd = DataCombo1(0).Enabled
cmdLink.Enabled = DataCombo1(0).Enabled
txtBox(6).Enabled = False
End Sub

Private Sub MyDDE_LostFocus()
'If SSTab1.TabEnabled(SSTab1.Tab) = True And mAdd = False Then SSTab1.SetFocus
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("DescID")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
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

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Descriptor Header] ([DescID], [Description],TypeID, UOM, UnitPrice, UnitCost, PartnerID, TAX, Note,CurrID) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "',N'" & DataCombo1(0).BoundText & "',N'" & txtBox(2) & "'," & CDbl(txtBox(3)) & ", " & CDbl(txtBox(4)) & ",N'" & DataCombo1(1).BoundText & "'," & Check1.Value & ",N'" & txtBox(5) & "',N'" & txtBox(6) & "')"
                     
    .PrepareUpdate = " UPDATE [Descriptor Header] Set CurrID =N'" & txtBox(6) & "', [Description] = N'" & txtBox(1) & "',TypeID=N'" & DataCombo1(0).BoundText & "', UOM=N'" & txtBox(2) & "', UnitPrice=" & CDbl(txtBox(3)) & ", UnitCost=" & CDbl(txtBox(4)) & ", PartnerID=N'" & DataCombo1(1).BoundText & "', TAX=" & Check1.Value & ", Note=N'" & txtBox(5) & "' WHERE     ([DescID] = N'" & txtBox(0) & "')"
    'MessageBox .PrepareUpdate
    .PrepareDelete = " DELETE FROM [Descriptor Header] WHERE   ([DescID] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub GridLayout()
'DataGrid1(0).Height = 2225
DataGrid1(0).width = 9390
Check1.BackColor = &HEAAF6F
End Sub

Private Sub OpenDesc()
RcDes.DBOpen "SELECT     TypeID, Description FROM         [Descriptor Type] ORDER BY Description", CNN, lckLockReadOnly
Set DataCombo1(0).RowSource = RcDes.DBRecordset
End Sub

Private Sub OpenPart()
RcPart.DBOpen "SELECT     PartnerID, CompanyName AS [Nama Perusahaan]  FROM         PartnerDB WHERE     (PartnerType = N'SUPPLIER') ORDER BY CompanyName", CNN, lckLockReadOnly
Set DataCombo1(1).RowSource = RcPart.DBRecordset
End Sub

Private Sub OpenHeader()
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Set Rc.DBRecordset = MyDDE.ActiveRecordset.Clone(adLockReadOnly)
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            With ListView1.ListItems.Add(, , Avdata(0, I))
                 .SubItems(1) = Avdata(1, I)
                 .SubItems(2) = Avdata(2, I)
            End With
        Next I
     Else
     End If
End With
Rc.CloseDB
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1:
            RcPartner.DBOpen " SELECT     [Cost Element Type] AS [Tipe Cost], Description AS Keterangan FROM         [Cost Element] ORDER BY [Cost Element Type]", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen " SELECT     * from [Currency Setup]", CNN, lckLockReadOnly
            
End Select
If RcPartner.Recordcount <> 0 Then
Select Case Index
       Case 1:
            mCall.FromTagActive = "COST ELEMENT"
       Case 2:
            mCall.FromTagActive = "Master Currency"
End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Cost Element Setup Belum Ada.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal Param As String)
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     [Cost Element].[Cost Element Type] AS [Tipe Cost], [Cost Element].Description AS Keterangan, [Descriptor Detail Costing].Cost,                        [Descriptor Detail Costing].Percentage FROM         [Cost Element] INNER JOIN                       [Descriptor Detail Costing] ON [Cost Element].[Cost Element Type] = [Descriptor Detail Costing].[Cost Element Type] WHERE     ([Descriptor Detail Costing].DescID = N'" & Param & "') ORDER BY [Descriptor Detail Costing].DescID", CNN, lckLockBatch
Set MyDDE.ChildRecordset = Rc.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub OpenQTY()
'Dim RcQty As New DBQuick
'RcQty.DBOpen "SELECT     SUM(QTYI) AS QTYI, SUM(QTYII) AS QTYII, SUM(QTYIII) AS QTYIII FROM         [History Issued] WHERE     (NoItem = N'" & MyDDE.GetFieldByName("NoItem") & "')", Cnn, lckLockReadOnly
'With RcQty.DBRecordset
'     If .Recordcount <> 0 Then
'        TxtReceipt(0) = FormatNumber(IIf(Not IsNull(.Fields(0)), .Fields(0), 0), 0)
'        TxtReceipt(1) = FormatNumber(IIf(Not IsNull(.Fields(1)), .Fields(1), 0), 0)
'        TxtReceipt(2) = FormatNumber(IIf(Not IsNull(.Fields(2)), .Fields(2), 0), 0)
'     Else
'        TxtReceipt(0) = 0
'        TxtReceipt(1) = 0
'        TxtReceipt(2) = 0
'     End If
'End With
'RcQty.DBOpen "SELECT     SUM(QTYI) AS QTYI, SUM(QTYII) AS QTYII, SUM(QTYIII) AS QTYIII FROM         [History Receipt] WHERE     (NoItem = N'" & MyDDE.GetFieldByName("NoItem") & "')", Cnn, lckLockReadOnly
'With RcQty.DBRecordset
'     If .Recordcount <> 0 Then
'        TxtReceipt(3) = FormatNumber(IIf(Not IsNull(.Fields(0)), .Fields(0), 0), 0)
'        TxtReceipt(4) = FormatNumber(IIf(Not IsNull(.Fields(1)), .Fields(1), 0), 0)
'        TxtReceipt(5) = FormatNumber(IIf(Not IsNull(.Fields(2)), .Fields(2), 0), 0)
'     Else
'        TxtReceipt(3) = 0
'        TxtReceipt(4) = 0
'        TxtReceipt(5) = 0
'     End If
'End With
'RcQty.CloseDB
'Set RcQty = Nothing
End Sub
