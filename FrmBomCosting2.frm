VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{2D9B7C69-DEB8-4853-83CE-4B327B4C1B03}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmBomCosting2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBomCosting2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Tag             =   "Product Costing"
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
      Height          =   5940
      Left            =   45
      ScaleHeight     =   5910
      ScaleWidth      =   9795
      TabIndex        =   1
      Top             =   0
      Width           =   9825
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Height          =   5250
         Left            =   105
         ScaleHeight     =   5220
         ScaleWidth      =   9525
         TabIndex        =   2
         Top             =   570
         Width           =   9555
         Begin TabDlg.SSTab SSTab1 
            Height          =   5070
            Left            =   75
            TabIndex        =   3
            Top             =   105
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   8943
            _Version        =   393216
            Style           =   1
            Tab             =   2
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
            TabCaption(0)   =   "BOM Listing"
            TabPicture(0)   =   "FrmBomCosting2.frx":6852
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Picture3"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Detail Costing"
            TabPicture(1)   =   "FrmBomCosting2.frx":686E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Picture4"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Analyze"
            TabPicture(2)   =   "FrmBomCosting2.frx":688A
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Picture5"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00EAAF6F&
               Height          =   4620
               Left            =   -74925
               ScaleHeight     =   4560
               ScaleWidth      =   9150
               TabIndex        =   7
               Top             =   375
               Width           =   9210
               Begin VB.TextBox txtBox 
                  Appearance      =   0  'Flat
                  DataField       =   "Kode Barang"
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
                  Left            =   1410
                  MaxLength       =   15
                  TabIndex        =   14
                  Tag             =   "Partner"
                  Text            =   "Text1"
                  Top             =   75
                  Width           =   2250
               End
               Begin VB.TextBox txtBox 
                  Appearance      =   0  'Flat
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
                  Height          =   315
                  Index           =   1
                  Left            =   6645
                  MaxLength       =   15
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   13
                  Tag             =   "Partner"
                  Text            =   "Text1"
                  Top             =   405
                  Width           =   2250
               End
               Begin VB.TextBox txtBox 
                  Appearance      =   0  'Flat
                  DataField       =   "Nama Barang"
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
                  Left            =   1410
                  MaxLength       =   50
                  TabIndex        =   12
                  Tag             =   "Partner"
                  Text            =   "Text1"
                  Top             =   405
                  Width           =   3945
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Fixed Cost"
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
                  Height          =   315
                  Index           =   3
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   13
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   10
                  Tag             =   "Partner"
                  Text            =   "Text1"
                  Top             =   735
                  Width           =   2250
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Average Cost"
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0;(#,##0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
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
                  Height          =   315
                  Index           =   4
                  Left            =   6645
                  Locked          =   -1  'True
                  MaxLength       =   13
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   9
                  Tag             =   "Partner"
                  Text            =   "Text1"
                  Top             =   735
                  Width           =   2250
               End
               Begin VB.TextBox txtBox 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  DataField       =   "Last Cost"
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
                  Height          =   315
                  Index           =   5
                  Left            =   1410
                  MaxLength       =   13
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   8
                  Tag             =   "Partner"
                  Text            =   "Text1"
                  Top             =   1065
                  Width           =   2250
               End
               Begin MSDataGridLib.DataGrid DataGrid1 
                  Bindings        =   "FrmBomCosting2.frx":68A6
                  Height          =   2205
                  Index           =   0
                  Left            =   135
                  TabIndex        =   11
                  Top             =   1440
                  Width           =   8910
                  _ExtentX        =   15716
                  _ExtentY        =   3889
                  _Version        =   393216
                  AllowUpdate     =   0   'False
                  Appearance      =   0
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
                  ColumnCount     =   3
                  BeginProperty Column00 
                     DataField       =   "Cost Element"
                     Caption         =   "Cost Element"
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
                  EndProperty
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama BOM                                                                                                    UOM"
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
                  Left            =   195
                  TabIndex        =   20
                  Top             =   450
                  Width           =   5625
               End
               Begin VB.Line Line1 
                  Index           =   2
                  X1              =   195
                  X2              =   1620
                  Y1              =   705
                  Y2              =   705
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "BOM ID"
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
                  Left            =   195
                  TabIndex        =   19
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Line Line1 
                  Index           =   0
                  X1              =   195
                  X2              =   1620
                  Y1              =   375
                  Y2              =   375
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fixed Cost                                                                                                    Average Cost"
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
                  Left            =   195
                  TabIndex        =   18
                  Top             =   780
                  Width           =   6255
               End
               Begin VB.Line Line1 
                  Index           =   4
                  X1              =   5475
                  X2              =   6900
                  Y1              =   705
                  Y2              =   705
               End
               Begin VB.Line Line1 
                  Index           =   1
                  X1              =   5475
                  X2              =   6900
                  Y1              =   1035
                  Y2              =   1035
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last Cost"
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
                  Left            =   195
                  TabIndex        =   17
                  Top             =   1110
                  Width           =   675
               End
               Begin VB.Line Line1 
                  Index           =   3
                  X1              =   195
                  X2              =   1620
                  Y1              =   1365
                  Y2              =   1365
               End
               Begin VB.Line Line1 
                  Index           =   5
                  X1              =   195
                  X2              =   1620
                  Y1              =   1035
                  Y2              =   1035
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cost Rollup"
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
                  TabIndex        =   16
                  Top             =   3750
                  Width           =   810
               End
               Begin VB.Label LblAmount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Label2"
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
                  Height          =   315
                  Left            =   6795
                  TabIndex        =   15
                  Top             =   3690
                  Width           =   2250
               End
               Begin VB.Line Line1 
                  Index           =   6
                  X1              =   5355
                  X2              =   6900
                  Y1              =   3990
                  Y2              =   3990
               End
            End
            Begin VB.PictureBox Picture3 
               Height          =   4620
               Left            =   -74925
               ScaleHeight     =   4560
               ScaleWidth      =   9150
               TabIndex        =   5
               Top             =   375
               Width           =   9210
               Begin MSComctlLib.ListView ListView1 
                  Height          =   4575
                  Left            =   0
                  TabIndex        =   6
                  Top             =   0
                  Width           =   9165
                  _ExtentX        =   16166
                  _ExtentY        =   8070
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  HotTracking     =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   15380335
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   6
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Kode Barang"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Nama Barang"
                     Object.Width           =   5644
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "UOM"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   3
                     Text            =   "Fixed Cost"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   4
                     Text            =   "Average Cost"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   5
                     Text            =   "Last Cost"
                     Object.Width           =   2540
                  EndProperty
               End
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00EAAF6F&
               Height          =   4620
               Left            =   75
               ScaleHeight     =   4560
               ScaleWidth      =   9150
               TabIndex        =   4
               Top             =   375
               Width           =   9210
            End
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmBomCosting2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAdd As Boolean
Private RcPart As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mFirstCaller, mKeyLoad As Boolean

Private Sub DataGrid1_AfterColEdit(Index As Integer, ByVal ColIndex As Integer)
If DataGrid1(0).Col = 2 And mAdd = True Then TotalTrans
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
'Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If mAdd = True Then
   DataGrid1(0).MarqueeStyle = dbgFloatingEditor
   Select Case DataGrid1(0).Col
          Case 2:
               DataGrid1(Index).AllowUpdate = mAdd
          Case Else
               'DataGrid1(Index).MarqueeStyle = dbgFloatingEditor
               DataGrid1(Index).AllowUpdate = False
   End Select
Else
   'DataGrid1(0).MarqueeStyle = dbgHighlightRow
   DataGrid1(0).AllowUpdate = mAdd
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mKeyLoad = False Then mKeyLoad = True Else mKeyLoad = False
If mKeyLoad = False Then ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
GridLayout
SSTab1.Tab = 0
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmBomCosting
    .SetPermissions = UserAddnewDenied
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT NoItem AS [Kode Barang], ItemName AS [Nama Barang], UOM, FixCost AS [Fixed Cost], AvgCost AS [Average Cost], LastCost AS [Last Cost] FROM         Inventory WHERE     (Manufacture = 1)"
End With
OpenHeader
Set mCall = New frmCaller
DataGrid1(0).AllowUpdate = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmBomCosting = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmBomCosting = Nothing
End If
End Sub

Private Sub Form_Resize()
GridLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mCall = Nothing
Set FrmBomCosting = Nothing
End Sub



Private Sub ListView1_DblClick()
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   SSTab1.Tab = 1
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If MyDDE.ActiveRecordset.Recordcount <> 0 Then
   MyDDE.FindStringData "[Kode Barang]='" & Item.Text & "'"
End If
End Sub

Private Sub mCall_BeforeUnload()
If FindOwnRecordset(MyDDE.ChildRecordset, "[Cost Element] = '" & MyDDE.ChildRecordset.Fields("Cost Element") & "'") = True Then
   MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Cost Element") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
   MyDDE.ChildRecordset.CancelBatch adAffectCurrent
   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
Else
   If Not IsNull(MyDDE.ChildRecordset.Fields("Cost Element")) = True Then
      If MyDDE.ChildRecordset.Fields("Cost Element") = "" Then
         MyDDE.ChildRecordset.CancelBatch adAffectCurrent
         If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
      End If
   End If
End If
mAdd = txtBox(3).Enabled
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case mCall.FromTagActive
       Case "COST ELEMENT":
            With MyDDE.ChildRecordset
                 .Fields("Cost Element") = mCall.GetFieldByName(0)
                 .Fields("Keterangan") = mCall.GetFieldByName(1)
                 .Fields("Cost") = 0
            End With
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            mAdd = True
            txtBox(0).SetFocus
            'Label2 = IndexAuto

            SSTab1.Tab = 1
       Case tmbEdit:
            txtBox(0).Enabled = False
            txtBox(1).Enabled = False
            txtBox(2).Enabled = False
            mAdd = True
            txtBox(3).SetFocus
            SSTab1.Tab = 1
            TotalTrans
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
            Select Case MyDDE.ChildRecordset.Status
                   Case 8: mAdd = False
                   Case Else
                        If MyDDE.ChildRecordset.Recordcount <> 0 Then
                           mAdd = True
                        Else
                           mAdd = False
                        End If
            End Select
            Else
               mAdd = False
            End If
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then OpenDetailPartner SSTab1.Tab
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  With MyDDE.ChildRecordset
                       .MoveFirst
                       If SendDataToServer("Delete From [BOM Costing Detail] WHERE  (NoItem = N'" & txtBox(0) & "')") = True Then
                          Do
                          If MyDDE.ChildRecordset.EOF Then Exit Do
                             SendDataToServer " INSERT INTO [BOM Costing Detail]" & _
                                              " ( NoItem, [Cost Element Type], CostValue)" & _
                                              " VALUES (N'" & txtBox(0) & "', N'" & .Fields("Cost Element") & "', " & .Fields("Cost") & ")"
                             .MoveNext
                          Loop
                       End If
                       .MoveLast
                  End With
               End If
               mAdd = False
            End If
       Case tmbPrint:
            CallRPTReport "BOM Costing List.rpt", "sELECT * FROM [BOM Costing List] Where [BOM ID] =N'" & txtBox(0) & "'"
       Case Else: 'mVarDataDc = False
End Select
mAdd = txtBox(3).Enabled
SSTab1.TabEnabled(0) = Not mAdd
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("Kode Barang")
TotalTrans
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
'               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
'               Else
'                  MyDDE.CancelTrans = True
'                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
'                  MyDDE.IsChildMemberReady = False
'               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  If mAdd = True Then txtBox(3) = CDbl(LblAmount)
                  PrepareQuery
               Else
                  'MessageBox "Date detail calendar belum ada.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
                  PrepareQuery
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDetail:
            MyDDE.CancelTrans = mFirstCaller
            If MyDDE.CancelTrans = True Then Exit Sub
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
               If MyDDE.ChildRecordset.Fields(2) = 0 Then
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

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    '.PrepareAppend = " INSERT INTO [Inventory]" & _
                     " (NoItem, ItemName, UOM, MethodeID, Phantom,Manufacture)" & _
                     " VALUES  (N'" & txtBox(0) & "', N'" & txtBox(2) & "', N'" & txtBox(1) & "', N'" & DataCombo1.BoundText & "', " & Check1.Value & ",1)"
'MessageBox .PrepareAppend
    .PrepareUpdate = " UPDATE [Inventory] Set FixCost=" & CDbl(txtBox(3)) & ",AvgCost=" & CDbl(txtBox(4)) & ",LastCost=" & CDbl(txtBox(5)) & " WHERE     ([NoItem] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Inventory] WHERE   ([NoItem] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2160
DataGrid1(0).Width = 8910
DataGrid1(0).Columns(0).Width = 2340.284
DataGrid1(0).Columns(1).Width = 4169.764
DataGrid1(0).Columns(2).Width = 1830.047
End Sub

Private Sub OpenDetail(ByVal Param As String)
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     [BOM Costing Detail].[Cost Element Type] AS [Cost Element], [Cost Element].Description AS Keterangan, [BOM Costing Detail].CostValue AS Cost FROM         [BOM Costing Detail] INNER JOIN                       [Cost Element] ON [BOM Costing Detail].[Cost Element Type] = [Cost Element].[Cost Element Type] WHERE     ([BOM Costing Detail].NoItem = N'" & Param & "') ORDER BY [BOM Costing Detail].[Cost Element Type]", CNN, lckLockBatch
Set MyDDE.ChildRecordset = Rc.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1(0).DataSource = MyDDE.ChildRecordset
End Sub

Private Sub OpenDetailPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 1: RcPartner.DBOpen "SELECT     [Cost Element Type] AS [Cost Element], Description AS Keterangan FROM         [Cost Element] ORDER BY [Cost Element Type]", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 1: mCall.FromTagActive = "COST ELEMENT"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
'    MsgBox Err.Description
    Err.Clear
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
                 .SubItems(3) = FormatNumber(Avdata(3, I), 0)
                 .SubItems(4) = FormatNumber(Avdata(4, I), 0)
                 .SubItems(5) = FormatNumber(Avdata(5, I), 0)
            End With
        Next I
     Else
     End If
End With
End Sub

Private Sub TotalTrans()
Dim Rc As New DBQuick
Dim I As Integer
Dim Avdata As Variant
Set Rc.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
LblAmount = 0
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            LblAmount = LblAmount + IIf(Not IsNull(Avdata(2, I)), Avdata(2, I), 0)
        Next I
        LblAmount = FormatNumber(LblAmount, 0)
     End If
End With
Set Avdata = Nothing
End Sub
