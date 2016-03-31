VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4494886D-7C13-4468-B1F9-FC63570B5F46}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMRPInventory 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12090
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12090
   Tag             =   "MRP Inventory"
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   405
      Left            =   5025
      TabIndex        =   28
      Top             =   6450
      Width           =   1530
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   405
      Left            =   3075
      TabIndex        =   25
      Top             =   6450
      Width           =   1530
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Mrp"
      Height          =   405
      Left            =   1530
      TabIndex        =   24
      Top             =   6450
      Width           =   1530
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Generate MRP"
      Height          =   390
      Left            =   60
      TabIndex        =   22
      Top             =   6450
      Width           =   1455
   End
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
      Height          =   6525
      Left            =   15
      ScaleHeight     =   6495
      ScaleWidth      =   11865
      TabIndex        =   20
      Top             =   30
      Width           =   11895
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Height          =   5580
         Left            =   135
         ScaleHeight     =   5550
         ScaleWidth      =   11640
         TabIndex        =   21
         Top             =   195
         Width           =   11670
         Begin VB.ComboBox List1 
            DataField       =   "Plan Horizon"
            Height          =   315
            ItemData        =   "FrmMRPInventory.frx":0000
            Left            =   5835
            List            =   "FrmMRPInventory.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Tag             =   "PO"
            Top             =   45
            Width           =   795
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Safety Stock"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   9480
            TabIndex        =   18
            Tag             =   "PO"
            Text            =   "50"
            Top             =   990
            Width           =   1470
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Require Date"
            Height          =   285
            Index           =   0
            Left            =   6450
            TabIndex        =   11
            Tag             =   "PO"
            Top             =   375
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yy"
            Format          =   51183619
            CurrentDate     =   38624
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Planning Horizon (                  )"
            DataField       =   "Order Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4515
            TabIndex        =   9
            Tag             =   "PO"
            Top             =   135
            Width           =   2595
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "QTY Order"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   6465
            TabIndex        =   17
            Tag             =   "PO"
            Text            =   "300"
            Top             =   990
            Width           =   1470
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Yield Percentage"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   9480
            TabIndex        =   15
            Tag             =   "PO"
            Text            =   "0"
            Top             =   690
            Width           =   1470
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Lead Time"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   6465
            TabIndex        =   14
            Tag             =   "PO"
            Text            =   "3"
            Top             =   690
            Width           =   1470
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Lot Size"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2670
            TabIndex        =   7
            Tag             =   "PO"
            Text            =   "100"
            Top             =   990
            Width           =   1470
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2670
            TabIndex        =   5
            Tag             =   "PO"
            Text            =   "20"
            Top             =   690
            Width           =   1470
         End
         Begin VB.TextBox txtMrp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2670
            TabIndex        =   3
            Tag             =   "PO"
            Text            =   "50"
            Top             =   390
            Width           =   1470
         End
         Begin VB.TextBox txtMrp 
            Appearance      =   0  'Flat
            DataField       =   "NoItem"
            Height          =   285
            Index           =   0
            Left            =   2670
            TabIndex        =   1
            Tag             =   "PO"
            Text            =   "FG/000000000005"
            Top             =   90
            Width           =   1470
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Use even multiples of Lot size"
            DataField       =   "Multiple Lot"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   8
            Tag             =   "PO"
            Top             =   1320
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid DgJadwal 
            Height          =   3780
            Left            =   90
            TabIndex        =   19
            Top             =   1650
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   6668
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            TabAcrossSplits =   -1  'True
            TabAction       =   2
            WrapCellPointer =   -1  'True
            RowDividerStyle =   3
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "End Date"
            Height          =   285
            Index           =   1
            Left            =   9480
            TabIndex        =   12
            Tag             =   "PO"
            Top             =   375
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yy"
            Format          =   51183619
            CurrentDate     =   38624
         End
         Begin VB.Label LblOrder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Make To Order"
            Height          =   195
            Left            =   7200
            TabIndex        =   23
            Top             =   135
            Width           =   1065
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   8250
            X2              =   10590
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   8250
            X2              =   10590
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   8250
            X2              =   10590
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Require Date                                                             Entry Date"
            Height          =   195
            Index           =   7
            Left            =   4560
            TabIndex        =   10
            Top             =   435
            Width           =   4470
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   4545
            X2              =   6885
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   4575
            X2              =   6915
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   4575
            X2              =   6915
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   330
            X2              =   2670
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   330
            X2              =   2670
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   345
            X2              =   2685
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   360
            X2              =   2700
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty Order                                                                  Safety Stock"
            Height          =   195
            Index           =   5
            Left            =   4560
            TabIndex        =   16
            Top             =   1035
            Width           =   4620
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Time in Weeks                                                  Yield percentage"
            Height          =   195
            Index           =   4
            Left            =   4560
            TabIndex        =   13
            Top             =   735
            Width           =   4860
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lot size (Use 1 for L4L)"
            Height          =   195
            Index           =   3
            Left            =   345
            TabIndex        =   6
            Top             =   1035
            Width           =   1650
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Units Allocated"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   4
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beginning Inventory"
            Height          =   195
            Index           =   1
            Left            =   345
            TabIndex        =   2
            Top             =   435
            Width           =   1455
         End
         Begin VB.Label LblMrp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Identifier"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   0
            Top             =   135
            Width           =   1035
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   27
      Top             =   6375
      Visible         =   0   'False
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmMRPInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcJadwal As New Recordset
Private mAwal, mAkhir As Integer
Dim mTotal, mReceipt, mGross As Variant
Private mList As String
Private vColindex As Integer

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
   LblOrder.Caption = "Make To Order"
   DTPicker1(0).Enabled = True
Else
   LblOrder.Caption = "Make To Stock"
   DTPicker1(0).Enabled = False
End If
End Sub

Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub CmdOK_Click()
mTotal = CDbl(txtMrp(6))
Select Case mList
       Case "Day": GenerateJadwal DayOfYear(DTPicker1(0)), HitungHari
       Case "Week": GenerateJadwal WeekOfYear(DTPicker1(0)), HitungHari
       Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
End Select
OpenCucakRowo MyDDE.GetFieldByName("NoItem")
End Sub

Private Sub Command1_Click()
'
RcJadwal.CursorLocation = adUseClient
RcJadwal.Open App.Path & "\FG5.Dat", , adOpenForwardOnly, adLockBatchOptimistic, adCmdFile
Set DgJadwal.DataSource = RcJadwal
GridLayout
End Sub

Private Sub Command2_Click()
RcJadwal.Save App.Path & "\" & Replace(Replace(txtMrp(0), "/", ""), "0", "") & ".Dat", adPersistADTG
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DgJadwal_AfterColEdit(ByVal ColIndex As Integer)
Dim Rc As New DBQuick
If ColIndex > 3 Then
With RcJadwal
     vColindex = ColIndex
     If DgJadwal.Columns(ColIndex).Value <> "" Then
        Rc.DBOpen "SELECT [BOM Component Detail].Component, Inventory.ItemName, SUM(Inventory.LeadTimeDays) AS LeadTimeDays,  SUM([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT) AS Qty FROM         [BOM Component Detail] INNER JOIN  Inventory ON [BOM Component Detail].Component = Inventory.NoItem INNER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE ([BOM Component Detail].NoItem = N'" & txtMrp(0) & "') AND (Inventory.Manufacture = 1) GROUP BY [BOM Component Detail].Component, Inventory.ItemName ORDER BY [BOM Component Detail].Component", Cnn, lckLockReadOnly
        With Rc.DBRecordset
             If .Recordcount <> 0 Then
                .MoveFirst
                Do
                   If .EOF = True Then Exit Do
'                   LoadJadwal .Fields(0), .Fields(2), .Fields(3)
                    HitungOnHand .Fields(2), Abs(.Fields(3)), 0, 10, IIf(DgJadwal.Columns(ColIndex).Value = "", "0", DgJadwal.Columns(ColIndex).Value)
                   .MoveNext
                Loop
             End If
        End With
        
     End If
End With
End If
End Sub

Private Sub DgJadwal_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DgJadwal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DgJadwal.Col <= 2 Then
   DgJadwal.MarqueeStyle = dbgHighlightRow
   DgJadwal.AllowUpdate = False
ElseIf DgJadwal.Col > 2 Then
   If RcJadwal.AbsolutePosition = 1 Then
      DgJadwal.MarqueeStyle = dbgFloatingEditor
      DgJadwal.AllowUpdate = True
   Else
      DgJadwal.MarqueeStyle = dbgHighlightRow
      DgJadwal.AllowUpdate = False
   End If
End If
End Sub

Private Sub DTPicker1_Change(Index As Integer)
DTPicker1_Click Index
End Sub

Private Sub DTPicker1_Click(Index As Integer)
If DTPicker1(1).Value > DTPicker1(0).Value Then
   DTPicker1(1).Value = DTPicker1(0).Value
End If
HitungHari
End Sub

Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
List1.ListIndex = 0
With MyDDE
     .EditModeReplace = False
     Set .BindForm = FrmMRPInventory
     .BindFormTAG = "PO"
     Set .ActiveConnection = Cnn
     .PrepareQuery = " SELECT  * FROM [MRP INVENTORY] ORDER BY NoItem"
End With
'Command1_Click
End Sub

Private Sub GenerateJadwal(ByVal vAwal As Integer, ByVal vAkhir As Integer)
Dim I As Integer
Set RcJadwal = Nothing
Set RcJadwal = New Recordset
With RcJadwal
     .Fields.Append "ITEM ID", adBSTR
     .Fields.Append "Planning Horizon", adBSTR
     .Fields.Append "Stock", adInteger
     For I = 1 To vAkhir
        .Fields.Append vAwal + I, adInteger
     Next
End With
RcJadwal.Open
Set DgJadwal.DataSource = RcJadwal
mAwal = vAwal
mAkhir = RcJadwal.Fields.Count
GridLayout
Err.Clear
End Sub

Private Sub GridLayout()
Dim I As Integer
DgJadwal.Splits.Add 1
DgJadwal.Splits(0).Columns(0).Visible = True
DgJadwal.Splits(0).Columns(1).Visible = True
DgJadwal.Splits(0).Columns(1).Width = 1750
DgJadwal.Splits(1).Columns(0).Visible = False
DgJadwal.Splits(1).Columns(1).Visible = False
DgJadwal.Splits(0).ScrollBars = dbgHorizontal
DgJadwal.Splits(1).ScrollBars = dbgBoth
DgJadwal.Splits(1).RecordSelectors = False
DgJadwal.Splits(0).AllowSizing = False
DgJadwal.Splits(1).AllowSizing = False
DgJadwal.Splits(1).Size = 2
DgJadwal.Splits(1).SizeMode = dbgScalable
DgJadwal.Splits(1).LeftCol = 0
For I = 0 To RcJadwal.Fields.Count - 1
    If I <= 1 Then
       DgJadwal.Columns(I).Width = 1700
       DgJadwal.Columns(I).Alignment = dbgLeft
    Else
       DgJadwal.Columns(I).Width = 600
       DgJadwal.Columns(I).Alignment = dbgRight
       DgJadwal.Columns(I).NumberFormat = "#,##0;(#,##0)"
       DgJadwal.Splits(0).Columns(I).Visible = False
       DgJadwal.Splits(1).Columns(I).AllowSizing = False
    End If
    DgJadwal.Columns(I).DividerStyle = dbgRaised
Next
DgJadwal.TabAcrossSplits = True
DgJadwal.Refresh
End Sub

Private Sub LoadJadwal(ByVal vBOMItemLoad As String, ByVal vLeadTime As Integer, Optional ByVal vJumlah As Double)
Dim I, j As Integer
With RcJadwal
     For I = 0 To 5
         If I = 0 Then .AddNew 0, vBOMItemLoad Else .AddNew
         Select Case I
                Case 0: .Fields(1) = "Gross Requirement"
                Case 1: .Fields(1) = "Schedule Receipt"
                Case 2: .Fields(1) = "On Hand"
                        .Fields(2) = vJumlah
                Case 3: .Fields(1) = "Net Requirement"
                Case 4: .Fields(1) = "Planned Order Receipt"
                Case 5: .Fields(1) = "Planned Order Release"
         End Select
     Next I
End With
mAkhir = mAkhir - vLeadTime
DgJadwal.Refresh
End Sub

Private Function OpenOnHand(ByVal vNoItem As String) As Double
OpenOnHand = 20
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RcJadwal.Close
Set RcJadwal = Nothing
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMRPInventory = Nothing
End Sub

Private Sub List1_Click()
If List1.Text <> "" Then mList = List1.Text
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'OpenCucakRowo MyDDE.GetFieldByName("NoItem")
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtMrp_Change(Index As Integer)
If Index <> 0 Then If txtMrp(Index) = "" Then txtMrp(Index) = 0
End Sub

Private Sub txtMrp_GotFocus(Index As Integer)
Block txtMrp(Index)
End Sub

Private Sub txtMrp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub txtMrp_KeyPress(Index As Integer, KeyAscii As Integer)
If Index <> 0 Then ValidNum KeyAscii
End Sub

Private Function HitungHari() As Long
Dim mTotal As Long
mTotal = DTPicker1(0).Value - DTPicker1(1).Value
Select Case mList
       Case "Day": HitungHari = mTotal
       Case "Week": HitungHari = Round(mTotal / 7)
       Case "Monthly": HitungHari = Round(mTotal / 30)
End Select
End Function

Private Sub HitungOnHand(ByVal vLeadTime As Double, ByVal vStockAwal As Double, ByVal vPlanned As Double, ByVal vJadwal As Double, ByVal vGross As Double)
On Error GoTo Hell
Dim I As Integer
Dim mTot As Variant
With RcJadwal
     If .Recordcount <> 0 Then
        .AbsolutePosition = .AbsolutePosition + 2
        vStockAwal = .Fields(2)
        For I = 3 To vColindex
            .Fields(I) = (vStockAwal + vPlanned + vJadwal) - vGross
            mTot = .Fields(I)
        Next
        .AbsolutePosition = .AbsolutePosition + 1
        .Fields(vColindex) = vGross - mTot
        
        .AbsolutePosition = .AbsolutePosition + 1
        .Fields(vColindex) = vGross - mTot
        
        .AbsolutePosition = .AbsolutePosition + 1
        .Fields(vColindex - vLeadTime) = vGross - mTot
        
        .AbsolutePosition = .AbsolutePosition + 1
        .Fields(vColindex - vLeadTime) = vStockAwal * CDbl(txtMrp(6))
        vColindex = DgJadwal.Col - vLeadTime
     End If
End With
Exit Sub
Hell:
Err.Clear
End Sub

Private Sub PrepareQuery()
SendDataToServer (" INSERT INTO [MRP INVENTORY] " & _
                  " (NoItem, [QTY Order], [Lot Size], [Multiple Lot], [Plan Horizon], [Order Type], [Require Date], [End Date], [Lead Time], [Yield Percentage], [Safety Stock])" & _
                  " VALUES (N'" & txtMrp(0) & "', " & CDbl(txtMrp(6)) & ", " & CDbl(txtMrp(3)) & ", " & BoolToInt(Check1.Value) & ", N'" & mList & "', " & BoolToInt(Check2.Value) & ", CONVERT(DATETIME, '" & Format(DTPicker1(0), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(DTPicker1(1), "dd/mm/yy") & "', 3), " & CDbl(txtMrp(4)) & ", " & CDbl(txtMrp(5)) & ", " & CDbl(txtMrp(7)) & ")")
                  
'SendDataToServer ("Delete From [MRP INVENTORY] where NoItem='" & txtMrp(0) & "'")
End Sub

Private Sub OpenCucakRowo(ByVal vKode As String)
Dim Rc As New DBQuick
Rc.DBOpen "SELECT [BOM Component Detail].Component, Inventory.ItemName, SUM(Inventory.LeadTimeDays) AS LeadTimeDays,  SUM([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT) AS Qty FROM         [BOM Component Detail] INNER JOIN  Inventory ON [BOM Component Detail].Component = Inventory.NoItem INNER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE ([BOM Component Detail].NoItem = N'" & vKode & "') AND (Inventory.Manufacture = 1) GROUP BY [BOM Component Detail].Component, Inventory.ItemName ORDER BY [BOM Component Detail].Component", Cnn, lckLockReadOnly
With Rc.DBRecordset
     If .Recordcount <> 0 Then
        Do
           If .EOF = True Then Exit Do
           LoadJadwal .Fields(0), .Fields(2), .Fields(3)
           .MoveNext
        Loop
     End If
End With
End Sub




