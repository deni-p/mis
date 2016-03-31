VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMRPLogic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Requirement Planning"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMRPLogic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11700
   Tag             =   "Material Requirement Planning"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   7485
      Left            =   0
      ScaleHeight     =   7485
      ScaleWidth      =   11700
      TabIndex        =   1
      Top             =   0
      Width           =   11700
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5700
         Picture         =   "FrmMRPLogic.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   218
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Test"
         Height          =   255
         Left            =   9840
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtMrp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "mps name"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   2670
         TabIndex        =   33
         Tag             =   "PO"
         Top             =   210
         Width           =   3030
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Generate"
         Height          =   255
         Left            =   9960
         TabIndex        =   31
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox List1 
         Appearance      =   0  'Flat
         DataField       =   "Plan Horizon"
         Height          =   315
         ItemData        =   "FrmMRPLogic.frx":6BDC
         Left            =   6450
         List            =   "FrmMRPLogic.frx":6BE9
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "PO"
         Top             =   855
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
         TabIndex        =   15
         Tag             =   "PO"
         Top             =   1845
         Width           =   2535
      End
      Begin VB.TextBox txtMrp 
         Appearance      =   0  'Flat
         DataField       =   "NoItem"
         Height          =   330
         Index           =   0
         Left            =   2670
         TabIndex        =   14
         Tag             =   "PO"
         Text            =   "FG/000000000005"
         Top             =   555
         Width           =   1470
      End
      Begin VB.TextBox txtMrp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "QTY"
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
         TabIndex        =   13
         Tag             =   "PO"
         Text            =   "50"
         Top             =   915
         Width           =   1470
      End
      Begin VB.TextBox txtMrp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Unit Aloc"
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
         TabIndex        =   12
         Tag             =   "PO"
         Text            =   "20"
         Top             =   1215
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
         TabIndex        =   11
         Tag             =   "PO"
         Text            =   "100"
         Top             =   1515
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
         TabIndex        =   10
         Tag             =   "PO"
         Text            =   "3"
         Top             =   1515
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
         Left            =   9495
         TabIndex        =   9
         Tag             =   "PO"
         Text            =   "98"
         Top             =   1485
         Width           =   1470
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
         TabIndex        =   8
         Tag             =   "PO"
         Text            =   "300"
         Top             =   1815
         Width           =   1470
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         Caption         =   "Planning Horizon                (                   )"
         DataField       =   "Order Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4515
         TabIndex        =   7
         Tag             =   "PO"
         Top             =   900
         Width           =   3915
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
         Left            =   9495
         TabIndex        =   5
         Tag             =   "PO"
         Text            =   "50"
         Top             =   1785
         Width           =   1470
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4140
         Picture         =   "FrmMRPLogic.frx":6C01
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   563
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DgJadwal 
         Height          =   4035
         Left            =   60
         TabIndex        =   2
         Top             =   2490
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   7117
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
         DataField       =   "Require Date"
         Height          =   285
         Index           =   0
         Left            =   6450
         TabIndex        =   6
         Tag             =   "PO"
         Top             =   1185
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   71630851
         CurrentDate     =   38624
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "End Date"
         Height          =   285
         Index           =   1
         Left            =   9480
         TabIndex        =   16
         Tag             =   "PO"
         Top             =   1170
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   71630851
         CurrentDate     =   38715
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   360
         X2              =   2700
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPS Name"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   32
         Top             =   210
         Width           =   750
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Identifier"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   30
         Top             =   615
         Width           =   1035
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Inventory"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units Allocated"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   28
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot size (Use 1 for L4L)"
         Height          =   195
         Index           =   3
         Left            =   345
         TabIndex        =   27
         Top             =   1560
         Width           =   1650
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lead Time in Weeks                                                   Yield percentage"
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   26
         Top             =   1560
         Width           =   4905
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Order                                                                  Safety Stock"
         Height          =   195
         Index           =   5
         Left            =   4560
         TabIndex        =   25
         Top             =   1860
         Width           =   4620
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   2700
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   345
         X2              =   2685
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   330
         X2              =   2670
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   330
         X2              =   2670
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   4575
         X2              =   6915
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4575
         X2              =   6915
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4545
         X2              =   6885
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Require Date                                                             Entry Date"
         Height          =   195
         Index           =   7
         Left            =   4560
         TabIndex        =   24
         Top             =   1230
         Width           =   4470
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   8250
         X2              =   10590
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   8250
         X2              =   10590
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   8250
         X2              =   10590
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label LblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Make To Order"
         Height          =   195
         Left            =   8505
         TabIndex        =   23
         Top             =   915
         Width           =   1065
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Travel Mate"
         DataField       =   "ItemName"
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
         Index           =   6
         Left            =   4605
         TabIndex        =   22
         Tag             =   "PO"
         Top             =   615
         Width           =   1020
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4545
         X2              =   6885
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Warning"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   21
         Top             =   6585
         Width           =   1275
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lead Time"
         Height          =   195
         Index           =   9
         Left            =   690
         TabIndex        =   20
         Top             =   6870
         Width           =   720
      End
      Begin VB.Label LblMrp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Stock"
         Height          =   195
         Index           =   10
         Left            =   495
         TabIndex        =   19
         Top             =   7125
         Width           =   915
      End
      Begin VB.Label LblStock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Index           =   0
         Left            =   1725
         TabIndex        =   18
         Top             =   6870
         Width           =   60
      End
      Begin VB.Label LblStock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Index           =   1
         Left            =   1725
         TabIndex        =   17
         Top             =   7125
         Width           =   60
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   7500
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmMRPLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcJadwal As New Recordset
Private RcPartner As New DBQuick
Private RcMPS As New Recordset
Private RcSchedule As New Recordset

Private WithEvents RcLoad As frmCaller
Attribute RcLoad.VB_VarHelpID = -1
Private mAwal As Integer
Private mAkhir As Integer
Private mCount As Integer
Private mTotal As Integer
Dim vColindex As Integer
Private mList As String
Private mRowLast As Long

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub Command1_Click()
OpenCucakRowo IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxx"), False
insert_FCQTY txtMrp(0), List1.Text  'digunakan untuk nampilkan gross yang dari MPS
insert_Schedule txtMrp(0) 'digunakan untuk nampil schedule recepit dari PO QTY Purchasing
refresh_grid
End Sub

Private Sub Command2_Click()
''insert_FCQTY txtMrp(0)
'
'Text1.Text = WeekNumber(DTPicker1(0).value) + 1
'insert_Schedule txtMrp(0)
'Text2.Text = DgJadwal.Columns(4).Caption
End Sub



Private Sub DgJadwal_DblClick()
'Text1.Text = DgJadwal.RowBookmark(DgJadwal.Row)
'Text1.Text = DgJadwal.Row & " y  " & RowSetAbsolute(5)
End Sub

Private Sub Form_Load()
Set RcLoad = New frmCaller
HiasFormManTell Picture2, Me
List1.ListIndex = 0
With MyDDE
     .EditModeReplace = False
     Set .BindForm = FrmMRPLogic
     .BindFormTAG = "PO"
     Set .ActiveConnection = CNN
      .PrepareQuery = " SELECT [MRP INVENTORY].[No ID], [MRP INVENTORY].NoItem, [MRP INVENTORY].[QTY Order], [MRP INVENTORY].[Lot Size], [MRP INVENTORY].[Multiple Lot], " & _
                      " [MRP INVENTORY].[Plan Horizon], [MRP INVENTORY].[Order Type], [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date]," & _
                      " [MRP INVENTORY].[Lead Time], [MRP INVENTORY].[Yield Percentage], [MRP INVENTORY].[Safety Stock]," & _
                      " ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY, [MRP INVENTORY].[Unit Aloc], Inventory.ItemName, [MRP INVENTORY].[MPS name]" & _
                      " FROM [MRP INVENTORY] INNER JOIN" & _
                      " Inventory ON [MRP INVENTORY].NoItem = Inventory.NoItem LEFT OUTER JOIN" & _
                      " [Inventory Tabel] ON [MRP INVENTORY].NoItem = [Inventory Tabel].NoItem" & _
                      " GROUP BY [MRP INVENTORY].[No ID], [MRP INVENTORY].NoItem, [MRP INVENTORY].[QTY Order], [MRP INVENTORY].[Lot Size]," & _
                      " [MRP INVENTORY].[Plan Horizon], [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date], [MRP INVENTORY].[Lead Time]," & _
                      " [MRP INVENTORY].[Yield Percentage], [MRP INVENTORY].[Safety Stock], [MRP INVENTORY].[Multiple Lot], [MRP INVENTORY].[Order Type]," & _
                      " ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0), [MRP INVENTORY].[Unit Aloc], Inventory.ItemName, [MRP INVENTORY].[MPS name]"
'MessageBox .PrepareQuery
End With
End Sub

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

Private Sub SaveLayout()
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim mNo As Integer
Dim RcTes As New Recordset
Dim StrKode As String
StrKode = ""
Set RcTes = RcJadwal.Clone(adLockReadOnly)
If SendDataToServer("Delete From [Mrp Detail] Where NoItem =N'" & txtMrp(0) & "'") = True Then
   For j = 1 To RcTes.Recordcount
       If k = 6 Then k = 0
       RcTes.AbsolutePosition = j
       If RcTes.Fields(0) <> "" Then StrKode = RcTes.Fields(0)
       If StrKode = "" Then StrKode = RcTes.Fields(0)
       mNo = 1
       k = k + 1
       For I = mAwal + 1 To mAwal + mCount + 1
           mNo = mNo + 1
           If mNo = 2 Then
              SendDataToServer (" INSERT INTO [MRP Detail] " & _
                                " (NoItem, Component,[Time Days],[List Value1],[No Urut]) " & _
                                " VALUES (N'" & txtMrp(0) & "', N'" & StrKode & "',0," & CDbl(IIf(Not IsNull(RcTes.Fields(mNo)), RcTes.Fields(mNo), 0)) & "," & k & ") ")
           Else
              SendDataToServer (" INSERT INTO [MRP Detail] " & _
                                " (NoItem, Component,[Time Days],[List Value1],[No Urut]) " & _
                                " VALUES (N'" & txtMrp(0) & "', N'" & StrKode & "'," & I - 1 & "," & CDbl(IIf(Not IsNull(RcTes.Fields(mNo)), RcTes.Fields(mNo), 0)) & "," & k & ") ")
           End If
       Next I
   Next j
End If
End Sub

Private Sub DgJadwal_AfterColEdit(ByVal ColIndex As Integer)
Dim Rc As New DBQuick
mRowLast = 0
If ColIndex > 2 Then
With RcJadwal
     vColindex = ColIndex
     If DgJadwal.Columns(ColIndex).Value <> "" Then
        Rc.DBOpen "SELECT [BOM Component Detail].Component, Inventory.ItemName, Inventory.LeadTimeDays AS LeadTimeDays, ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS Qty, Inventory.MinStock, [BOM Component Detail].QTYUsage FROM [BOM Component Detail] INNER JOIN Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE     ([BOM Component Detail].NoItem = N'" & MyDDE.GetFieldByName("NoItem") & "') ORDER BY [BOM Component Detail].Component", CNN, lckLockReadOnly
        With Rc.DBRecordset
             If .Recordcount <> 0 Then
                .MoveFirst
                Do
                   If .EOF = True Then Exit Do
                   If .AbsolutePosition = 1 Then
                      LoadJadwal .Fields(0), IIf(Not IsNull(MyDDE.GetFieldByName("Lead Time")), IIf(Not IsNull(MyDDE.GetFieldByName("Lead Time")), MyDDE.GetFieldByName("Lead Time"), 0), 0), .Fields(3), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), False, , ColIndex, .Fields(1)
                   Else
'                      LoadJadwal .Fields(0), .Fields(2), .Fields(3), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), False, .Fields(5), ColIndex - .Fields(2), .Fields(1)
                   End If
                   mRowLast = mRowLast + 6
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
If cmdLink(0).Enabled = True Then
    If DgJadwal.col <= 2 Then
       DgJadwal.MarqueeStyle = dbgHighlightRow
       DgJadwal.AllowUpdate = False
    ElseIf (DgJadwal.col > 2) Then
       If RcJadwal.AbsolutePosition = 1 Or RcJadwal.AbsolutePosition = 2 Then
          DgJadwal.MarqueeStyle = dbgFloatingEditor
          DgJadwal.AllowUpdate = True
       Else
          DgJadwal.MarqueeStyle = dbgHighlightRow
          DgJadwal.AllowUpdate = False
       End If
    End If
End If
End Sub

Private Sub DTPicker1_Change(Index As Integer)
Dim vWeek As Integer
Dim mlistPlan As Integer

'DTPicker1_Click Index
mTotal = CDbl(txtMrp(6))
vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
mList = IIf(Not IsNull(MyDDE.GetFieldByName("Plan Horizon")), MyDDE.GetFieldByName("Plan Horizon"), "Week")
mlistPlan = HitungHari
If mlistPlan = 0 Then
   mlistPlan = DTPicker1(0).Value - DTPicker1(1).Value
End If
Select Case mList
       Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), mlistPlan
       Case "Week": GenerateJadwal vWeek, mlistPlan
       Case Else: GenerateJadwal vWeek, mlistPlan
End Select
End Sub

Private Sub DTPicker1_Click(Index As Integer)
If DTPicker1(1).Value > DTPicker1(0).Value Then
   DTPicker1(1).Value = DTPicker1(0).Value
End If
'HitungHari
End Sub

Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub GenerateJadwal(ByVal vAwal As Integer, ByVal vAkhir As Integer)
On Error Resume Next
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
mCount = vAkhir
GridLayout
Set MyDDE.ChildRecordset = RcJadwal
Err.Clear
End Sub

Private Sub GridLayout()
Dim I As Integer
If DgJadwal.Splits.Count < 2 Then DgJadwal.Splits.Add 1
DgJadwal.Splits(0).Columns(0).Visible = True
DgJadwal.Splits(0).Columns(1).Visible = True
DgJadwal.Splits(0).Columns(1).width = 1750
DgJadwal.Splits(1).Columns(0).Visible = False
DgJadwal.Splits(1).Columns(1).Visible = False
DgJadwal.Splits(0).ScrollBars = dbgHorizontal
DgJadwal.Splits(1).ScrollBars = dbgBoth
DgJadwal.Splits(1).RecordSelectors = False
DgJadwal.Splits(0).AllowSizing = False
DgJadwal.Splits(1).Size = 2
DgJadwal.Splits(1).SizeMode = dbgScalable
DgJadwal.Splits(1).LeftCol = 0
For I = 0 To RcJadwal.Fields.Count - 1
    If I <= 1 Then
       DgJadwal.Columns(I).width = 1700
       DgJadwal.Columns(I).Alignment = dbgLeft
    Else
       DgJadwal.Columns(I).width = 600
       DgJadwal.Columns(I).Alignment = dbgRight
       DgJadwal.Columns(I).NumberFormat = "#,##0;(#,##0)"
       DgJadwal.Splits(0).Columns(I).Visible = False
       DgJadwal.Splits(1).Columns(I).AllowSizing = True
    End If
    DgJadwal.Columns(I).DividerStyle = dbgRaised
Next
DgJadwal.TabAcrossSplits = True
DgJadwal.Refresh
End Sub

Private Sub OpenCucakRowo(ByVal vKode As String, Optional ByVal Tipical As Boolean = False)
On Error Resume Next
Dim Rc As New DBQuick
Dim RcDetail As New Recordset
Dim vWeek As Integer
Dim I As Integer
Dim iJ As Integer
Dim mLast As Integer
Dim Avdata As Variant
'Dim mStart As Boolean
mRowLast = 0
mTotal = CDbl(txtMrp(6))
vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
mList = IIf(Not IsNull(MyDDE.GetFieldByName("Plan Horizon")), MyDDE.GetFieldByName("Plan Horizon"), "Week")
Select Case MyDDE.GetFieldByName("Plan Horizon")
       Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), HitungHari
       Case "Week": GenerateJadwal vWeek, HitungHari
'       Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
       Case Else: GenerateJadwal vWeek, HitungHari
End Select
'Tolong Jarno ben sak garis wae, debug-e dhek sql cek enak....Please.
If Tipical = False Then
   Rc.DBOpen " SELECT [BOM Component Detail].Component, Inventory.ItemName, Inventory.LeadTimeDays AS LeadTimeDays,  ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY, [BOM Component Detail].QTYUsage FROM  [BOM Component Detail] INNER JOIN  Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN                       [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE  ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY [BOM Component Detail].Component", CNN, lckLockReadOnly
   With Rc.DBRecordset
        If .Recordcount <> 0 Then
           'LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), MyDDE.GetFieldByName("QTY") - MyDDE.GetFieldByName("Unit Aloc"), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
           LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), txtMrp(1) - txtMrp(2), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
           Do
              If .EOF = True Then Exit Do
                 LoadJadwal .Fields(0), .Fields(2), .Fields(3), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True '.Fields(3)
                 mRowLast = mRowLast + 6
                .MoveNext
           Loop
        End If
   End With
  
Else
  Rc.DBOpen "Shape {SELECT [MRP Detail].Component, [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date] FROM [MRP INVENTORY] INNER JOIN [MRP Detail] ON [MRP INVENTORY].NoItem = [MRP Detail].NoItem WHERE ([MRP INVENTORY].NoItem = N'" & vKode & "') GROUP BY [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date], [MRP Detail].Component ORDER BY [MRP Detail].Component} Append({SELECT [MRP Detail].Component AS Component, [MRP Detail].[Time Days] AS [Plan Horizon], [MRP Detail].[List Value1] AS Amount, [MRP Detail].[No Urut] FROM [MRP Detail] INNER JOIN Inventory ON [MRP Detail].NoItem = Inventory.NoItem WHERE     (Inventory.Manufacture = 1) AND ([MRP Detail].NoItem = N'" & vKode & "') GROUP BY [MRP Detail].[Time Days], [MRP Detail].[List Value1], [MRP Detail].[No Urut], [MRP Detail].Component ORDER BY [MRP Detail].Component, [MRP Detail].[No Urut]} As ChildMD Relate Component To Component)", CNN, lckLockBatch
  With Rc.DBRecordset
       If .Recordcount <> 0 Then
'          vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
'          mList = MyDDE.GetFieldByName("Plan Horizon")
'          Select Case MyDDE.GetFieldByName("Plan Horizon")
'                 Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), HitungHari
'                 Case "Week": GenerateJadwal vWeek, HitungHari
''                 Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
'                 Case Else: GenerateJadwal vWeek, HitungHari
'          End Select
          Set RcDetail = Rc.DBRecordset("ChildMD").UnderlyingValue
'            RcDetail.MoveFirst
'            MsgBox RcDetail.GetString(adClipString)
'             LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), MyDDE.GetFieldByName("QTY") - MyDDE.GetFieldByName("Unit Alloc"), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
             Do
               If .EOF Then Exit Do
                  mLast = 1
                  iJ = 0
                  If RcDetail.Recordcount <> 0 Then
                        Avdata = RcDetail.Getrows(RcDetail.Recordcount, adBookmarkFirst)
                        For I = 0 To UBound(Avdata, 2)
                            iJ = iJ + 1
                            If mLast <> Avdata(3, I) Then iJ = 1
                            Select Case Avdata(3, I)
                                   Case 1:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(0) = .Fields(0): RcJadwal.Fields(1) = "Gross Requirement"
                                   Case 2:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Schedule Receipt"
                                   Case 3:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "On Hand"
                                   Case 4:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Net Requirement"
                                   Case 5:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Plan Order Receipt"
                                   Case 6:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Plan Order Release"
                            End Select
'                            MsgBox RcJadwal.Fields(10).Name
                            RcJadwal.Fields((2 + iJ) - 1) = Avdata(2, I)
                            'RcJadwal.Fields(iJ - 1) = Avdata(2, I)
                            mLast = Avdata(3, I)
                        Next I
                  End If
               .MoveNext
             Loop
             .MoveFirst
       Else
       End If
  End With
End If
End Sub

Private Sub LoadJadwal(ByVal vBOMItemLoad As String, ByVal vLeadTime As Integer, ByVal vOnHand As Double, ByVal vSafetyStock As Double, Optional ByVal Update_OR_New As Boolean = False, Optional ByVal vQTyUsage As Long, Optional ByVal vLastColindex As Integer, Optional ByVal vItemName As String)
On Error Resume Next
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim vSc As Long
Dim vPoRc As Long
Dim vGr As Long
Dim vOh As Long
Dim vTot As Long
Dim vPoRR As Long
Dim mTmpRow As Long
With RcJadwal
'     For I = 0 To 5
         'Prepare Field
         If Update_OR_New = True Then .AddNew 0, vBOMItemLoad Else .AbsolutePosition = RowSetAbsolute(1)
         .Fields(1) = "Gross Requirement"
         If Update_OR_New = True Then .AddNew Else .AbsolutePosition = RowSetAbsolute(2)
         .Fields(1) = "Schedule Receipt"
         If mRowLast >= 6 Then
            DgJadwal.row = (mRowLast) - 1
            vPoRR = Val(DgJadwal.Columns(mAkhir))
            DgJadwal.row = DgJadwal.row + 1
            '.Fields(mAkhir) = vPoRR * vQTyUsage
            vLastColindex = mAkhir
            .AbsolutePosition = RowSetAbsolute(2)
         End If
        ' .Fields(1) = "Schedule Receipt" 'lama
        ' '.Fields(0) = vItemName
         If Update_OR_New = True Then .AddNew Else .AbsolutePosition = RowSetAbsolute(3)
         .Fields(1) = "On Hand"
         If Update_OR_New = True Then .Fields(2) = (vOnHand)
         'Rumus On Hand
         If Update_OR_New = True Then .AddNew Else .AbsolutePosition = RowSetAbsolute(4)
         .Fields(1) = "Net Requirement"
         If Update_OR_New = True Then .AddNew Else .AbsolutePosition = RowSetAbsolute(5)
         .Fields(1) = "Plan Order Receipt"
         If Update_OR_New = True Then .AddNew Else .AbsolutePosition = RowSetAbsolute(6)
         .Fields(1) = "Plan Order Release"
         .MoveFirst
         If Update_OR_New = False Then mRowLast = 0
         For I = 1 To .Recordcount
             .AbsolutePosition = I
              For j = 3 To .Fields.Count - 1
              'Gros Req
                 If (I Mod 6) = 0 Then
                    If (I >= 6) Then
                       DgJadwal.row = (mRowLast + 6) - 1
                       vPoRR = CDbl(IIf(DgJadwal.Columns(j) <> "", DgJadwal.Columns(j), 0))
                       'MsgBox .AbsolutePosition Mod 6
                       '.AbsolutePosition = I + 1
                       DgJadwal.row = DgJadwal.row + 1
                       If mAkhir - vLeadTime <= 2 Then
                          GoTo Err_HorizonDays
                          Exit For
                       End If
                       .Fields(j).Value = (vPoRR * 2)
'                       If .Fields(j).Value <> 0 Then
'                          MsgBox "dfl;gklsdjfgks;k"
'                       End If
'                       .AbsolutePosition = .AbsolutePosition + 1
                    End If
                 End If
                 
                 'End Gros Req
                 'On Hand
                 vOh = MPSFormula(3, j - 1)
                 vSc = MPSFormula(2, j)
                 'vPoRc = MPSFormula(5, j - 1) 'lama
                 vPoRc = MPSFormula(5, j)
                 vGr = MPSFormula(1, j)
                 .AbsolutePosition = RowSetAbsolute(3)
                 .Fields(j).Value = (vOh + vSc + vPoRc) - vGr
                 'End Hand
                 'Net Requirement
                 'vOh = MPSFormula(3, j) 'lama
                 vOh = MPSFormula(3, j - 1)
                 vSc = MPSFormula(2, j)
                 vPoRc = MPSFormula(5, j)
                 vGr = MPSFormula(1, j)
                 .AbsolutePosition = RowSetAbsolute(4)
                 .Fields(j).Value = IIf(vGr = 0, 0, IIf((((vGr + vSafetyStock) - vSc) - vOh) < 0, 0, (((vGr + vSafetyStock) - vSc) - vOh)))
                 'End Net
                 'PORC
                  vPoRc = MPSFormula(4, j)
                 .AbsolutePosition = RowSetAbsolute(5)
                 '.Fields(j).value = RoundUp(vPoRc, MyDDE.GetFieldByName("Lot Size"))
                 .Fields(j).Value = (vPoRc * MyDDE.GetFieldByName("Lot Size"))
                 'End PORC
                 'PORR
                 vPoRR = MPSFormula(5, vLastColindex)
                 .AbsolutePosition = RowSetAbsolute(6)
                 .Fields(vLastColindex - vLeadTime).Value = (vPoRR / (Val(txtMrp(5)) / 100))
                 mAkhir = vLastColindex - vLeadTime
             Next j
'             .MoveNext
            If Update_OR_New = False Then If (I Mod 6) = 0 Then mRowLast = mRowLast + 6
         Next I
         .MoveFirst
End With
DgJadwal.Refresh
Err.Clear
Exit Sub
Err_HorizonDays:
    LblStock(0) = "Kolom Horizon Plan tidak mencukupi untuk Plan Order Release."
    
Err_SafetyStock:
    LblStock(1) = "Safety stock lebih besar dari On Hand"

End Sub

Private Function MPSFormula(ByVal vMpsRows As Integer, ByVal vMpsCol As Integer) As Variant
Select Case vMpsRows
       Case 1:
            DgJadwal.row = (vMpsRows + mRowLast) - 1
            MPSFormula = CDbl(IIf(DgJadwal.Columns(vMpsCol) <> "", DgJadwal.Columns(vMpsCol), 0))
       Case 2:
            DgJadwal.row = (vMpsRows + mRowLast) - 1
            MPSFormula = CDbl(IIf(DgJadwal.Columns(vMpsCol) <> "", DgJadwal.Columns(vMpsCol), 0))
       Case 3:
            DgJadwal.row = (vMpsRows + mRowLast) - 1
            MPSFormula = CDbl(IIf(DgJadwal.Columns(vMpsCol) <> "", DgJadwal.Columns(vMpsCol), 0))
       Case 4:
            DgJadwal.row = (vMpsRows + mRowLast) - 1
            MPSFormula = CDbl(IIf(DgJadwal.Columns(vMpsCol) <> "", DgJadwal.Columns(vMpsCol), 0))
       Case 5:
            DgJadwal.row = (vMpsRows + mRowLast) - 1
            MPSFormula = CDbl(IIf(DgJadwal.Columns(vMpsCol) <> "", DgJadwal.Columns(vMpsCol), 0))
End Select
If IsNull(MPSFormula) Or IsEmpty(MPSFormula) Or MPSFormula = "" Or MPSFormula = "0" Then MPSFormula = 0
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RcJadwal.Close
Set RcJadwal = Nothing
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMPSNew = Nothing
End Sub

Private Sub List1_Click()
If List1.Text <> "" Then mList = List1.Text
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
'            txtBox(0).Enabled = False
             cmdLink(0).Enabled = True
             cmdLink(1).Enabled = True
       Case tmbAddNew:
            
            With MyDDE
                 .GetFieldByName("Lead Time") = 0
                 .GetFieldByName("QTY") = 0
                 .GetFieldByName("Unit Aloc") = 0
                 .GetFieldByName("Safety Stock") = 0
'                 .GetFieldByName("UOM") = "KG"
'                 .GetFieldByName("MutasiDate") = dDateBegin
            End With
'            txtBox(0).SetFocus
'            mVarAdd = True
       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               If OpenPartner = True Then CancelDetailTrans
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
       Case tmbDelete:
'            If MyDDE.IsChildMemberReady = True Then
'               SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & txtBox(0) & "') ")
'            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SaveLayout
            End If
       Case tmbPrint:
       CallRPTReport "MPS Report.rpt", "Select * from [MPS Report] Where [Item ID] = '" & MyDDE.GetFieldByName("NoItem") & "'"
'            CallRPTReport "Raw Material.Rpt", "Select * from [Raw Material] Where [Kode RM]='" & txtBox(0) & "'"
'       Case Else: 'mVarDataDc = False
End Select
cmdLink(0).Enabled = txtMrp(0).Enabled
cmdLink(1).Enabled = txtMrp(0).Enabled
GridLayout
DgJadwal.Splits(1).Columns(0).Visible = False
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            DgJadwal.AllowUpdate = True
       Case tmbAddNew:
            cmdLink(0).Enabled = True
            DgJadwal.AllowUpdate = True
'            With MyDDE
'                 .GetFieldByName("MutasiID") = MyData.PrepareIndex(tmbTransaksiMutasiPenjualan, 5, cboRakit(1).BoundText, cboRakit(1).BoundText & "/")
'                 .GetFieldByName("Notes") = "-"
'                 .GetFieldByName("UOM") = "KG"
'                 .GetFieldByName("MutasiDate") = dDateBegin
'            End With
'            txtBox(0).SetFocus
'            mVarAdd = True
       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               If OpenPartner = True Then CancelDetailTrans
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
       Case tmbDelete:
'            If MyDDE.IsChildMemberReady = True Then
'               SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & txtBox(0) & "') ")
'            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  'MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                  'PrepareQuery
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
            DgJadwal.AllowUpdate = False
       Case tmbPrint:
'            CallRPTReport "Raw Material.Rpt", "Select * from [Raw Material] Where [Kode RM]='" & txtBox(0) & "'"
''       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenCucakRowo IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxx"), True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub RcLoad_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case RcLoad.FromTagActive
       Case "Bom List":
            MyDDE.GetFieldByName("NoItem") = RcLoad.GetFieldByName(0)
            MyDDE.GetFieldByName("ItemName") = RcLoad.GetFieldByName(1)
            MyDDE.GetFieldByName("Plan Horizon") = "Week"
            MyDDE.GetFieldByName("Require Date") = DTPicker1(0).Value
            MyDDE.GetFieldByName("End Date") = DTPicker1(1).Value
        Case "MPS List":
            txtMrp(8) = RcLoad.GetFieldByName(0)
End Select
End Sub

Private Sub txtMrp_Change(Index As Integer)
If Index <> 8 Then
    If Index <> 0 Then If txtMrp(Index) = "" Then txtMrp(Index) = 0
End If
End Sub

Private Sub txtMrp_GotFocus(Index As Integer)
Block txtMrp(Index)
End Sub

Private Sub txtMrp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub txtMrp_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If Index <> 0 Then ValidNum KeyAscii
End Sub

Private Function HitungHari() As Long
Dim mTotal As Long
mTotal = CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)) - CDate(IIf(Not IsNull(MyDDE.GetFieldByName("End Date")), MyDDE.GetFieldByName("End Date"), Date))
Select Case mList
       Case "Day": HitungHari = mTotal
       Case "Week": HitungHari = Round(mTotal / 7)
       Case "Monthly": HitungHari = Round(mTotal / 30)
End Select
End Function

Private Sub PrepareQuery()
With MyDDE
     .PrepareAppend = " INSERT INTO [MRP INVENTORY] " & _
                      " (NoItem, [QTY Order], [Lot Size], [Multiple Lot], [Plan Horizon], [Order Type], [Require Date], [End Date], [Lead Time], [Yield Percentage], [Safety Stock],[MPS name])" & _
                      " VALUES (N'" & txtMrp(0) & "', " & CDbl(txtMrp(6)) & ", " & CDbl(txtMrp(3)) & ", " & BoolToInt(Check1.Value) & ", N'" & mList & "', " & BoolToInt(Check2.Value) & ", CONVERT(DATETIME, '" & Format(DTPicker1(0), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(DTPicker1(1), "dd/mm/yy") & "', 3), " & CDbl(txtMrp(4)) & ", " & CDbl(txtMrp(5)) & ", " & CDbl(txtMrp(7)) & ",'" & txtMrp(8) & "')"
     .PrepareUpdate = " UPDATE [MRP INVENTORY]" & _
                      " SET [QTY Order] = " & CDbl(txtMrp(6)) & ", [Lot Size] = " & CDbl(txtMrp(3)) & ", [Multiple Lot] = " & BoolToInt(Check1.Value) & ", [Plan Horizon] = N'" & List1.Text & "', [Order Type] = " & BoolToInt(Check2.Value) & ", [Require Date] = CONVERT(DATETIME, '" & Format(DTPicker1(0).Value, "dd/mm/yy") & "'," & _
                      " 3), [End Date] = CONVERT(DATETIME, '" & Format(DTPicker1(1).Value, "dd/mm/yy") & "', 3), [Lead Time] = " & CDbl(txtMrp(4)) & ", [Yield Percentage] = " & CDbl(txtMrp(5)) & ", [Safety Stock] = " & CDbl(txtMrp(7)) & ",[mps name]='" & txtMrp(8) & "'  WHERE  (NoItem = N'" & MyDDE.GetFieldByName("NoItem") & "')"
     .PrepareDelete = " Delete From [MRP INVENTORY] where NoItem='" & txtMrp(0) & "'"
End With
End Sub

Private Function RoundUp(ByVal TotalAmount As Double, ByVal AmountMultiPlier As Variant) As Double
   On Error Resume Next
   Dim vTmpVal As Double
   vTmpVal = TotalAmount / AmountMultiPlier
   If Int(vTmpVal) = vTmpVal Then
      RoundUp = TotalAmount
   Else
      vTmpVal = Int(vTmpVal) + 1
      RoundUp = vTmpVal * AmountMultiPlier
   End If
   Err.Clear
End Function

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
             If txtMrp(8) = "" Then
                'gunakan apabila tidak pakai MPS
                RcPartner.DBOpen "SELECT NoItem AS [BOM Id], ItemName AS Keterangan, UOM AS UOM FROM Inventory WHERE     (Manufacture = 1) ORDER BY NoItem", CNN, lckLockReadOnly
             Else
                'di gunakan apabila pakai MPS
                RcPartner.DBOpen "SELECT  [MPS Detail].NoItem, Inventory.ItemName, Inventory.UOM, [MPS Header].No_MPS " & _
                                 "FROM    [MPS Header] INNER JOIN " & _
                                 "[MPS Detail] ON [MPS Header].No_MPS = [MPS Detail].No_MPS INNER JOIN " & _
                                 "Inventory ON [MPS Detail].NoItem = Inventory.NoItem " & _
                                 "WHERE ([MPS Header].No_MPS = '" & txtMrp(8) & "')" & _
                                 " GROUP BY [MPS Detail].NoItem, Inventory.ItemName, Inventory.UOM, [MPS Header].No_MPS", CNN, lckLockReadOnly
             End If
       Case 1:
            RcPartner.DBOpen "SELECT No_MPS, Description, Periode_no, Periode_type,[Require Date] ,[End Date] FROM [MPS Header] order by no_mps", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            RcLoad.FromTagActive = "Bom List"
          Case 1:
            RcLoad.FromTagActive = "MPS List"
   End Select
   Set RcLoad.FormData = RcPartner.DBRecordset
   RcLoad.LookUp Me
 
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Function RowSetAbsolute(ByVal vRowActive As Integer) As Long
If mRowLast = 0 Then
  RowSetAbsolute = vRowActive
Else
  RowSetAbsolute = mRowLast + vRowActive
End If
End Function


Private Sub insert_FCQTY(VnoItem As String, Vperiode As String)
On Error Resume Next
Dim Irow, jCol As Integer

Set RcMPS = New Recordset
'digunakan untuk ambil nilai fcqty di tabel MPS
'RcMPS.Open "SELECT     dbo.[MPS Header].No_MPS, dbo.[MPS Header].Description, dbo.[MPS Header].Periode_no, dbo.[MPS Header].Periode_type," & _
'           "dbo.[MPS Header].[Require Date], dbo.[MPS Header].[End Date], dbo.[MPS Detail].NoItem, dbo.[MPS Detail].fcast_item, dbo.[MPS Detail].Time_Days," & _
'           "dbo.[MPS Detail].list_value1 , dbo.[MPS Detail].no_urut " & _
'           " FROM dbo.[MPS Header] INNER JOIN " & _
'           " dbo.[MPS Detail] ON dbo.[MPS Header].No_MPS = dbo.[MPS Detail].No_MPS " & _
'           " WHERE (dbo.[MPS Detail].NoItem = N'" & VnoItem & "') AND (dbo.[MPS Detail].fcast_item = 'FCQTY') AND (dbo.[MPS Header].Periode_type = '" & Vperiode & "')", CNN, adOpenKeyset, adLockReadOnly
           
RcMPS.Open "SELECT  [MPS Header].No_MPS, [MPS Header].Description, [MPS Header].Periode_no, [MPS Header].Periode_type," & _
           "[MPS Header].[Require Date],[MPS Header].[End Date], [MPS Detail].NoItem, [MPS Detail].fcast_item, [MPS Detail].Time_Days," & _
           "[MPS Detail].list_value1 , [MPS Detail].no_urut " & _
           " FROM [MPS Header] INNER JOIN " & _
           " [MPS Detail] ON [MPS Header].No_MPS = [MPS Detail].No_MPS " & _
           " WHERE ([MPS Detail].NoItem = N'" & VnoItem & "') AND ([MPS Detail].fcast_item = 'FCQTY') AND ([MPS Header].Periode_type = '" & Vperiode & "') and (([MPS Header].[Require Date]>='" & Format(DTPicker1(1).Value, "mm/dd/yyyy") & "') and ([MPS Header].[end Date]<='" & Format(DTPicker1(0).Value, "mm/dd/yyyy") & "'))", CNN, adOpenKeyset, adLockReadOnly


txtMrp(8) = RcMPS.Fields("no_mps")
If RcMPS.Recordcount > 0 Then
    Irow = 0
    RcJadwal.MoveFirst
    For Irow = 0 To 5 '(mRowLast + 5) 'set baris
        jCol = 3
            RcMPS.MoveFirst
            Do While jCol <= (mCount + 2) Or RcMPS.EOF <> True
                DgJadwal.row = Irow 'DgJadwal.RowBookmark(DgJadwal.Row)
                DgJadwal.col = jCol
                DgJadwal.Columns(jCol).Value = RcMPS.Fields("list_value1")
                jCol = jCol + 1
                RcMPS.MoveNext
            Loop
           RcJadwal.MoveNext
           Irow = Irow + 5
    Next Irow
Else
    MessageBox "Data FCQTY di MPS Tidak Ada Untuk Gross", "FrmMRPLogic", msgOkOnly, msgExclamation
End If

'RcJadwal.MoveFirst   'digunakan untuk semua row
'For Irow = 0 To (mRowLast + 5) 'set baris
'    jCol = 3
'        RcMPS.MoveFirst
'        Do While jCol <= (mCount + 2) Or RcMPS.EOF <> True
'            DgJadwal.Row = Irow 'DgJadwal.RowBookmark(DgJadwal.Row)
'            DgJadwal.Col = jCol
'            DgJadwal.Columns(jCol).value = RcMPS.Fields("list_value1")
'            jCol = jCol + 1
'            RcMPS.MoveNext
'        Loop
'       Text1.Text = DgJadwal.RowBookmark(DgJadwal.Row)
'       RcJadwal.MoveNext
'       Irow = Irow + 5
'Next Irow
End Sub

Private Sub insert_Schedule(VnoItem As String)
On Error Resume Next
Dim Irow, jCol As Integer

'query ini di gunakan untuk mengambil nilai yg ada di po detail
Set RcSchedule = New Recordset
RcSchedule.Open "SELECT dbo.[PO Order].PurchaseID, dbo.[Detail PO].NoItem, dbo.[Detail PO].QTYPO, dbo.[Detail PO].ScheduleDate, dbo.[PO Order].Status," & _
                  "SUM(dbo.[Detail PO].QTYPO) As sumQTY " & _
                  "FROM dbo.[PO Order] INNER JOIN " & _
                  "dbo.[Detail PO] ON dbo.[PO Order].PurchaseID = dbo.[Detail PO].PurchaseID " & _
                  "WHERE (LEFT(dbo.[PO Order].PurchaseID, 2) = 'PO') AND (dbo.[PO Order].Status = 0 OR " & _
                  " dbo.[PO Order].Status = 2) AND (dbo.[Detail PO].NoItem = N'" & VnoItem & "')" & _
                  " GROUP BY dbo.[PO Order].PurchaseID, dbo.[Detail PO].NoItem, dbo.[Detail PO].QTYPO, dbo.[Detail PO].ScheduleDate, dbo.[PO Order].Status ", CNN, adOpenKeyset, adLockReadOnly
                  

If RcSchedule.Recordcount > 0 Then
    Irow = 0
    RcJadwal.MoveFirst
    For Irow = 0 To 5 '(mRowLast + 5) 'set baris
        jCol = 3
            Do While jCol <= (mCount + 2)
                DgJadwal.row = Irow + 1 'DgJadwal.RowBookmark(DgJadwal.Row)
                DgJadwal.col = jCol
                RcSchedule.MoveFirst
                Do While RcSchedule.EOF <> True
                    If DgJadwal.Columns(jCol).Caption = WeekNumber(RcSchedule.Fields("scheduledate")) + 1 Then
                        DgJadwal.Columns(jCol).Value = RcSchedule.Fields("sumqty") 'ambil nilai
                        DgJadwal_AfterColEdit jCol 'refresh nilai
                    End If
                    RcSchedule.MoveNext
                Loop
                jCol = jCol + 1
            Loop
           RcJadwal.MoveNext
           Irow = Irow + 5
    Next Irow
Else
    MessageBox "Data QTY  di Purchasing Tidak Ada Untuk Schedule ", FrmMRPLogic.Caption, msgOkOnly, msgExclamation
End If
End Sub

Private Sub refresh_grid()
On Error Resume Next
Dim Irow, jCol As Integer
'di gunakan untuk refresh nilai
If RcJadwal.Recordcount > 0 Then
RcJadwal.MoveFirst
For Irow = 0 To 5 '(mRowLast + 5) 'set baris
        jCol = 3
            Do While jCol <= (mCount + 2)
                DgJadwal.row = Irow 'DgJadwal.RowBookmark(DgJadwal.Row)
                DgJadwal.col = jCol
                RcSchedule.MoveFirst
                DgJadwal_AfterColEdit jCol
                jCol = jCol + 1
            Loop
           RcJadwal.MoveNext
           Irow = Irow + 5
    Next Irow
End If
End Sub
