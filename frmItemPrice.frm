VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form frmItemPrice 
   AutoRedraw      =   -1  'True
   Caption         =   "Seting Harga"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   9465
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   5535
      Left            =   60
      ScaleHeight     =   5475
      ScaleWidth      =   9210
      TabIndex        =   9
      Top             =   285
      Width           =   9270
      Begin VB.PictureBox Picture2 
         Height          =   5145
         Left            =   180
         ScaleHeight     =   5085
         ScaleWidth      =   8820
         TabIndex        =   10
         Top             =   195
         Width           =   8880
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4800
            Index           =   0
            Left            =   105
            TabIndex        =   0
            Tag             =   "PRICE"
            Top             =   165
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   8467
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   2
            RowHeight       =   15
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "Itemname"
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
            BeginProperty Column01 
               DataField       =   "Pricein"
               Caption         =   "Harga Beli"
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
            BeginProperty Column02 
               DataField       =   "HargaJual"
               Caption         =   "Harga Jual"
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
               MarqueeStyle    =   4
               BeginProperty Column00 
                  ColumnWidth     =   3254.74
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1560.189
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1560.189
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtBox 
            DataField       =   "NoItem"
            Height          =   315
            Index           =   0
            Left            =   5670
            MaxLength       =   15
            TabIndex        =   1
            Tag             =   "PRICE"
            Top             =   345
            Width           =   1935
         End
         Begin VB.TextBox txtBox 
            DataField       =   "ItemName"
            Height          =   315
            Index           =   1
            Left            =   5670
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "PRICE"
            Top             =   675
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Serial Supplier"
            Height          =   315
            Index           =   2
            Left            =   5670
            MaxLength       =   25
            TabIndex        =   3
            Tag             =   "PRICE"
            Top             =   990
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "UOM"
            Height          =   315
            Index           =   3
            Left            =   5670
            MaxLength       =   25
            TabIndex        =   4
            Tag             =   "PRICE"
            Top             =   1320
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "PPN"
            Height          =   315
            Index           =   4
            Left            =   5670
            MaxLength       =   3
            TabIndex        =   5
            Tag             =   "PRICE"
            Top             =   1635
            Width           =   675
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Markup"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   7530
            MaxLength       =   3
            TabIndex        =   6
            Tag             =   "PRICE"
            Top             =   1635
            Width           =   675
         End
         Begin VB.TextBox txtBox 
            DataField       =   "PriceIn"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   5670
            MaxLength       =   12
            TabIndex        =   7
            Tag             =   "PRICE"
            Top             =   1965
            Width           =   1905
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item/Service ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   4080
            TabIndex        =   24
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Item/Service"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   3795
            TabIndex        =   23
            Top             =   705
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Supplier ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   3960
            TabIndex        =   22
            Top             =   1035
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   5
            Left            =   5160
            TabIndex        =   21
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPn              %   Margin              %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   6
            Left            =   5190
            TabIndex        =   20
            Top             =   1680
            Width           =   3270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   4995
            TabIndex        =   19
            Top             =   1995
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   4905
            TabIndex        =   18
            Top             =   2280
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   7
            Left            =   4650
            TabIndex        =   17
            Top             =   2595
            Width           =   885
         End
         Begin VB.Label lblHarga 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   5670
            TabIndex        =   16
            Top             =   2295
            Width           =   3045
         End
         Begin VB.Label lblHarga 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   5670
            TabIndex        =   15
            Top             =   2610
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   8
            Left            =   5175
            TabIndex        =   14
            Top             =   2910
            Width           =   360
         End
         Begin VB.Label lblHarga 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   5670
            TabIndex        =   13
            Top             =   2895
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Jual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   9
            Left            =   4575
            TabIndex        =   12
            Top             =   3225
            Width           =   960
         End
         Begin VB.Label lblHarga 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   5670
            TabIndex        =   11
            Top             =   3180
            Width           =   3045
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   5835
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1217
      BindFormTAG     =   "PRICE"
   End
End
Attribute VB_Name = "frmItemPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mEdit As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
OpenDataHarga
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MyDDE.ClearRecordset
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMaximized Then
   Me.Height = MainMenu.ScaleHeight
   Me.Width = MainMenu.ScaleWidth
End If
HiasForm Picture1, Me
CenterForm Picture2
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            txtBox(0).Enabled = False
            txtBox(1).Enabled = False
            txtBox(2).Enabled = False
            txtBox(3).Enabled = False
            mEdit = True
            txtBox(6) = Harga(MyDDE.GetFieldByName("NoItem"))
       Case tmbPrint:
            Dim Mprint As New frmReportView
            With Mprint
                 .QuerySource = "Select * from rptPriceList Order By NoItem"
                 .ReportName = "Price List.rpt"
                 .Show
            End With
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
With MyDDE
    lblHarga(1) = FormatNumber(IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0) * (.GetFieldByName("Markup") / 100))
    lblHarga(2) = FormatNumber(CDbl(IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0) * (.GetFieldByName("Markup") / 100)) + CDbl(IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0)))
    lblHarga(3) = FormatNumber(IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0) * (.GetFieldByName("PPn") / 100))
    lblHarga(4) = FormatNumber(CDbl(IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0) * (.GetFieldByName("Markup") / 100)) + CDbl(IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0)) + IIf(Not IsNull(.GetFieldByName("PriceIn")), .GetFieldByName("PriceIn"), 0) * (.GetFieldByName("PPn") / 100))
End With
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            PrepareQuery
            mEdit = False
End Select
End Sub

Private Sub txtBox_Change(Index As Integer)
If Index = 4 Or Index = 5 Or Index = 6 Then
   If mEdit = True Then
      If txtBox(Index) = "" Then txtBox(Index) = "0"
      lblHarga(1) = FormatNumber(IIf(Not IsNull(txtBox(6)), txtBox(6), 0) * (txtBox(5) / 100)) '+ IIf(Not IsNull(txtBox(6)), txtBox(6), 0))
      lblHarga(2) = FormatNumber(CDbl(IIf(Not IsNull(txtBox(6)), txtBox(6), 0) * (txtBox(5) / 100)) + CDbl(IIf(Not IsNull(txtBox(6)), txtBox(6), 0)))
      lblHarga(3) = FormatNumber(IIf(Not IsNull(txtBox(6)), txtBox(6), 0) * (txtBox(4) / 100))
      lblHarga(4) = FormatNumber(CDbl(IIf(Not IsNull(txtBox(6)), txtBox(6), 0) * (txtBox(5) / 100)) + CDbl(IIf(Not IsNull(txtBox(6)), txtBox(6), 0)) + IIf(Not IsNull(txtBox(6)), txtBox(6), 0) * (txtBox(4) / 100))
   End If
End If
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
ValidNum KeyAscii
End Sub

Private Sub OpenDataHarga()
With MyDDE
     .EditModeReplace = False
     Set .BindForm = frmItemPrice
     Set .ActiveConnection = Cnn
     .PrepareQuery = " SELECT *, PriceIn * (PPn / 100) + PriceIn * (Markup / 100) + PriceIn AS HargaJual from Inventory Order By NoItem"
     .BindFormTAG = "PRICE"
     .EditModeReplace = False
     .SetPermissions = UserAddnewDeleteDenied
     .IsChildMemberReady = True
End With
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO Inventory (NoItem, WareHouse, NoGroup,  ItemName, Merk, [Serial Supplier], UOM, MinStock, MaxStock, StatusItem, PPn, Markup)" & _
                     " VALUES (N'" & txtBox(0) & "', N'" & .GetFieldByName("WareHouse") & "', N'" & .GetFieldByName("NoGroup") & "', N'" & txtBox(1) & "', N'" & .GetFieldByName("Merk") & "', N'" & IIf(Not IsNull(txtBox(2)), txtBox(2), "-") & "', N'" & txtBox(3) & "', " & CDbl(.GetFieldByName("MinStock")) & ", " & CDbl(.GetFieldByName("MaxStock")) & ", N'" & .GetFieldByName("StatusItem") & "', " & CDbl(.GetFieldByName("PPn")) & ", " & CDbl(.GetFieldByName("Markup")) & ")"
    .PrepareUpdate = " UPDATE Inventory" & _
                     " SET PPn = " & CDbl(txtBox(4)) & " , Markup = " & CDbl(txtBox(5)) & ",PriceIn=" & CDbl(txtBox(6)) & " WHERE     (NoItem = N'" & txtBox(0) & "') AND (StatusItem = N'ITEM')"
    .PrepareDelete = " DELETE FROM Inventory WHERE   (StatusItem = N'ITEM') AND (NoItem = N'" & txtBox(0) & "')"
End With
End Sub

Private Function Harga(ByVal NoItem As String) As Long
Dim RcHrg As New Recordset
RcHrg.CursorLocation = adUseClient
RcHrg.Open "SELECT MAX(PriceIn) AS PriceIn FROM         [Inventory Tabel] GROUP BY NoItem, LEFT(RefTrans, 2) HAVING (NoItem = N'" & NoItem & "') AND (LEFT(RefTrans, 2) = N'rn')", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
Harga = 0
With RcHrg
     If .Recordcount <> 0 Then
         Harga = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     End If
     .Close
End With
Set RcHrg = Nothing
End Function

