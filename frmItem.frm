VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form frmItem 
   Caption         =   "Master Item"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   6765
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
      ForeColor       =   &H80000015&
      Height          =   6570
      Left            =   105
      ScaleHeight     =   6510
      ScaleWidth      =   8850
      TabIndex        =   2
      Top             =   0
      Width           =   8910
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   5760
         Left            =   405
         ScaleHeight     =   5700
         ScaleWidth      =   8085
         TabIndex        =   3
         Top             =   540
         Width           =   8145
         Begin VB.TextBox txtBox 
            DataField       =   "MaxStock"
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
            Left            =   3810
            MaxLength       =   4
            TabIndex        =   10
            Tag             =   "Partner"
            Top             =   2460
            Width           =   675
         End
         Begin VB.TextBox txtBox 
            DataField       =   "MinStock"
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
            Left            =   2055
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "Partner"
            Top             =   2445
            Width           =   675
         End
         Begin VB.TextBox txtBox 
            DataField       =   "UOM"
            Height          =   315
            Index           =   4
            Left            =   2055
            MaxLength       =   15
            TabIndex        =   8
            Tag             =   "Partner"
            Top             =   2085
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Serial Supplier"
            Height          =   315
            Index           =   3
            Left            =   2055
            MaxLength       =   25
            TabIndex        =   7
            Tag             =   "Partner"
            Top             =   1350
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Merk"
            Height          =   315
            Index           =   2
            Left            =   5730
            MaxLength       =   25
            TabIndex        =   6
            Tag             =   "Partner"
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtBox 
            DataField       =   "ItemName"
            Height          =   315
            Index           =   1
            Left            =   2055
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Partner"
            Top             =   615
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "NoItem"
            Height          =   315
            Index           =   0
            Left            =   2055
            MaxLength       =   15
            TabIndex        =   4
            Tag             =   "Partner"
            Top             =   240
            Width           =   1935
         End
         Begin MSDataListLib.DataCombo CboGudang 
            DataField       =   "WareHouse"
            Height          =   330
            Index           =   0
            Left            =   2055
            TabIndex        =   11
            Tag             =   "Partner"
            Top             =   1725
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "WareHouse Name"
            BoundColumn     =   "WareHouse"
            Text            =   ""
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2580
            Index           =   0
            Left            =   210
            TabIndex        =   12
            Tag             =   "Partner"
            Top             =   2925
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   4551
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
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "NoItem"
               Caption         =   "No Item"
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
               DataField       =   "WareHouse"
               Caption         =   "Gudang"
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
               DataField       =   "NoGroup"
               Caption         =   "Kelompok"
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
               DataField       =   "ItemName"
               Caption         =   "Nama Item/Service"
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
            BeginProperty Column05 
               DataField       =   "Serial Supplier"
               Caption         =   "Serial Supplier ID"
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
               DataField       =   "UOM"
               Caption         =   "Unit"
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
               DataField       =   "MinStock"
               Caption         =   "MinStock"
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
               DataField       =   "MaxStock"
               Caption         =   "MaxStock"
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
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo CboGudang 
            DataField       =   "NoGroup"
            Height          =   330
            Index           =   1
            Left            =   2055
            TabIndex        =   13
            Tag             =   "Partner"
            Top             =   975
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "Group Name"
            BoundColumn     =   "NoGroup"
            Text            =   ""
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min. Stok:              Max. Stok:"
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
            Left            =   1005
            TabIndex        =   20
            Top             =   2505
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit:"
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
            Left            =   1545
            TabIndex        =   19
            Top             =   2175
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gudang:"
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
            Left            =   1215
            TabIndex        =   18
            Top             =   1815
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Supplier ID:"
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
            Left            =   345
            TabIndex        =   17
            Top             =   1440
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelompok:"
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
            Left            =   1020
            TabIndex        =   16
            Top             =   1035
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Item/Service:                                                     Merk:"
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
            Left            =   195
            TabIndex        =   15
            Top             =   630
            Width           =   5490
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item/Service ID:"
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
            Left            =   465
            TabIndex        =   14
            Top             =   270
            Width           =   1515
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6765
      TabIndex        =   0
      Top             =   4035
      Width           =   6765
      Begin VB.Label Batal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2->Edit   F3->Tambah   F4->Delete   F5->Simpan   F6->Batal   F7->Bantuan   F8->Print   F9->Keluar"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   7380
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   21
      Top             =   3570
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   820
      BindFormTAG     =   "Partner"
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyData As New clsMaster

Private Sub CboGudang_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
Set CboGudang(0).RowSource = MyData.OpenGudang
Set CboGudang(1).RowSource = MyData.OpenKelompok
OpenDB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      Set MyData = Nothing
      MyDDE.ClearRecordset
   Else
      Cancel = True
   End If
Else
   Set MyData = Nothing
   MyDDE.ClearRecordset
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
HiasForm Picture1, Me
CenterForm Picture2
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmItemData = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            txtBox(0) = MyData.PrepareIndex(tmbInventory, 5, "", "IT/")
            SetelKosong
            mVarDataDc = True
            txtBox(0).Enabled = False
            txtBox(1).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            mVarDataDc = True
            If txtBox(1).Enabled = True Then txtBox(1).SetFocus
       Case tmbPrint:
            CallRPTReport "Tabel Item.rpt"
       Case Else: mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               MyDDE.GetFieldByName("WareHouse") = CboGudang(0).BoundText
               MyDDE.GetFieldByName("NoGroup") = CboGudang(1).BoundText
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub OpenDB()
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmItemData
    .BindFormTAG = "Partner"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "Select * from Inventory where StatusItem='ITEM' Order By NoItem"
End With
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO Inventory ( NoItem, WareHouse, NoGroup,  ItemName, Merk, [Serial Supplier], UOM, MinStock, MaxStock, StatusItem) " & _
                     " VALUES (N'" & txtBox(0) & "', N'" & CboGudang(0).BoundText & "', N'" & CboGudang(1).BoundText & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "', N'" & ValidString(txtBox(3)) & "', N'" & ValidString(txtBox(4)) & "' ," & _
                     "  N'" & CDbl(txtBox(5)) & "', N'" & CDbl(txtBox(5)) & "',  N'ITEM')"
                     
    .PrepareUpdate = " UPDATE    Inventory Set WareHouse = N'" & CboGudang(0).BoundText & "', NoGroup = N'" & CboGudang(1).BoundText & "', ItemName = N'" & ValidString(txtBox(1)) & "', Merk = N'" & ValidString(txtBox(2)) & "', [Serial Supplier] = N'" & ValidString(txtBox(3)) & "', UOM = N'" & ValidString(txtBox(4)) & "', MinStock = N'" & CDbl(txtBox(5)) & "', MaxStock = N'" & CDbl(txtBox(6)) & "'" & _
                     " WHERE     (StatusItem = N'ITEM') AND (NoItem = N'" & ValidString(txtBox(0)) & "')"
                     
    .PrepareDelete = " DELETE FROM Inventory WHERE   (StatusItem = N'ITEM') AND (NoItem = N'" & ValidString(txtBox(0)) & "')"
End With
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
       Case 5, 6:
            ValidNum KeyAscii
       Case Else:
End Select
End Sub

Private Sub txtBox_Validate(Index As Integer, Cancel As Boolean)
If txtBox(Index) = "" Then
   MessageBox "Data Tidak Boleh Kosong..........!", "Peringatan", msgOkOnly
   Cancel = True
End If
End Sub


Private Sub SetelKosong()
With MyDDE
     .GetFieldByName("ItemName") = "-"
     .GetFieldByName("Merk") = "-"
     .GetFieldByName("Serial Supplier") = "-"
     .GetFieldByName("UOM") = "PCS"
     .GetFieldByName("MinStock") = 0
     .GetFieldByName("MaxStock") = 0
End With
End Sub


