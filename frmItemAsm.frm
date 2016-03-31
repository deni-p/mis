VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{2A1DDFC8-F968-4ED0-BD2F-8A462E9BB934}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmItemAsm 
   Caption         =   "Pengolahan Bahan Baku"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmItemAsm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
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
      Height          =   6300
      Left            =   90
      ScaleHeight     =   6240
      ScaleWidth      =   10050
      TabIndex        =   9
      Top             =   0
      Width           =   10110
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   5100
         Left            =   105
         ScaleHeight     =   5040
         ScaleWidth      =   9495
         TabIndex        =   10
         Top             =   645
         Width           =   9555
         Begin MSDataGridLib.DataGrid DGPurchase 
            Height          =   2730
            Left            =   105
            TabIndex        =   6
            Top             =   1665
            Width           =   9270
            _ExtentX        =   16351
            _ExtentY        =   4815
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
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
            ColumnCount     =   7
            BeginProperty Column00 
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
            BeginProperty Column01 
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
            BeginProperty Column02 
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
            BeginProperty Column03 
               DataField       =   "Qty In"
               Caption         =   "QTY. In"
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
            BeginProperty Column04 
               DataField       =   "Qty Used"
               Caption         =   "Qty Proses"
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
            BeginProperty Column05 
               DataField       =   "Srink"
               Caption         =   "Persen"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   5
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Total"
               Caption         =   "Persen B"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   5
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2610.142
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915.024
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "UOM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   2
            Tag             =   "ASM"
            Top             =   855
            Width           =   3450
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Asm Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   510
            Width           =   3450
         End
         Begin MSDataListLib.DataCombo cboRakit 
            DataField       =   "WareHouse"
            Height          =   330
            Index           =   0
            Left            =   6150
            TabIndex        =   4
            Tag             =   "ASM"
            Top             =   150
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Asm Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   0
            Tag             =   "ASM"
            Top             =   165
            Width           =   3450
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Notes"
            Height          =   315
            Index           =   2
            Left            =   1545
            MaxLength       =   200
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   4530
            Width           =   4575
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Date Asm"
            Height          =   315
            Left            =   6150
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   495
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dddd dd/MMMM/yyyy"
            Format          =   59441155
            CurrentDate     =   38272
         End
         Begin MSDataListLib.DataCombo cboRakit 
            DataField       =   "NoGroup"
            Height          =   330
            Index           =   1
            Left            =   1800
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   1200
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Line Line1 
            Index           =   6
            X1              =   255
            X2              =   1980
            Y1              =   4830
            Y2              =   4830
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   5325
            X2              =   7050
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   5340
            X2              =   7065
            Y1              =   465
            Y2              =   465
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   135
            X2              =   1860
            Y1              =   1515
            Y2              =   1515
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   120
            X2              =   1845
            Y1              =   1170
            Y2              =   1170
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   105
            X2              =   1830
            Y1              =   825
            Y2              =   825
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   165
            X2              =   1890
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelompok Barang"
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
            Left            =   135
            TabIndex        =   17
            Top             =   1215
            Width           =   1605
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
            Index           =   4
            Left            =   135
            TabIndex        =   16
            Top             =   915
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
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
            Left            =   135
            TabIndex        =   15
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Bikin"
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
            Left            =   5325
            TabIndex        =   14
            Top             =   540
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
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
            Left            =   135
            TabIndex        =   13
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gudang"
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
            Left            =   5325
            TabIndex        =   12
            Top             =   195
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
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
            Left            =   300
            TabIndex        =   11
            Top             =   4575
            Width           =   1065
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   6960
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   873
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmItemAsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcGudang As New DBQuick
Private RcGroup As New DBQuick
Private RcDetail As New DBQuick
Private RcPartner As New DBQuick
Private MyData As New clsTransaksi
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mVarAdd As Boolean

Private Sub cboRakit_Click(Index As Integer, Area As Integer)
If Index = 1 And mVarAdd = True Then MyDDE.GetFieldByName("Asm Order") = MyData.PrepareIndex(tmbAsmOrder, 10, cboRakit(1).BoundText, cboRakit(1).BoundText & "/")
End Sub

Private Sub cboRakit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim mTmp, mVarVal As Variant

Select Case DGPurchase.Col
       Case 3, 4:
            mVarVal = CekStock(MyDDE.ChildRecordset("Kode Barang")) - MyDDE.ChildRecordset.Fields(DGPurchase.Columns(ColIndex).DataField)
            If mVarVal < 0 Then
               MessageBox "Stock Tidak Cukup Untuk Melakukan Transaksi." & vbCrLf & "Stok Kurang -> " & mVarVal & " Untuk Memenuhi Transaksi SC", "Peringatan", msgOkOnly
               MyDDE.ChildRecordset.Fields(DGPurchase.Columns(ColIndex).DataField) = 0
            Else
               mTmp = DGPurchase.Columns(3) - DGPurchase.Columns(4)
               DGPurchase.Columns(6) = (mTmp / 1000) * (1 / 100)
            End If
End Select
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Select Case DGPurchase.Col
       Case 3, 4: DGPurchase.AllowUpdate = True
       Case Else: DGPurchase.AllowUpdate = False
End Select
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
DTPicker1.Value = dDateBegin
OpenGudang
OpenKelompok
With MyDDE
    .SetPermissions = UserEditDeleteDenied
    .EditModeReplace = False
    Set .BindForm = frmItemAsm
    .BindFormTAG = "ASM"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT     [Asm Order], [Date Asm], [Date Issued], Notes, WareHouse, [Asm Name], UOM, NoGroup FROM         [Prepare Inventory]"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGudang.CloseDB
RcGroup.CloseDB
RcDetail.CloseDB
RcPartner.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Resize()

HiasForm Picture1, Me
CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmItemAsm = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
With MyDDE.ChildRecordset
     .Fields(0) = mCall.GetFieldByName(0)
     .Fields(1) = mCall.GetFieldByName(1)
     .Fields(2) = mCall.GetFieldByName(2)
     .Fields(3) = mCall.GetFieldByName("Stok")
     .Fields(4) = 0
     .Fields("Harga") = mCall.GetFieldByName(5)
End With
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            txtBox(0).Enabled = False
       Case tmbAddNew:
            With MyDDE
                 .GetFieldByName("Asm Order") = MyData.PrepareIndex(tmbAsmOrder, 10, cboRakit(1).BoundText, cboRakit(1).BoundText & "/")
                 .GetFieldByName("Notes") = "-"
                 .GetFieldByName("UOM") = "KG"
                 .GetFieldByName("Date Asm") = dDateBegin
            End With
            txtBox(0).SetFocus
            mVarAdd = True
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               OpenPartner
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & txtBox(0) & "') ")
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
            End If
       Case tmbPrint:
            CallRPTReport "Raw Material.Rpt", "Select * from [Raw Material] Where [Kode RM]='" & txtBox(0) & "'"
'       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Opendetail IIf(Not IsNull(MyDDE.GetFieldByName("Asm Order")), MyDDE.GetFieldByName("Asm Order"), "XXXXX")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada. Harap diisi dulu.", "Peringatan", msgOkOnly
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
            mVarAdd = False
       
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
            mVarAdd = False
       Case tmbCancel: mVarAdd = False

End Select
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Prepare Inventory]" & _
                     " ([Asm Order], [Date Asm], Notes, WareHouse, [Asm Name], UOM, NoGroup)" & _
                     " VALUES  (N'" & txtBox(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & txtBox(2) & "', N'" & cboRakit(0).BoundText & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(3)) & "', N'" & cboRakit(1).BoundText & "')"

    .PrepareUpdate = " UPDATE [Prepare Inventory]" & _
                     " SET [Date Asm] = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Notes = N'" & txtBox(2) & "'," & _
                     " WareHouse = N'" & cboRakit(0).BoundText & "', [Asm Name] = N'" & ValidString(txtBox(1)) & "', UOM = N'" & ValidString(txtBox(3)) & "', NoGroup = N'" & cboRakit(1).BoundText & "'" & _
                     " WHERE ([Asm Order] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Prepare Inventory] WHERE     ([Asm Order] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub OpenGudang()
RcGudang.DBOpen "Select * from WareHouse order by WareHouse", Cnn, lckLockReadOnly
Set cboRakit(0).RowSource = RcGudang.DBRecordset
End Sub

Private Function OpenKelompok()
RcGroup.DBOpen "Select * from [Inventory Group] where status =0 order by NoGroup", Cnn, lckLockReadOnly
Set cboRakit(1).RowSource = RcGroup.DBRecordset
End Function

Private Sub SimpanDetail()
If CekDataItem = False Then
   SendDataToServer (" INSERT INTO Inventory" & _
                     " (NoItem, WareHouse, NoGroup, ItemName, UOM,StatusItem)" & _
                     " VALUES (N'" & txtBox(0) & "', N'" & cboRakit(0).BoundText & "', N'" & cboRakit(1).BoundText & "', N'" & txtBox(1) & "', N'" & txtBox(3) & "',N'OLAHAN')")
Else
   SendDataToServer (" UPDATE    Inventory" & _
                     " Set WareHouse = N'" & cboRakit(0).BoundText & "', NoGroup = N'" & cboRakit(1).BoundText & "', ItemName = N'" & txtBox(1) & "'," & _
                     " UOM = N'" & txtBox(3) & "', StatusItem = N'OLAHAN'" & _
                     " WHERE     (NoItem = N'" & txtBox(0) & "')")

End If
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Raw Materials] WHERE     ([Asm Order] = N'" & txtBox(0) & "')") = True Then
           .MoveFirst
           Do
             If .EOF = True Then Exit Do
             If SendDataToServer(" INSERT INTO [Raw Materials] " & _
                                 " ([Asm Order], NoItem, [Qty In], [Qty Used],Harga)" & _
                                 " VALUES (N'" & txtBox(0) & "', N'" & .Fields(0) & "', " & CCur(.Fields("Qty In")) & "," & .Fields("Harga") & ", " & CCur(.Fields("Qty Used")) & ")") = True Then
                'If SendDataToServer("DELETE FROM [Inventory Tabel] WHERE     (RefTrans = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields(0) & "')") = True Then
                SendARItem .Fields(0), CCur(.Fields("Qty In")), .Fields("Harga"), txtBox(0), DTPicker1.Value, .Fields("Harga"), "IAB"
                'End If
             End If
             .MoveNext
           Loop
           .MoveLast
           SendAPItem txtBox(0), CCur(.Fields("Qty Used")), .Fields("Harga"), txtBox(0), DTPicker1.Value, "IAJ"
        End If
     End If
End With
End Sub

Private Sub Opendetail(ByVal ParamString As String)
RcDetail.DBOpen " SELECT [Raw Materials].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM, [Raw Materials].[Qty In], [Raw Materials].[Qty Used], (([Raw Materials].[Qty In] - [Raw Materials].[Qty Used]) / 1000) * (1 / 100) AS Srink, Inventory.PriceIn AS Total, [Raw Materials].Harga" & _
                " FROM [Raw Materials] INNER JOIN Inventory ON [Raw Materials].NoItem = Inventory.NoItem WHERE ([Raw Materials].[Asm Order] = N'" & ParamString & "') ORDER BY [Raw Materials].NoItem", Cnn
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Sub OpenPartner()
RcPartner.DBOpen "SELECT     Inventory.NoItem AS [No Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit, Inventory.PPn, SUM([Inventory Tabel].QTY_IN)                        - SUM([Inventory Tabel].QTY_OUT) AS Stok, MAX([Inventory Tabel].PriceIn) AS Harga FROM         Inventory INNER JOIN                       [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem INNER JOIN                       [Inventory Group] ON Inventory.NoGroup = [Inventory Group].NoGroup GROUP BY Inventory.NoItem, Inventory.ItemName, Inventory.UOM, Inventory.PPn, [Inventory Group].Status, Inventory.StatusItem HAVING      (SUM([Inventory Tabel].QTY_IN) - SUM([Inventory Tabel].QTY_OUT) > 0) AND ([Inventory Group].Status = 0) AND (Inventory.StatusItem <> N'OLAHAN') ORDER BY Inventory.NoItem", Cnn, lckLockReadOnly
If RcPartner.Recordcount <> 0 Then
    Set mCall = New frmCaller
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.SetFormat(4) = "#,##0"
    mCall.SetFormat(5) = "#,##0"
    mCall.SetAlignmentFormat(4) = 1
    mCall.SetAlignmentFormat(5) = 1
    mCall.Caption = "MASTER BARANG"
    mCall.LookUp Me
    If FindOwnRecordset(MyDDE.ChildRecordset, "[Kode Barang] = '" & mCall.GetFieldByName("No Barang") & "'") = True Then
       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
       DGPurchase.SetFocus
    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   MyDDE.CancelTrans = True
End If
RcPartner.CloseDB
Set mCall = Nothing
End Sub

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT     SUM(QTY_IN) - SUM(QTY_OUT) AS Stock FROM         [Inventory Tabel] WHERE     (NoItem = N'" & NoItem & "') HAVING      (SUM(QTY_IN) - SUM(QTY_OUT) <> 0)", Cnn, lckLockReadOnly
'MessageBox "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) <> N'DN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')"
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
End With
RcCek.CloseDB
End Function

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Function CekDataItem() As Boolean
Dim Rcitem As New DBQuick
Rcitem.DBOpen "select * from Inventory where NoItem=N'" & txtBox(0) & "'", Cnn, lckLockReadOnly
With Rcitem
     If .Recordcount <> 0 Then
        CekDataItem = True
     End If
End With
Rcitem.CloseDB
End Function
