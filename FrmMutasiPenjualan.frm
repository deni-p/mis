VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C5BD1BD0-C880-4C3C-8176-E61FC2E2B3F5}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmMutasiPenjualan 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10830
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
   Icon            =   "FrmMutasiPenjualan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Tag             =   "Mutasi Penjualan"
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
      Height          =   6300
      Left            =   90
      ScaleHeight     =   6270
      ScaleWidth      =   9930
      TabIndex        =   9
      Top             =   0
      Width           =   9960
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00EAAF6F&
         Height          =   5100
         Left            =   90
         ScaleHeight     =   5040
         ScaleWidth      =   9495
         TabIndex        =   10
         Top             =   630
         Width           =   9555
         Begin VB.TextBox txtBox 
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
         Begin VB.TextBox txtBox 
            DataField       =   "MutasiID"
            Height          =   330
            Index           =   0
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   0
            Tag             =   "ASM"
            Top             =   90
            Width           =   3450
         End
         Begin VB.TextBox txtBox 
            DataField       =   "MutasiName"
            Height          =   330
            Index           =   1
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   420
            Width           =   3450
         End
         Begin VB.TextBox txtBox 
            DataField       =   "UOM"
            Height          =   330
            Index           =   3
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   2
            Tag             =   "ASM"
            Top             =   765
            Width           =   3450
         End
         Begin MSDataListLib.DataCombo cboRakit 
            DataField       =   "WareHouse"
            Height          =   330
            Index           =   0
            Left            =   6150
            TabIndex        =   4
            Tag             =   "ASM"
            Top             =   75
            Width           =   3150
            _ExtentX        =   5556
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
            ColumnCount     =   6
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
               DataField       =   "QTY Retur"
               Caption         =   "QTY Mutasi"
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
            BeginProperty Column05 
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
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Date Asm"
            Height          =   315
            Left            =   6150
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   420
            Width           =   3150
            _ExtentX        =   5556
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
            CustomFormat    =   "dddd dd/MMMM/yyyy"
            Format          =   53346307
            CurrentDate     =   38272
         End
         Begin MSDataListLib.DataCombo cboRakit 
            DataField       =   "NoGroup"
            Height          =   330
            Index           =   1
            Left            =   1800
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   1110
            Width           =   3450
            _ExtentX        =   6085
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
            Caption         =   "Keterangan :"
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
            TabIndex        =   17
            Top             =   4575
            Width           =   1185
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
            Left            =   5385
            TabIndex        =   16
            Top             =   120
            Width           =   705
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
            Left            =   570
            TabIndex        =   15
            Top             =   135
            Width           =   1170
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
            Top             =   465
            Width           =   765
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
            Left            =   555
            TabIndex        =   13
            Top             =   465
            Width           =   1200
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
            Left            =   1365
            TabIndex        =   12
            Top             =   825
            Width           =   375
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
            TabIndex        =   11
            Top             =   1125
            Width           =   1605
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   6600
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
      LimitRecordData =   "1"
   End
End
Attribute VB_Name = "FrmMutasiPenjualan"
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
If Index = 1 And mVarAdd = True Then MyDDE.GetFieldByName("MutasiID") = MyData.PrepareIndex(tmbTransaksiMutasiPenjualan, 5, cboRakit(1).BoundText, cboRakit(1).BoundText & "/")
End Sub

Private Sub cboRakit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
'Dim mTmp, mVarVal As Variant

Select Case DGPurchase.Col
       Case 3, 4:
'            mVarVal = CekStock(MyDDE.ChildRecordset("Kode Barang")) - MyDDE.ChildRecordset.Fields(DGPurchase.Columns(ColIndex).DataField)
'            If mVarVal < 0 Then
'               MessageBox "Stock Tidak Cukup Untuk Melakukan Transaksi." & vbCrLf & "Stok Kurang -> " & mVarVal & " Untuk Memenuhi Transaksi SC", "Peringatan", msgOkOnly
'               MyDDE.ChildRecordset.Fields(DGPurchase.Columns(ColIndex).DataField) = 0
'            Else
'               mTmp = DGPurchase.Columns(3) - DGPurchase.Columns(4)
'               'DGPurchase.Columns(6) = (mTmp / 1000) * (1 / 100)
'            End If
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
HiasForm Picture1, Me
CenterForm Picture2, Me
DTPicker1.Value = dDateBegin
OpenGudang
OpenKelompok
With MyDDE
    .SetPermissions = UserEditDeleteDenied
    .LimitRecordData = 1
    .EditModeReplace = False
    Set .BindForm = FrmMutasiPenjualan
    .BindFormTAG = "ASM"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT MutasiID, NoGroup, MutasiName, MutasiDate, Notes, WareHouse, UOM FROM [Mutasi Jual] ORDER BY MutasiID"
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

'HiasForm Picture1, Me
'CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmMutasiPenjualan = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
With MyDDE.ChildRecordset
     .Fields(0) = mCall.GetFieldByName(0)
     .Fields(1) = mCall.GetFieldByName(1)
     .Fields(2) = mCall.GetFieldByName(2)
     .Fields(3) = mCall.GetFieldByName("Stok")
     .Fields(5) = mCall.GetFieldByName("Harga")
     .Fields(6) = mCall.GetFieldByName("ppn")
     .Fields("Doc Reff") = mCall.GetFieldByName("Doc Reff")
End With
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            txtBox(0).Enabled = False
       Case tmbAddNew:
            With MyDDE
                 .GetFieldByName("MutasiID") = MyData.PrepareIndex(tmbTransaksiMutasiPenjualan, 5, cboRakit(1).BoundText, cboRakit(1).BoundText & "/")
                 .GetFieldByName("Notes") = "-"
                 .GetFieldByName("UOM") = "KG"
                 .GetFieldByName("MutasiDate") = dDateBegin
            End With
            txtBox(0).SetFocus
            mVarAdd = True
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner = True Then CancelDetailTrans
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
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("MutasiID")), MyDDE.GetFieldByName("MutasiID"), "XXXXX")
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
'       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               OpenPartner
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
End Select
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Mutasi Jual]" & _
                     " ([MutasiID], [MutasiDate], Notes, WareHouse, [MutasiName], UOM, NoGroup)" & _
                     " VALUES  (N'" & txtBox(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & txtBox(2) & "', N'" & cboRakit(0).BoundText & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(3)) & "', N'" & cboRakit(1).BoundText & "')"

    .PrepareUpdate = " UPDATE [Mutasi Jual]" & _
                     " SET [MutasiDate] = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), Notes = N'" & txtBox(2) & "'," & _
                     " WareHouse = N'" & cboRakit(0).BoundText & "', [MutasiName] = N'" & ValidString(txtBox(1)) & "', UOM = N'" & ValidString(txtBox(3)) & "', NoGroup = N'" & cboRakit(1).BoundText & "'" & _
                     " WHERE ([MutasiID] = N'" & txtBox(0) & "')"

    .PrepareDelete = " DELETE FROM [Mutasi Jual] WHERE     ([MutasiID] = N'" & txtBox(0) & "') "
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
SendDataToServer (" INSERT INTO Inventory" & _
                  " (NoItem, WareHouse, NoGroup, ItemName, UOM,StatusItem)" & _
                  " VALUES (N'" & txtBox(0) & "', N'" & cboRakit(0).BoundText & "', N'" & cboRakit(1).BoundText & "', N'" & txtBox(1) & "', N'" & txtBox(3) & "',N'MUTASI JUAL')")
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Detail mutasijual] WHERE     ([MutasiID] = N'" & txtBox(0) & "')") = True Then
           .MoveFirst
           Do
             If .EOF = True Then Exit Do
             If SendDataToServer(" INSERT INTO [Detail mutasijual] " & _
                                 " ([MutasiID], NoItem, [Qty In], [Qty Used],VAT,Price,ReturID)" & _
                                 " VALUES (N'" & txtBox(0) & "', N'" & .Fields(0) & "', " & CCur(.Fields("qty retur")) & ", " & CCur(.Fields("qty retur")) & "," & .Fields("ppn") & "," & .Fields("Harga") & ",N'" & .Fields("Doc Reff") & "')") = True Then
'                If SendDataToServer("DELETE FROM [Inventory Tabel] WHERE     (RefTrans = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields(0) & "')") = True Then
'                   SendAPItem .Fields(0), CCur(.Fields("Qty retur")), .Fields("Harga"), txtBox(0), DTPicker1.Value, "MJ"
'                   SendDataToServer ("UPDATE    [Detail Retur] SET              StatusItem = 1 WHERE     (ReturID = N'" & .Fields("Doc Reff") & "') AND (NoItem = N'" & .Fields(0) & "')")
'                End If
             End If
             .MoveNext
           Loop
           .MoveLast
           'SendAPItem txtBox(0), CCur(.Fields("Qty Used")), 0, txtBox(0), DTPicker1.Value
        End If
     End If
End With
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
 RcDetail.DBOpen " SELECT     [Detail MutasiJual].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang],Inventory.UOM, [Detail MutasiJual].[Qty In] AS [QTY Retur],  [Detail MutasiJual].[Qty Used] AS [QTY Mutasi], [Detail MutasiJual].Price AS Harga, [Detail MutasiJual].Vat AS Ppn,[Detail MutasiJual].ReturID AS [Doc Reff] FROM         [Detail MutasiJual] INNER JOIN                       Inventory ON [Detail MutasiJual].NoItem = Inventory.NoItem WHERE     ([Detail MutasiJual].MutasiID = N'" & ParamString & "') ORDER BY [Detail MutasiJual].NoItem", Cnn
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner() As Boolean
RcPartner.DBOpen "SELECT     Inventory.NoItem AS [No Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM AS Unit, [Detail Retur].VAT AS Ppn,                        [Detail Retur].[Retur Jual] AS Stok, [Detail Retur].Price AS Harga, [Detail Retur].ReturID AS [Doc Reff] FROM         Inventory INNER JOIN                       [Inventory Group] ON Inventory.NoGroup = [Inventory Group].NoGroup INNER JOIN                       [Detail Retur] ON Inventory.NoItem = [Detail Retur].NoItem WHERE     (LEFT([Detail Retur].ReturID, 2) = N'RJ') and ([detail retur].StatusItem=0) ORDER BY Inventory.NoItem", Cnn, lckLockReadOnly
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
       CancelDetailTrans
       DGPurchase.SetFocus
    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   OpenPartner = True
End If
RcPartner.CloseDB
Set mCall = Nothing
End Function

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) <> N'DN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", Cnn, lckLockReadOnly
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

Private Sub CancelDetailTrans()
If MyDDE.ChildRecordset.Recordcount <> 0 Then
  If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveNext
  If MyDDE.ChildRecordset.EOF And MyDDE.ChildRecordset.Recordcount > 0 Then MyDDE.ChildRecordset.MoveLast
End If
End Sub
