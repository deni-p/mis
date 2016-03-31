VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmMutasiGudang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mutasi Gudang"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMutasiGudang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Tag             =   "Stock Transfer"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   9945
      TabIndex        =   6
      Top             =   0
      Width           =   9945
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
         Height          =   315
         Index           =   2
         Left            =   1635
         MaxLength       =   200
         TabIndex        =   4
         Tag             =   "ASM"
         Top             =   4530
         Width           =   4575
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "NoMutasi"
         Height          =   330
         Index           =   0
         Left            =   1635
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "ASM"
         Top             =   90
         Width           =   3450
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal"
         Height          =   330
         Left            =   1635
         TabIndex        =   3
         Tag             =   "ASM"
         Top             =   810
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   139329539
         CurrentDate     =   38272
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Bindings        =   "frmMutasiGudang.frx":6852
         Height          =   3030
         Left            =   105
         TabIndex        =   5
         Top             =   1335
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   5345
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Kode Gudang"
            Caption         =   "Kode Gudang"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "Jumlah Masuk"
            Caption         =   "Stok Asal"
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
            DataField       =   "Jumlah Keluar"
            Caption         =   "Jumlah Mutasi"
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
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboRakit 
         DataField       =   "WareHouse"
         Height          =   330
         Index           =   0
         Left            =   1635
         TabIndex        =   2
         Tag             =   "ASM"
         Top             =   450
         Width           =   3450
         _ExtentX        =   6085
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
      Begin VB.Line Line1 
         Index           =   2
         X1              =   105
         X2              =   1680
         Y1              =   1110
         Y2              =   1110
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
         Left            =   150
         TabIndex        =   10
         Top             =   4575
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Tujuan"
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
         Left            =   150
         TabIndex        =   9
         Top             =   510
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Mutasi"
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
         Left            =   150
         TabIndex        =   8
         Top             =   135
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Mutasi"
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
         Left            =   150
         TabIndex        =   7
         Top             =   870
         Width           =   1395
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   1680
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   105
         X2              =   1680
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   105
         X2              =   1680
         Y1              =   4830
         Y2              =   4830
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5100
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmMutasiGudang"
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
Private mVarDocReff As String

Private Sub cboRakit_Click(Index As Integer, Area As Integer)
If Index = 1 And mVarAdd = True Then MyDDE.GetFieldByName("MutasiID") = MyData.PrepareIndex(tmbTransaksiMutasiPenjualan, 5, cboRakit(1).BoundText, cboRakit(1).BoundText & "/")
End Sub

Private Sub cboRakit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
'Dim mTmp, mVarVal As Variant

Select Case DGPurchase.col
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

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Select Case DGPurchase.col
       Case 4: DGPurchase.AllowUpdate = True
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
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
OpenGudang
'OpenKelompok
With MyDDE
    .SetPermissions = UserEditDeleteDenied
    .LimitRecordData = 1
    .EditModeReplace = False
    Set .BindForm = frmMutasiGudang
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  [Mutasi Gudang].NoMutasi, [Mutasi Gudang].WareHouse, WareHouse.[WareHouse Name], [Mutasi Gudang].Datetrans as Tanggal,                        [Mutasi Gudang].RefNotes FROM         [Mutasi Gudang] INNER JOIN                       WareHouse ON [Mutasi Gudang].WareHouse = WareHouse.WareHouse WHERE     ([Mutasi Gudang].Validasi = 0) ORDER BY [Mutasi Gudang].NoMutasi"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGudang.CloseDB
RcGroup.CloseDB
RcDetail.CloseDB
RcPartner.CloseDB
RcPartner.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
'
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMutasiGudang = Nothing
End Sub

Private Sub mCall_BeforeUnload()
If IsNull(MyDDE.ChildRecordset.Fields(0)) = True Then
   'If MyDDE.ChildRecordset.Fields(0) = "" Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   'End If
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
With MyDDE.ChildRecordset
     'MsgBox .Fields(3).Name
     .Fields(0) = mCall.GetFieldByName("Kode Gudang")
     .Fields(1) = mCall.GetFieldByName(1)
     .Fields(2) = mCall.GetFieldByName(0)
     .Fields(3) = mCall.GetFieldByName("Stok")
     .Fields(4) = 0
     .Fields(5) = HppProce(.Fields(1))
     .Fields(6) = mVarDocReff
End With
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            txtBox(0).Enabled = False
       Case tmbAddNew:
            txtBox(0).Enabled = False
            With MyDDE
                 .GetFieldByName("NoMutasi") = MyData.PrepareIndex(tmbTransaksiMutasiGudang, 5, "", TglIndex)
                 .GetFieldByName("Refnotes") = "Mutasi Gudang"
                 .GetFieldByName("Tanggal") = dDateBegin
            End With
            cboRakit(0).SetFocus
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
            CallRPTReport "Mutasi Internal Persediaan.rpt", "Select * from [Mutasi Internal Persediaan] Where [NoMutasi]='" & txtBox(0) & "'"
'       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("NoMutasi")), MyDDE.GetFieldByName("NoMutasi"), "XXXXX")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada. Harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
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
    .PrepareAppend = " INSERT INTO [Mutasi Gudang]" & _
                     " (NoMutasi, WareHouse, Datetrans, RefNotes)" & _
                     " VALUES     (N'" & FNumText(txtBox(0)) & "', N'" & cboRakit(0).BoundText & "', CONVERT(DATETIME, '" & Format(FDatePicker(DTPicker1.Value, "dd/mm/yy")) & "', 3), N'" & ValidString(FNumText(txtBox(2))) & "')"
'MsgBox .PrepareAppend
    .PrepareUpdate = " UPDATE    [Mutasi Gudang]" & _
                     " Set WareHouse = N'" & cboRakit(0).BoundText & "', Datetrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), RefNotes = N'" & ValidString(FNumText(txtBox(2))) & "'" & _
                     " WHERE     (NoMutasi = N'" & FNumText(txtBox(0)) & "')"

    .PrepareDelete = " DELETE FROM [Mutasi Gudang] WHERE     (NoMutasi = N'" & FNumText(txtBox(0)) & "') "
End With
End Sub

Private Sub OpenGudang()
RcGudang.DBOpen "Select * from WareHouse order by WareHouse", CNN, lckLockReadOnly
Set cboRakit(0).RowSource = RcGudang.DBRecordset
End Sub

Private Sub SimpanDetail()
'SendDataToServer (" INSERT INTO Inventory" & _
                  " (NoItem, WareHouse, NoGroup, ItemName, UOM,StatusItem)" & _
                  " VALUES (N'" & txtBox(0) & "', N'" & cboRakit(0).BoundText & "', N'" & cboRakit(1).BoundText & "', N'" & txtBox(1) & "', N'" & txtBox(3) & "',N'MUTASI JUAL')")
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Detail MutasiGudang] WHERE     ([NoMutasi] = N'" & FNumText(txtBox(0)) & "')") = True Then
           .MoveFirst
           Do
             If .EOF = True Then Exit Do
             If SendDataToServer(" INSERT INTO [Detail MutasiGudang]" & _
                                 " (WareHouse, NoMutasi, NoItem, [Jumlah Masuk], [Jumlah Keluar], Harga, [Doc Reff])" & _
                                 "  VALUES     (N'" & .Fields("Kode Gudang") & "', N'" & FNumText(txtBox(0)) & "', N'" & .Fields("Kode Barang") & "', " & CDbl(.Fields("Jumlah Masuk")) & ", " & CDbl(.Fields("Jumlah Keluar")) & ", " & CCur(.Fields("Harga")) & ", N'" & .Fields("Doc Reff") & "')") = True Then
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
 RcDetail.DBOpen " SELECT     [Detail MutasiGudang].WareHouse AS [Kode Gudang], [Detail MutasiGudang].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang],                        [Detail MutasiGudang].[Jumlah Masuk], [Detail MutasiGudang].[Jumlah Keluar], [Detail MutasiGudang].Harga ,[Detail MutasiGudang].[Doc Reff] FROM         [Detail MutasiGudang] INNER JOIN                       Inventory ON [Detail MutasiGudang].NoItem = Inventory.NoItem WHERE     ([Detail MutasiGudang].NoMutasi = N'" & ParamString & "') ORDER BY [Detail MutasiGudang].NoItem", CNN
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner() As Boolean
RcPartner.DBOpen "SELECT     Inventory.ItemName AS [Nama Barang], Inventory.NoItem AS [Kode Barang], SUM([Inventory Tabel].QTY_IN) AS [Qty Stok],  WareHouse.[WareHouse Name] AS [Nama Gudang], Inventory.WareHouse AS [Kode Gudang] FROM         [Inventory Tabel] LEFT OUTER JOIN Inventory INNER JOIN WareHouse ON Inventory.WareHouse = WareHouse.WareHouse ON [Inventory Tabel].NoItem = Inventory.NoItem WHERE     ([Inventory Tabel].TypeTrans <> N'AR') GROUP BY WareHouse.[WareHouse Name], Inventory.NoItem, Inventory.ItemName, Inventory.WareHouse ORDER BY WareHouse.[WareHouse Name]", CNN, lckLockReadOnly
If RcPartner.Recordcount <> 0 Then
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.FromTagActive = "MASTER BARANG"
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If

End Function

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) <> N'DN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
End With
RcCek.CloseDB
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

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

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "MG-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Function HppProce(ByVal NoItem As String) As Double
Dim RcHpp As New DBQuick
RcHpp.DBOpen "SELECT     RefTrans, HPP FROM         [Inventory Tabel] WHERE     (LockFIFO = 0) AND (TypeTrans = N'AP') AND (NoItem = N'" & NoItem & "') AND (StockTmp <> 0)", CNN, lckLockReadOnly
With RcHpp
     If .Recordcount <> 0 Then
        HppProce = IIf(Not IsNull(.Fields(1)), .Fields(1), 0)
        mVarDocReff = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     Else
        HppProce = 0
        mVarDocReff = ""
     End If
     
End With
RcHpp.CloseDB
End Function

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1514.835
DGPurchase.Columns(1).width = 1649.764
DGPurchase.Columns(2).width = 3060.284
DGPurchase.Columns(3).width = 1244.976
DGPurchase.Columns(4).width = 1244.976
End Sub
