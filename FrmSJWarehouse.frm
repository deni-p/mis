VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmSJWarehouse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pengeluaran Barang"
   ClientHeight    =   6105
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "FrmSJWarehouse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   5535
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1005
      BindFormTAG     =   "SPP"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   11505
      TabIndex        =   12
      Top             =   0
      Width           =   11505
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "gudang_tujuan"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   4
         Left            =   6840
         MaxLength       =   35
         TabIndex        =   7
         Tag             =   "SPP"
         Top             =   510
         Width           =   2955
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "gudang_asal"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   3
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "SPP"
         Top             =   150
         Width           =   2955
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Person"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   2
         Left            =   1590
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1245
         Width           =   2415
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "No Pol"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   1
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   2
         Top             =   885
         Width           =   2415
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "TransID"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1590
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "SPP"
         Top             =   150
         Width           =   2400
      End
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   9795
         Picture         =   "FrmSJWarehouse.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   158
         Width           =   330
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   5
         Left            =   6840
         MaxLength       =   200
         TabIndex        =   9
         Tag             =   "SPP"
         Top             =   1245
         Width           =   4515
      End
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   9795
         Picture         =   "FrmSJWarehouse.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   518
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   3615
         Left            =   60
         TabIndex        =   10
         Top             =   1830
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
            DataField       =   "DNID"
            Caption         =   "No Permintaan"
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
            DataField       =   "noItem"
            Caption         =   "Kode Brg"
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
            DataField       =   "InternalName"
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
            DataField       =   "UOM"
            Caption         =   "Satuan"
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
            DataField       =   "Qty_IN"
            Caption         =   "Jml Stock"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ActQty"
            Caption         =   "Jml Permintaan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Qty_Receive"
            Caption         =   "Jml Dikirim"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "sak"
            Caption         =   "Jml Coli"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "supplier"
            Caption         =   "Supplier"
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
            Size            =   248
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1590
         TabIndex        =   1
         Tag             =   "SPP"
         Top             =   510
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   64684035
         CurrentDate     =   38272
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "Jam"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   6840
         TabIndex        =   8
         Tag             =   "SPP"
         Top             =   870
         Width           =   1995
         _ExtentX        =   3519
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
         Format          =   64684034
         CurrentDate     =   36494
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         Height          =   270
         Index           =   0
         Left            =   4470
         TabIndex        =   21
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   5445
         TabIndex        =   20
         Top             =   938
         Width           =   465
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   5400
         X2              =   6900
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   5445
         TabIndex        =   19
         Top             =   578
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5400
         X2              =   6900
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   5445
         TabIndex        =   18
         Top             =   218
         Width           =   285
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   5400
         X2              =   6900
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sopir"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   17
         Top             =   1313
         Width           =   360
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   120
         X2              =   1620
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pol"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   16
         Top             =   953
         Width           =   510
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   1620
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   120
         X2              =   1620
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   1620
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   578
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pengiriman"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   218
         Width           =   1080
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   5400
         X2              =   6900
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   5445
         TabIndex        =   13
         Top             =   1313
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmSJWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDetail                                          As New DBQuick
Private RcPartner                                         As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private MEdit, mEditPO, mFirstCaller, mVarDetailPOClose   As Boolean
Private mAccount                                          As String
Dim strSQL As String


Private Sub cmdLink_Click(Index As Integer)
   OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
   If (ColIndex = 6) Or (ColIndex = 7) Then
      On Error GoTo xErr
      If (DGPurchase.Columns(ColIndex) = 0) Then
         MessageBox "Nilai Tidak Boleh Nol", "Peringatan"
         If ColIndex = 6 Then
            DGPurchase.Columns(ColIndex) = IIf(Val(DGPurchase.Columns(5)) <= Val(DGPurchase.Columns(4)), DGPurchase.Columns(5), DGPurchase.Columns(4))
         Else
            DGPurchase.Columns(ColIndex) = 1
         End If
      End If
      If (DGPurchase.Columns(6) > DGPurchase.Columns(4)) Then
         MessageBox "Stok yg ada Hanya " & DGPurchase.Columns(4) & " !", "Peringatan"
         DGPurchase.Columns(6) = DGPurchase.Columns(4)
      End If
      If (DGPurchase.Columns(6) > DGPurchase.Columns(5)) Then
         MessageBox "Stok yg Diminta cuma " & DGPurchase.Columns(5) & " !", "Peringatan"
         DGPurchase.Columns(6) = IIf(Val(DGPurchase.Columns(5)) <= Val(DGPurchase.Columns(4)), DGPurchase.Columns(5), DGPurchase.Columns(4))
      End If
   End If
Exit Sub
xErr:
   MessageBox "Isikan Dengan Angka ", "Peringatan"
   DGPurchase.Columns(ColIndex) = 1
   Err.Clear
End Sub

Private Sub DGPurchase_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If MEdit = True Then
    DGPurchase.AllowUpdate = True
End If
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
'DataError = 0
'Response = 0
End Sub


Private Sub Dir1_Change()

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()

HiasFormManTell Picture2, Me

mVarDetailPOClose = False
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
With MyDDE
     .EditModeReplace = False
     Set .BindForm = Me
     Set .ActiveConnection = CNN
     .PrepareQuery = "SELECT * from TransData where status = 0 and TypeTrans='SS'"
     strSQL = "SELECT transdata.TransID, transdata.EmpID, transdata.DateTrans, transdata.DateIssued, " & _
            " transdata.Status, transdata.RefNotes, transdata.TypeTrans, transdata.[No Pol], " & _
            " transdata.WareHouse, transdata.jam, transdata.tujuan, transdata.person, " & _
            " WareHouse_1.[WareHouse Name] AS gudang_asal, WareHouse.[WareHouse Name] AS gudang_tujuan " & _
            " FROM transdata LEFT OUTER JOIN WareHouse ON transdata.tujuan = WareHouse.WareHouse " & _
            " LEFT OUTER JOIN WareHouse AS WareHouse_1 ON transdata.WareHouse = WareHouse_1.WareHouse " & _
            " WHERE (transdata.TypeTrans = N'SS')"
            
    .PrepareQuery = strSQL
     .SetPermissions = aksess.MayDo("Pengiriman RL")
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
Set mCall = Nothing
End Sub


Private Sub mCall_BeforeUnload()
On Error Resume Next
            If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Then
                  If MyDDE.ChildRecordset.Fields(0) = "" Then
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  End If
               End If
            End If
            mFirstCaller = False
            If DGPurchase.Enabled = True Then
               DGPurchase.AllowUpdate = True
               DGPurchase.col = 3
               DGPurchase.SetFocus
            End If
   
End Sub

Private Sub mCall_CallLinkForm()
If mCall.FromTagActive = "Inventory List" Then
   FrmItemData.SetFocus
   FrmItemData.ZOrder (0)
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Dim rsSupplier As New DBQuick
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
Select Case TagForm:
          
          
     Case "Daftar Penerimaan"
     'Case "Inventory List":
          'Field QTY_in  DIISI dengan  -> Jml Stock di Gudang
          'Field QTY_OUt DIISI dengan  -> Qty yg diminta
          'Field ActQty  DIISI dengan  -> QTY yg diminta setelah dikurangi brg yg pernah dikirim (aktual)
          
          MyDDE.ChildRecordset.Fields("TransID") = MyDDE.GetFieldByName("TransID")
          MyDDE.ChildRecordset.Fields("DNID") = mCall.GetFieldByName("No Permintaan")
          MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("Kode brg")
          'MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("Nama Barang")
          MyDDE.ChildRecordset.Fields("InternalName") = mCall.GetFieldByName("Nama Barang")
          MyDDE.ChildRecordset.Fields("UOM") = mCall.GetFieldByName("Satuan")
          MyDDE.ChildRecordset.Fields("QTy_in") = GetSOH(mCall.GetFieldByName("Kode Brg"))
          MyDDE.ChildRecordset.Fields("ActQty") = mCall.GetFieldByName("jml")
          MyDDE.ChildRecordset.Fields("Qty_Out") = mCall.GetFieldByName("QtyRequest")
          MyDDE.ChildRecordset.Fields("Qty_Receive") = IIf(Val(MyDDE.ChildRecordset.Fields("ActQty")) <= Val(MyDDE.ChildRecordset.Fields("QTy_in")), MyDDE.ChildRecordset.Fields("ActQty"), MyDDE.ChildRecordset.Fields("QTy_in"))
          MyDDE.ChildRecordset.Fields("sak") = 1
           
          rsSupplier.DBOpen "SELECT  TransData.PartnerId, PartnerDB.CompanyName, [inventory Tabel].refTrans " & _
                            "FROM PartnerDB INNER JOIN TransData ON PartnerDB.PartnerID = TransData.PartnerId RIGHT OUTER JOIN " & _
                                 "[Inventory Tabel] ON TransData.TransID = [Inventory Tabel].RefTrans " & _
                            "where [inventory Tabel].noItem='" & mCall.GetFieldByName("Kode brg") & _
                                    "' and [inventory tabel].lockFIFO =0 " & _
                            "order by [inventory tabel].dateTrans", CNN, lckLockReadOnly
                            
         
          If rsSupplier.DBRecordset.Recordcount > 0 Then
             MyDDE.ChildRecordset.Fields("supplier") = IIf(IsNull(rsSupplier.DBRecordset.Fields(1)), "Tidak Terdeteksi", rsSupplier.DBRecordset.Fields(1))
             MyDDE.ChildRecordset.Fields("supplier ID") = IIf(IsNull(rsSupplier.DBRecordset.Fields(0)), " ", rsSupplier.DBRecordset.Fields(0))
             MyDDE.ChildRecordset.Fields("Referense") = IIf(IsNull(rsSupplier.DBRecordset.Fields(2)), " ", rsSupplier.DBRecordset.Fields(2))
          Else
             MyDDE.ChildRecordset.Fields("supplier") = "-"
             MyDDE.ChildRecordset.Fields("supplier ID") = " "
             MyDDE.ChildRecordset.Fields("Referense") = " "
          End If
          
          Set rsSupplier = Nothing
          
    Case "Gudang":
        MyDDE.GetFieldByName("warehouse") = mCall.GetFieldByName("kode")
        MyDDE.GetFieldByName("gudang_asal") = mCall.GetFieldByName("Gudang")
    Case "Gudang Tujuan":
        MyDDE.GetFieldByName("Tujuan") = mCall.GetFieldByName("kode")
        MyDDE.GetFieldByName("gudang_tujuan") = mCall.GetFieldByName("Gudang")
        
End Select
End If
End Sub

Private Function GetSOH(KodeBrg As String) As Double
   Dim rsSOH As New DBQuick
   rsSOH.DBOpen "Select sum(stockTmp) from [inventory Tabel] where (lockFIFO=0) and (NoItem='" & KodeBrg & "') and (lokasiGdg='" & MyDDE.GetFieldByName("WareHouse") & "') group by NoItem", CNN, lckLockReadOnly
   If rsSOH.DBRecordset.Recordcount > 0 Then
      GetSOH = rsSOH.DBRecordset.Fields(0)
   Else
      GetSOH = 0
   End If
   Set rsSOH = Nothing
End Function

Private Sub MergeDoubleItem()
   
End Sub


Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If MEdit = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Exit Sub
End If
With DGPurchase
     Select Case .col
            Case 0, 1, 2:
                .AllowUpdate = False
            Case Else:
                .AllowUpdate = True
     End Select
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
   PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit, tmbDelete:
       Case tmbDetail:
            If txtBox(1) = "" Then MyDDE.GetFieldByName("Note") = "-"
            If MyDDE.CancelTrans = False Then
                If MyData.CheckGridKosong(MyDDE.ChildRecordset, "Qty_SPP") = True Then
                   MyDDE.CancelTrans = True
                   MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly
                End If
            End If
       Case tmbSave:
            If txtBox(1) = "" Then MyDDE.GetFieldByName("Note") = "-"
            If MyDDE.CheckEmptyControl = False Then
               If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
                  MyDDE.GetFieldByName("SPP_Date") = DTPicker1.Value
                  PrepareQuery
                  If txtBox(3).Text = txtBox(4).Text Then
                     MessageBox "Gudang asal dan tujuan tidak boleh sama !", "Peringatan", msgOkOnly, msgCrtical
                     MyDDE.IsChildMemberReady = False
                  End If
               Else
                  MyDDE.IsChildMemberReady = False
                  MyDDE.CancelTrans = True
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim IDGen As New IDGenerator
txtBox(0).Enabled = False
CmdLink(0).Enabled = False
CmdLink(1).Enabled = False

Select Case AdReasonActiveDb
       Case tmbEdit:
            CmdLink(0).Enabled = True
            CmdLink(1).Enabled = True
            MEdit = True
            mEditPO = True
            Call DGPurchase_RowColChange(DGPurchase.row, DGPurchase.col)
            
       Case tmbAddNew:
            MEdit = True
            CmdLink(0).Enabled = True
            CmdLink(1).Enabled = True
            DTPicker1.Value = Now
            DTPicker2.Value = Now
            MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
            MyDDE.GetFieldByName("TransID") = IDGen.GetID("SS")
            MyDDE.GetFieldByName("jam") = DTPicker2.Value
            MyDDE.GetFieldByName("status") = False
            MyDDE.GetFieldByName("refNotes") = "-"
            DTPicker1.SetFocus
            
       Case tmbSave:
'            If MyDDE.IsChildMemberReady = True Then
'               SimpanDetail
'               MEdit = False
'               mEditPO = False
'               OpenDetail txtBox(0)
'               mVarDetailPOClose = False
'               CmdLink(0).Enabled = False
'            Else
'               MessageBox "Detail Item  belum ada datanya.", "Peringatan", msgOkOnly
'            End If

            If MyDDE.IsChildMemberReady = True Then
               Dim aStatus As String
               With MyDDE.ChildRecordset
               If .Recordcount <> 0 Then
                   If SendDataToServer("DELETE FROM  [backflush_line] WHERE     (IDTrans = N'" & txtBox(0) & "')") = True Then
                        .MoveFirst
                        Do
                          If .EOF Then Exit Do
                            ' SendDataToServer (" UPDATE    backflush_line" & _
                                               " Set [Qty Warehouse] = " & CDbl(.Fields("Qty Warehouse")) & _
                                               " WHERE     (IDTrans = N'" & txtBox(0) & "') AND (NoItem = N'" & .Fields("Item ID") & "')")
                            aStatus = IIf(DGPurchase.Columns(4) = DGPurchase.Columns(5), "1", "0")
                            
                            SendDataToServer "update backflush_line set [Qty Required]=" & Val(DGPurchase.Columns(4)) - Val(DGPurchase.Columns(5)) & ",status=" & aStatus & " where idx='" & .Fields("idx") & "'"
                            
                            SendDataToServer (" INSERT INTO backflush_line (IDTrans, StageID,                                                          OrderID,                    NoItem,                       Description,                       UOM,                        Lokasi,                       Cost,                          [Qty Warehouse],                          [Qty Received],                      [Qty Required],                          ResourcesID)" & _
                                                        " VALUES (N'" & txtBox(0) & "', N'" & IIf(IsNull(.Fields("StageID")), "", .Fields("StageID")) & "', N'" & lblSupplier(0) & "', N'" & .Fields("Item ID") & "', N'" & .Fields("Description") & "', N'" & .Fields("UOM") & "', N'" & .Fields("Lokasi") & "',  " & CDbl(.Fields("Cost")) & ",  " & CDbl(.Fields("Qty Warehouse")) & "," & CDbl(.Fields("Qty Received")) & "," & .Fields("Qty Received") & ", N'" & .Fields("ResourcesID") & "')")
                            
                            SendDataToServer "update backflush_header set status=1 where IDTrans='" & .Fields("Description") & "'"
                            
                            If lblSupplier(0).Caption = "" Then
                              'UpdateStock .Fields("Item ID"), Val(DGPurchase.Columns(5))
                            Else
                              SendARItem .Fields("Item ID"), CDbl(.Fields("Qty Warehouse")), CDbl(.Fields("Cost")), txtBox(0), DTPicker1.Value, CDbl(.Fields("Cost")), "MR", True
                              'SendQTY .Fields("Item ID")
                            End If
                          .MoveNext
                        Loop
                        .MoveFirst
                        'ClosedPO
                   End If
               End If
               End With
            End If

            
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               MEdit = False
               mVarDetailPOClose = False
             End If
             
       Case tmbDetail:
               OpenPartner 3
               MEdit = True
               mVarDetailPOClose = False
               
       Case tmbPrint:
               Dim aReport As New utility
               aReport.CallReportView "select * from ReportSjGudang where TransID='" & MyDDE.GetFieldByName("TransID") & "'", "ReportSJWarehouse.rpt", ReportPath, "Surat Penyerahan Barang"
               Set aReport = Nothing
End Select

Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("TransID")
MEdit = False
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen "select warehouse as Kode,[Warehouse Name] as Gudang ,locations as lokasi from WareHouse", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen "select warehouse as Kode,[Warehouse Name] as Gudang ,locations as lokasi from WareHouse", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT [Remainder PO].NoItem, Inventory.InternalName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100) + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", CNN, lckLockReadOnly
       Case 3:
            'field QtyRequest -> adalah jml Awal Barang yg diminta
            'Field Jml        -> adalah sisa jml Barang yg diminta
            
           RcPartner.DBOpen "SELECT backflush_header.DateTrans AS Tanggal, " & _
                                   "backflush_line.Batch_lot as [Batch / Lot] , " & _
                                   "backflush_line.[Qty Required] AS [Quote Qty], " & _
                                   "backflush_line.UOM, " & _
                                   "backflush_header.[Issued BY], " & _
                                   "backflush_line.Cost, " & _
                                   "backflush_line.StageID, " & _
                                   "backflush_line.ResourcesID, " & _
                                   "backflush_line.IDX, " & _
                                   "backflush_header.dept, " & _
                                   "backflush_line.NoItem AS [Item ID], " & _
                                   "backflush_line.Lokasi AS WareHouse, " & _
                                   "backflush_line.penggunaan, " & _
                                   "backflush_line.description, " & _
                                   "backflush_line.IDTRans " & _
                           " FROM Inventory INNER JOIN " & _
                               " backflush_line ON Inventory.NoItem = backflush_line.NoItem INNER JOIN " & _
                               " backflush_header ON backflush_line.IDTrans = backflush_header.IDTrans " & _
                           " WHERE (LEFT(backflush_line.IDTrans, 2) = 'FR') AND (backflush_line.Status <> 1) AND (backflush_header.TypeTrans <> N'MI') AND (backflush_line.[Qty Required] <> 0) and backflush_header.approved_by is not null and left(backflush_line.noItem,2) = 'BB'" & _
                           " ORDER BY backflush_header.DateTrans desc , backflush_header.[Issued BY]", CNN, lckLockReadOnly
            
            'mFirstCaller = True
       Case 4:
            RcPartner.DBOpen "Select Code as Kode, Description as Keterangan,  [Bal_ Account Type], [Bal_ Account No_] from TermMethod ", CNN, lckLockReadOnly
       Case 5:
            RcPartner.DBOpen "Select No_ as Kode, Description as Keterangan, [Gen_ Prod_ Posting Group],  [Tax Group Code], [VAT Prod_ Posting Group], [Search Description], [Global Dimension 1 Code], [Global Dimension 2 Code] from item_charge ", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Gudang"
            mCall.CaptionLink = "Gudang"
          Case 1:
            mCall.FromTagActive = "Gudang Tujuan"
          Case 2:
            mCall.FromTagActive = "Remindier"
          Case 3:
            mCall.FromTagActive = "Daftar Penerimaan"
            mCall.CaptionLink = "Barang"
            'If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
          Case 4:
            mCall.FromTagActive = "Term Method"
            mCall.CaptionLink = "Term Method"
          Case 5:
            mCall.FromTagActive = "Item Charge"
            mCall.CaptionLink = "Item Charge"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
Set RcDetail = New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
'RcDetail.DBOpen "SELECT [Detail TransData].TransID,[Detail TransData].NoItem,[inventory].itemName,inventory.UOM,[Detail TransData].QTY_Out, [Detail TransData].QTY_Receive,[Detail TransData].QTY_IN,[Detail TransData].Referense, [Detail TransData].ActQTY,[Detail TransData].sak FROM [Detail TransData] inner join inventory on [Detail TransData].noItem = inventory.noItem  where [Detail TransData].TransID='" & ParameterString & "'", CNN, lckLockBatch
RcDetail.DBOpen "SELECT * FROM QDetailSJGudang where TransID='" & ParameterString & "'", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Sub SimpanDetail()
Dim jmlSisa As Double
Dim jmlAmbil As Double
Dim SisaStock As Double
Dim rsBalance As New DBQuick
With MyDDE.ChildRecordset

'** Field DNID sebagai referensi No Permintaan **
'** Field Reference sebagai referensi untuk mencari supplier **

     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [detail TransData] WHERE TransID = '" & MyDDE.GetFieldByName("TransID") & "'") = True Then
               While Not .EOF
               '** Simpan ke Detail TransData **
                  jmlSisa = Val(.Fields("ActQty")) - Val(.Fields("Qty_receive"))
                  SendDataToServer " INSERT INTO [detail TransData] (TransID, NoItem, DateTrans, Qty_Receive, Qty_Out,Referense,ActQty,sak,DNID) " & _
                                   " VALUES ('" & MyDDE.GetFieldByName("TransID") & _
                                            "','" & .Fields("NoItem") & _
                                            "','" & Format(MyDDE.GetFieldByName("DateTrans"), "yyyy-MM-dd") & _
                                            "', " & FQty(.Fields("qty_receive")) & _
                                            " , " & FQty(.Fields("qty_out")) & _
                                            " ,'" & .Fields("Referense") & _
                                            "', " & FQty(jmlSisa) & _
                                            " , " & FQty(.Fields("sak")) & _
                                            " ,'" & .Fields("DNID") & "')"
                                            
               '** update stock gudang [inventory tabel]
                   rsBalance.DBOpen "select Qty_Out,stockTmp,noIdx from [Inventory Tabel] where NoItem='" & .Fields("NoItem") & "' and lockFIFO=0 order by DateTrans ", CNN, lckLockReadOnly
                   If rsBalance.DBRecordset.Recordcount > 0 Then
                     jmlAmbil = Val(.Fields("Qty_Receive"))
                     While Not rsBalance.DBRecordset.EOF
                        If jmlAmbil > 0 Then
                           If jmlAmbil >= rsBalance.DBRecordset.Fields(1) Then
                              SendDataToServer "update [inventory tabel] set Qty_Out= Qty_in, StockTmp=0, lockFIFO=1 where NoIdx='" & rsBalance.DBRecordset.Fields("NoIdx") & "'"
                              jmlAmbil = jmlAmbil - Val(rsBalance.DBRecordset.Fields(1))
                           ElseIf jmlAmbil < rsBalance.DBRecordset.Fields(1) Then
                              SisaStock = Val(rsBalance.DBRecordset.Fields(1)) - jmlAmbil
                              SendDataToServer "update [Inventory tabel] set Qty_out = Qty_out + " & jmlAmbil & ", StockTmp = " & SisaStock & " where NoIdx='" & rsBalance.DBRecordset.Fields("NoIdx") & "'"
                              jmlAmbil = 0
                           End If
                           rsBalance.DBRecordset.MoveNext
                        Else
                           rsBalance.DBRecordset.MoveLast
                        End If
                     Wend
                   End If
                   Set rsBalance = Nothing
                  .MoveNext
               Wend
           End If
           .MoveLast
           DGPurchase.Refresh
     End If
End With
End Sub


Private Sub PrepareQuery()

On Error Resume Next
Dim strSQL As String
With MyDDE
   strSQL = " INSERT INTO  TransData (TransID,EmpID,DateTrans,DateIssued,Status,refNotes,TypeTrans,[No Pol],WareHouse,jam,tujuan,person) " & _
            " Values ('" & .GetFieldByName("TransID") & _
                   "','" & MainMenu.StatusBar1.Panels(1).Text & _
                   "','" & Format(.GetFieldByName("DateTrans"), "yyyy-MM-dd") & _
                   "','" & Format(Now, "yyyy-MM-dd") & _
                   "',0,'" & txtBox(5).Text & _
                   "','SS','" & .GetFieldByName("No Pol") & _
                   "','" & .GetFieldByName("WareHouse") & _
                   "','" & Format(DTPicker2.Value, "yyyy-MM-dd hh:mm:ss") & _
                   "','" & .GetFieldByName("tujuan") & _
                   "','" & .GetFieldByName("person") & "')"
                   
   .PrepareAppend = strSQL
                     
    strSQL = " UPDATE TransData set EmpID    = '" & MainMenu.StatusBar1.Panels(1).Text & _
                               "', DateTrans = '" & Format(.GetFieldByName("DateTrans"), "yyyy-MM-dd") & _
                               "', refNotes   = '" & .GetFieldByName("refNotes") & _
                               "', [No Pol]  = '" & .GetFieldByName("No Pol") & _
                               "', WareHouse = '" & .GetFieldByName("warehouse") & _
                               "', jam       = '" & Format(DTPicker2.Value, "yyyy-MM-dd hh:mm:ss") & _
                               "', Tujuan    = '" & .GetFieldByName("tujuan") & _
                               "', Person    = '" & .GetFieldByName("person") & _
             " where TRansID ='" & .GetFieldByName("TransID") & "'"
             
    .PrepareUpdate = strSQL
                     
    .PrepareDelete = " DELETE FROM  TransData WHERE (TransID = '" & .GetFieldByName("TransID") & "')"
End With
Err.Clear
End Sub

Private Function CekDetailItem(ByVal PoNumber As String, ByVal NoItemData As String) As Boolean
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT NoItem, SPPID FROM QuerySPP WHERE     (NoItem = N'" & NoItemData & "') AND (SPPID = N'" & PoNumber & "')", CNN, lckLockReadOnly
If RcCek.Recordcount <> 0 Then CekDetailItem = True
RcCek.CloseDB
End Function

Private Function CekGridKosong() As Boolean
Dim RcKsg As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Dim Temp As String
Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcKsg
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            Temp = IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), "")
            If (Temp <> "") Or (Temp <> "0") Then
                If Val(IIf(Not IsNull(Avdata(4, I)), Avdata(4, I), 0)) = 0 Then
                   MessageBox "Quantity harus diisi.", "Peringatan"
                   CekGridKosong = True
                   Exit For
                End If
            Else
               MessageBox "Data Item Tidak Lengkap.Harap Dicek Dulu", "Peringatan"
               CekGridKosong = True
               Exit For
            End If
        Next I
     Else
        CekGridKosong = True
     End If
End With
RcKsg.CloseDB
End Function

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
RcCek.Open "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) = N'RN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
     .Close
End With
Set RcCek = Nothing
End Function




