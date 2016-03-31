VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{594F23A7-88F5-4C02-866B-8E877A62F75C}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmProsesBahan 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10365
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10365
   Tag             =   "Proses Pengolahan Bahan Baku"
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
      Height          =   5055
      Left            =   75
      ScaleHeight     =   5025
      ScaleWidth      =   10155
      TabIndex        =   7
      Top             =   0
      Width           =   10185
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00EAAF6F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4530
         Left            =   135
         ScaleHeight     =   4470
         ScaleWidth      =   9765
         TabIndex        =   8
         Top             =   300
         Width           =   9825
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "Doc Reff"
            Height          =   315
            Index           =   4
            Left            =   4530
            TabIndex        =   4
            Tag             =   "PO"
            Top             =   1425
            Width           =   2085
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "UOM"
            Height          =   315
            Index           =   2
            Left            =   2235
            MaxLength       =   25
            TabIndex        =   2
            Tag             =   "PO"
            Top             =   1095
            Width           =   4380
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "QTY Order"
            Height          =   315
            Index           =   3
            Left            =   2235
            MaxLength       =   3
            TabIndex        =   3
            Tag             =   "PO"
            Top             =   1425
            Width           =   1395
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "ItemName"
            Height          =   315
            Index           =   1
            Left            =   2235
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "PO"
            Top             =   765
            Width           =   4380
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "Kode Bahan"
            Height          =   315
            Index           =   0
            Left            =   2235
            MaxLength       =   15
            TabIndex        =   0
            Tag             =   "PO"
            Top             =   435
            Width           =   4380
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Height          =   1950
            Left            =   120
            TabIndex        =   5
            Tag             =   "Partner"
            Top             =   1905
            Width           =   9570
            _ExtentX        =   16880
            _ExtentY        =   3440
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "Kode Barang"
               Caption         =   "No Barang"
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
               DataField       =   "QTY Used"
               Caption         =   "QTY"
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
               DataField       =   "Harga"
               Caption         =   "Harga"
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
               MarqueeStyle    =   3
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
         Begin VB.Line Line1 
            Index           =   4
            X1              =   3690
            X2              =   5880
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   270
            X2              =   2460
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   270
            X2              =   2400
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Direncanakan                                 Doc Reff"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   270
            TabIndex        =   6
            Top             =   1470
            Width           =   4155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   270
            TabIndex        =   13
            Top             =   1147
            Width           =   330
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   270
            X2              =   2445
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang Diproduksi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   270
            TabIndex        =   12
            Top             =   810
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang Diproduksi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   270
            TabIndex        =   11
            Top             =   480
            Width           =   1920
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   7
            Left            =   5835
            TabIndex        =   10
            Top             =   4065
            Width           =   795
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   7455
            TabIndex        =   9
            Top             =   4035
            Width           =   2235
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   270
            X2              =   2610
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   5835
            X2              =   9180
            Y1              =   4305
            Y2              =   4305
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   5400
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmProsesBahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private clsMytr                                           As New DBQuick
Private RcUang                                            As New DBQuick
Private RcDetail                                          As New DBQuick
Private RcPartner                                         As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private mEdit, mEditPO, mFirstCaller, mVarDetailPOClose   As Boolean
Private mAccount                                          As String



'
'Private Sub CboBayar_KeyDown(KeyCode As Integer, Shift As Integer)
'KeyEnter KeyCode
'End Sub
'
'Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
'KeyEnter KeyCode
'End Sub
'
'Private Sub CboUang_KeyDown(KeyCode As Integer, Shift As Integer)
'KeyEnter KeyCode
'End Sub
'
'Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
'KeyEnter KeyCode
'End Sub
'
'Private Sub cmdLink_Click(Index As Integer)
' OpenPartner Index
'End Sub
'
'Private Sub cmdLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF4 Then OpenPartner Index
'End Sub
'
'Private Sub DGPurchase_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'If mEdit = True Then
'   If ColIndex = 3 Then
'      If IsStatusPO(MyDDE.ChildRecordset.Fields("NoItem")) = True Then
'         MessageBox "Kode Barang " & vbCrLf & MyDDE.ChildRecordset.Fields("NoItem") & vbCrLf & "tidak bisa diedit,karena barang sudah dikirim Oleh Supplier " & vbCrLf & lblBank(0) & vbCrLf & " dan telah diterima bagian gudang.", "Peringatan", msgOkOnly
'         DGPurchase.AllowUpdate = False
'         DGPurchase.Columns(ColIndex).Value = MyDDE.ChildRecordset.Fields("QTYPO")
'         mVarDetailPOClose = True
'      Else
'         DGPurchase.AllowUpdate = True
'      End If
'   End If
'End If
'End Sub
'
'Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
'DataError = 0
'Response = 0
'End Sub
'
'Private Sub DTPicker1_Change()
'If mEdit = True Then If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.Fields("ScheduleDate").Value = DTPicker1.Value + CDbl(txtBox(1))
'End Sub
'
'Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then KeyEnter KeyCode
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'ScanKey KeyCode, Shift, MyDDE
'End Sub
'
Private Sub Form_Load()
CenterForm Picture2, Me
HiasForm Picture1, Me
Set mCall = New frmCaller
With MyDDE
     .EditModeReplace = False
     Set .BindForm = FrmProsesBahan
     .BindFormTAG = "PO"
     Set .ActiveConnection = Cnn
     .PrepareQuery = "SELECT     * FROM         [Tabel Bahan Baku] ORDER BY [Kode Bahan]"
End With
End Sub
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
RcUang.CloseDB
clsMytr.CloseDB
Set mCall = Nothing
End Sub
'
'Private Sub Form_Resize()
'On Error Resume Next
'
'Err.Clear
'End Sub
'
Private Sub Form_Unload(Cancel As Integer)
Set FrmProsesBahan = Nothing
End Sub
'
Private Sub mCall_BeforeUnload()
On Error Resume Next
Select Case mCall.FromTagActive
       Case "MASTER BARANG":
            If FindOwnRecordset(MyDDE.ChildRecordset, "[Kode Barang] = '" & mAccount & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Kode Barang") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            End If
'            mFirstCaller = False
            DGPurchase.SetFocus
'       Case "MASTER BANK":
'            CboUang.SetFocus
'       Case "MASTER SUPPLIER":
'            txtBox(1).SetFocus
End Select
End Sub
'
'Private Sub mCall_CallLinkForm()
'If mCall.FromTagActive <> "MASTER BARANG" Then
'   frmMasterSup.SetFocus
'   frmMasterSup.ZOrder (0)
'Else
'   FrmItemData.SetFocus
'   FrmItemData.ZOrder (0)
'End If
'End Sub
'
Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
Select Case TagForm:

       Case "MASTER BARANG":
            MyDDE.ChildRecordset.Fields(0) = mCall.GetFieldByName(0)
            MyDDE.ChildRecordset.Fields(1) = mCall.GetFieldByName(1)
            MyDDE.ChildRecordset.Fields(2) = mCall.GetFieldByName(2)
            MyDDE.ChildRecordset.Fields(3) = 1
            MyDDE.ChildRecordset.Fields(4) = mCall.GetFieldByName(4)
            mAccount = mCall.GetFieldByName(0)
End Select
End If
End Sub
'
Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
'Dim I As Integer
'Dim mStok As Long
'Dim mTmp As Variant
'Select Case ColIndex
'       Case 3, 4, 5:
''            If CBool(IIf(Not IsNull(MyDDE.ChildRecordset.Fields("StatusTrans")), MyDDE.ChildRecordset.Fields("StatusTrans"), False)) = False Then
'               If CDbl(DGPurchase.Columns(ColIndex).Value) <> 0 Then
'                  mTmp = (DGPurchase.Columns(3) * DGPurchase.Columns(4)) * (DGPurchase.Columns(5) / 100) + (DGPurchase.Columns(3) * DGPurchase.Columns(4))
'                  DGPurchase.Columns(7).Value = mTmp
'               Else
'                  mTmp = (DGPurchase.Columns(3) * DGPurchase.Columns(4))
'                  DGPurchase.Columns(7).Value = mTmp
'               End If
''            Else
''               MessageBox "Data Tidak Bisa Diedit Karena Digunakan Oleh Receive Notes Transaksi", "Peringatan", msgOkOnly
''               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
''            End If
'End Select
HitungTotal
End Sub
'
'Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
'If mEdit = False Then Exit Sub
'Call Form_KeyDown(KeyCode, Shift)
'End Sub
'
Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
   Exit Sub
End If
With DGPurchase
     Select Case .Col
            Case 3:
                DGPurchase.MarqueeStyle = dbgFloatingEditor
                .AllowUpdate = True
'            Case 3:
'                If IsDetailOK(MyDDE.ChildRecordset.Fields("NoItem")) = True Then
'                   DGPurchase.MarqueeStyle = dbgHighlightRow
'                   .AllowUpdate = False
''                   MessageBox "Kode Barang " & MyDDE.ChildRecordset.Fields("NoItem") & vbCrLf & " tidak bisa diedit,karena barang sudah dikirim Oleh Supplier " & vbCrLf & lblBank(0) & vbCrLf & " dan telah diterima oleh bagian gudang.", "Peringatan", msgOkOnly
'                Else
'                   DGPurchase.MarqueeStyle = dbgFloatingEditor
'                   .AllowUpdate = True
'                End If
            Case Else:
                            DGPurchase.MarqueeStyle = dbgHighlightRow
                .AllowUpdate = False

     End Select
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub
'
Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit, tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = False
               If MyDDE.CancelTrans = True Then MessageBox "Transaksi PO Tidak Bisa Diedit.Karena Transaksi sudah divalidasi."
            End If
            'DGPurchase.Columns(0).Button = False
       Case tmbDetail:
            
            If MyDDE.CancelTrans = False Then
                If MyData.CheckGridKosong(MyDDE.ChildRecordset, "fldtotal") = True Then
                   MyDDE.CancelTrans = True
                   MessageBox "Data transaksi belum lengkap." & "Silahkan dicek kembali.", "Peringatan", msgOkOnly
                Else
                   MyDDE.CancelTrans = mFirstCaller
                End If
            Else
               'MessageBox "Tidak bisa menambah detail PO ,karena barang sudah dikirim Oleh Supplier " & lblBank(0) & " dan telah diterima bagian gudang.", "Peringatan", msgOkOnly
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
'               If CekGridKosong = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
'                  'MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                  PrepareQuery
'               Else
'                  MyDDE.IsChildMemberReady = False
'               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub
'
Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'On Error Resume Next
'txtBox(0).Enabled = False
'lblBank(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            mEdit = True
            Text1(0).Enabled = False
            Text1(1).SetFocus
       Case tmbAddNew:
            mEdit = True
            Text1(0).SetFocus
            Text1(3) = 0
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               mEdit = False
            End If

       Case tmbCancel:
'            If MyDDE.ChildRecordset.Recordcount = 0 Then
'               mEdit = False
'               DGPurchase.Columns(6).Visible = True
'               DGPurchase.Columns(7).Visible = False
'               If Me.Caption = "P.O Transaksi" Then chkPo.Enabled = False
'               mVarDetailPOClose = False
'             Else
'               DGPurchase.Columns(6).Visible = False
'               DGPurchase.Columns(7).Visible = True
'               'mEdit = True
'             End If
       Case tmbDetail:
            If mFirstCaller = False Then
               OpenPartner 0
'               DGPurchase.Columns(6).Visible = False
'               DGPurchase.Columns(7).Visible = True
'               mEdit = True
'               mVarDetailPOClose = False
            End If
       Case tmbPrint:
            CallRPTReport "Pengolahan BB.rpt", "Select * From [Pengolahan BB] where [Kode Bahan] ='" & Text1(0) & "'"
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select
'CmdLink(0).Enabled = mEdit
'CmdLink(1).Enabled = mEdit
'cboType.Enabled = mEdit
Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("Kode Bahan")
HitungTotal
End Sub
'
Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:

Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT     Inventory.NoItem AS [No Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM, SUM([Inventory Tabel].QTY_IN)   - SUM([Inventory Tabel].QTY_OUT) AS Stok, MAX([Inventory Tabel].PriceIn) AS Harga FROM         Inventory INNER JOIN" & _
                             " [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem GROUP BY Inventory.NoItem, Inventory.ItemName, Inventory.UOM HAVING      (SUM([Inventory Tabel].QTY_IN) - SUM([Inventory Tabel].QTY_OUT) > 0) ORDER BY Inventory.NoItem", Cnn, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen MyData.UploadQuery("BANK", MyDDE.GetFieldByName("PartnerID")), Cnn, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT [Remainder PO].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100)   + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", Cnn, lckLockReadOnly
       Case 3:
'            RcPartner.DBOpen "SELECT     NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn FROM         Inventory WHERE     (Manufacture = " & cboType.ListIndex & ") ORDER BY NoItem", Cnn, lckLockReadOnly
'            mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "MASTER BARANG"
'            mCall.txtCari = lblBank(0)
'            mCall.CaptionLink = "Supplier"
          Case 1:
            mCall.FromTagActive = "MASTER BANK"
'            mCall.txtCari = lblBank(1)
          Case 2:
            mCall.FromTagActive = "REMINDER"
'            mCall.txtCari = lblBank(1)
          Case 3:
            mCall.FromTagActive = "MASTER BARANG"
            mCall.CaptionLink = "Barang"
            If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me

Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub
'
Private Sub OpenDetail(ByVal ParameterString As String)
Set RcDetail = New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
RcDetail.DBOpen " SELECT     [Detail Bahan Baku].NoItem AS [Kode Barang], Inventory.ItemName AS [Nama Barang], Inventory.UOM, [Detail Bahan Baku].[QTY Used],    [Detail Bahan Baku].PriceList AS Harga FROM         Inventory INNER JOIN  [Detail Bahan Baku] ON Inventory.NoItem = [Detail Bahan Baku].NoItem WHERE     ([Detail Bahan Baku].[Kode Bahan] = N'" & ParameterString & "')", Cnn, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset

End Sub
'
Private Sub SimpanDetail()
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [Detail Bahan Baku] WHERE     ([Kode Bahan] = N'" & Text1(0) & "')") = True Then
           Do
              If .EOF = True Then Exit Do
              SendDataToServer " INSERT INTO [Detail Bahan Baku]" & _
                               " ([Kode Bahan], NoItem, [QTY Used], PriceList)" & _
                               " VALUES     (N'" & Text1(0) & "', N'" & .Fields("Kode Barang") & "', " & .Fields("QTY USED") & ", " & .Fields("Harga") & ")"
              .MoveNext
           Loop
           End If
           .MoveLast
           DGPurchase.Refresh
     End If
End With
End Sub
'
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub
'
'Private Sub txtBox_Change(Index As Integer)
'If Index = 2 And mEdit = True Then
'   If txtBox(Index) = "" Then txtBox(Index) = 0
'   If CInt(txtBox(Index)) > 100 Then txtBox(Index) = 0
'   MyDDE.GetFieldByName("Discount") = txtBox(Index)
'   HitungTotal
'ElseIf Index = 1 And mEdit = True Then
'   If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.Fields("ScheduleDate").Value = DTPicker1.Value + CDbl(txtBox(Index))
'End If
'End Sub
'
'Private Sub txtBox_GotFocus(Index As Integer)
'Block txtBox(Index)
'End Sub
'
'Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then KeyEnter KeyCode
'End Sub
'
'Private Function TglIndex() As String
'Dim TglHari, TglBulan, TglTahun As String
'TglIndex = "PO/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
'End Function
'
Private Sub HitungTotal()
On Error Resume Next
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mTotal  As Variant

Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotal = 0
With RcTotal
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        ' 3 = QTY  4 = Harga
        For I = 0 To UBound(Avdata, 2)
            mTotal = mTotal + Avdata(3, I) * Avdata(4, I)
        Next I
     Else
        mTotal = 0
     End If
End With
LblAmount(0) = FormatNumber(mTotal, 0)
RcTotal.CloseDB
Set Avdata = Nothing
Set mTotal = Nothing
Set RcTotal = Nothing
Err.Clear
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Tabel Bahan Baku]" & _
                     " ([Doc Reff],[Kode Bahan], ItemName, UOM, [QTY Order], Ordered)" & _
                     " VALUES     (N'" & Text1(4) & "',N'" & Text1(0) & "', N'" & Text1(1) & "', N'" & Text1(2) & "', " & Text1(1) & ", 0)"

    .PrepareUpdate = " UPDATE    [Tabel Bahan Baku]" & _
                     " Set [Doc Reff]=N'" & Text1(4) & "', ItemName = N'" & Text1(1) & "', UOM = N'" & Text1(2) & "', [QTY Order] = " & Text1(3) & " WHERE     ([Kode Bahan] = N'" & Text1(0) & "')"

    .PrepareDelete = " DELETE FROM  [Tabel Bahan Baku] WHERE ([Kode Bahan] = N'" & Text1(0) & "')"
End With
Err.Clear
End Sub
'
'Private Function IsHeaderOk(ByVal NoPo As String) As Boolean
'Dim RcIs As New DBQuick
'RcIs.DBOpen "SELECT  StatusSJ FROM [PO Order] WHERE     (PurchaseID = N'" & NoPo & "')", Cnn, lckLockReadOnly
'IsHeaderOk = False
'With RcIs
'     If .Recordcount <> 0 Then IsHeaderOk = CBool(.Fields(0))
'End With
'RcIs.CloseDB
'End Function
'
'Private Function IsStatusPO(Optional ByVal NoItem As String) As Boolean
'Dim RcIs As New DBQuick
'If NoItem = "" Then
'   RcIs.DBOpen "SELECT SUM(QTY_Receive) AS QTY FROM [Detail TransData] WHERE     (DNID = N'" & txtBox(0) & "')", Cnn, lckLockReadOnly
'Else
'   RcIs.DBOpen "SELECT     QTY_Receive AS QTY FROM         [Detail TransData] WHERE     (DNID = N'" & txtBox(0) & "') AND (NoItem = N'" & NoItem & "')", Cnn, lckLockReadOnly
'End If
'With RcIs
'     If .Recordcount <> 0 Then If .Fields(0) <> 0 Then IsStatusPO = True
'End With
'RcIs.CloseDB
'End Function
'
''Private Function IsDetailOK(ByVal Noitem As String) As Boolean
''Dim RcIs As New DBQuick
''RcIs.DBOpen "SELECT     [Detail PO].StatusTrans FROM         [Detail PO] INNER JOIN                       [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE     ([PO Order].PurchaseID = N'" & txtBox(0) & "') AND ([Detail PO].NoItem = N'" & Noitem & "') GROUP BY [Detail PO].StatusTrans HAVING      ([Detail PO].StatusTrans = 1)", Cnn, lckLockReadOnly
''With RcIs
''     If .Recordcount <> 0 Then IsDetailOK = CBool(.Fields(0))
''End With
''RcIs.CloseDB
''Set RcIs = Nothing
''End Function
'
'Private Sub OpenTypeBayarPO()
'clsMytr.DBOpen MyData.UploadQuery("franco beli"), Cnn, lckLockReadOnly
'Set CboBayar.RowSource = clsMytr.DBRecordset
'End Sub
'
'Private Sub MataUang()
'RcUang.DBOpen MyData.UploadQuery("mata uang"), Cnn, lckLockReadOnly
'Set CboUang.RowSource = RcUang.DBRecordset
'End Sub
'
'Private Sub UpdateTotal()
'Dim rcUpdate As New DBQuick
'Dim iLast, mRow As Integer
'Dim Avdata As Variant
'Set rcUpdate.DBRecordset = MyDDE.ChildRecordset.Clone(adLockBatchOptimistic)
'With rcUpdate
'     If .Recordcount <> 0 Then
'        mRow = MyDDE.ChildRecordset.AbsolutePosition
'        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
'        For iLast = 0 To UBound(Avdata, 2)
'            .AbsolutePosition = iLast + 1
'            .Fields("Tmp") = Avdata(7, iLast)
'        Next iLast
'     End If
'End With
'Set MyDDE.ChildRecordset = rcUpdate.DBRecordset.Clone(adLockBatchOptimistic)
'If MyDDE.ChildRecordset.Recordcount <> 0 Then
'   MyDDE.ChildRecordset.AbsolutePosition = mRow
'End If
'rcUpdate.CloseDB
'End Sub
'
'Private Function CekDetailItem(ByVal PoNumber As String, ByVal NoItemData As String) As Boolean
'Dim RcCek As New DBQuick
'RcCek.DBOpen "SELECT NoItem, PurchaseID FROM [Detail PO] WHERE     (NoItem = N'" & NoItemData & "') AND (PurchaseID = N'" & PoNumber & "')", Cnn, lckLockReadOnly
'If RcCek.Recordcount <> 0 Then CekDetailItem = True
'RcCek.CloseDB
'End Function
'
'Private Sub ListTotalDeliver(ByVal ParamString As String)
'Dim RcDN As New DBQuick
'If ParamString = "" Then ParamString = "XXXXX"
'RcDN.DBOpen "SELECT DateTrans FROM TransData GROUP BY DateTrans, PurchaseID HAVING      (PurchaseID = N'" & ParamString & "')", Cnn, lckLockReadOnly
'With RcDN
'     If .Recordcount <> 0 Then
'        LblDeliVer = Abs(CDate(Format(DTPicker1.Value, "dd/mm/yyyy")) - CDate(Format(.Fields(0), "dd/mm/yyyy")))
'     Else
'        LblDeliVer = 0
'     End If
'End With
'End Sub
'
'Private Function CekGridKosong() As Boolean
'Dim RcKsg As New DBQuick
'Dim Avdata As Variant
'Dim I As Integer
'Dim Temp As String
'Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
'With RcKsg
'     If .Recordcount <> 0 Then
'        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
'        For I = 0 To UBound(Avdata, 2)
'            Temp = IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), "")
'            If Temp <> "" Then
'                If Val(Avdata(3, I)) = 0 Or Val(Avdata(4, I)) = 0 Then
'                   MessageBox "Data Item Untuk Quantity Atau Harga Ada Yang Berisi NOl.Harap Dicek Dulu", "Peringatan"
'                   CekGridKosong = True
'                   Exit For
'                End If
'            Else
'               MessageBox "Data Item Tidak Lengkap.Harap Dicek Dulu", "Peringatan"
'               CekGridKosong = True
'               Exit For
'            End If
'        Next I
'     Else
'        CekGridKosong = True
'     End If
'End With
'RcKsg.CloseDB
'End Function
'
'Private Function CekStock(ByVal NoItem As String) As Long
'Dim RcCek As New Recordset
'RcCek.CursorLocation = adUseClient
'RcCek.Open "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) = N'RN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
'With RcCek
'     If .Recordcount <> 0 Then
'        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
'     Else
'        CekStock = 0
'     End If
'     .Close
'End With
'Set RcCek = Nothing
'End Function
'
'Private Sub CekBankName(ByVal PartnerId As String, ByVal NoRekening As String)
'Dim RcBnk As New DBQuick
'RcBnk.DBOpen "SELECT     Account, [Bank Name] FROM         [Bank Partner] WHERE     (PartnerID = N'" & PartnerId & "') AND (Account = N'" & NoRekening & "')", Cnn, lckLockReadOnly
'With RcBnk
'     If .Recordcount <> 0 Then
'         lblBank(1) = .Fields(1)
'     Else
'         lblBank(1) = ""
'     End If
'End With
'RcBnk.CloseDB
'End Sub
'
'
'
Private Sub Picture2_Click()

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then
   ValidNum KeyAscii
End If
End Sub
