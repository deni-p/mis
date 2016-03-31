VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surat Permintaan Pembelian"
   ClientHeight    =   7725
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   11820
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
   Icon            =   "FrmPRequest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
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
      Height          =   6975
      Left            =   75
      ScaleHeight     =   6945
      ScaleWidth      =   11640
      TabIndex        =   5
      Top             =   45
      Width           =   11670
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00EAAF6F&
         Height          =   6375
         Left            =   120
         ScaleHeight     =   6315
         ScaleWidth      =   11385
         TabIndex        =   6
         Top             =   345
         Width           =   11445
         Begin MSDataListLib.DataCombo cmbDept 
            DataField       =   "kode_dept"
            DataSource      =   "MyDDE"
            Height          =   315
            Left            =   7440
            TabIndex        =   12
            Tag             =   "SPP"
            Top             =   150
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "kode_dept"
            Text            =   "DataCombo1"
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Note"
            DataSource      =   "MyDDE"
            Height          =   1155
            Index           =   1
            Left            =   120
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   9
            Tag             =   "SPP"
            Text            =   "FrmPRequest.frx":6852
            Top             =   5040
            Width           =   11115
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Height          =   3735
            Left            =   120
            TabIndex        =   2
            Top             =   960
            Width           =   11220
            _ExtentX        =   19791
            _ExtentY        =   6588
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "NoItem"
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
               DataField       =   "ItemName"
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
               DataField       =   "uom"
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
               DataField       =   "QTY_SPP"
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
               DataField       =   "keperluan"
               Caption         =   "Keperluan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "#,##0;(#,##0)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Note"
               Caption         =   "Keterangan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "#,##0;(#,##0)"
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
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2505.26
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "SPPID"
            DataSource      =   "MyDDE"
            Height          =   315
            Index           =   0
            Left            =   1620
            MaxLength       =   15
            TabIndex        =   0
            Tag             =   "SPP"
            Top             =   150
            Width           =   3315
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "SPP_DATE"
            DataSource      =   "MyDDE"
            Height          =   330
            Left            =   1620
            TabIndex        =   1
            Tag             =   "SPP"
            Top             =   480
            Width           =   3315
            _ExtentX        =   5847
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
            Format          =   60882947
            CurrentDate     =   38272
         End
         Begin VB.CheckBox chkPo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "P.O Reminder"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3225
            TabIndex        =   4
            Top             =   180
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dept"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   6120
            TabIndex        =   11
            Top             =   210
            Width           =   345
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   6120
            X2              =   7620
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   4800
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   150
            X2              =   1650
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   150
            X2              =   1650
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   8
            Top             =   555
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order No."
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   210
            Width           =   720
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   7155
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1005
      BindFormTAG     =   "SPP"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcDetail                                          As New DBQuick
Attribute RcDetail.VB_VarHelpID = -1
Private RcPartner                                         As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private mEdit, mEditPO, mFirstCaller, mVarDetailPOClose   As Boolean
Private mAccount                                          As String
Private lParams                                           As String
Private RcDept                                            As New DBQuick


Public Property Let IDParams(vData As String)
   lParams = vData
End Property

Private Sub DGPurchase_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If mEdit = True Then
    DGPurchase.AllowUpdate = True
End If
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
'DataError = 0
'Response = 0
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub LoadDept()
   RcDept.DBOpen "select * from dept", CNN, lckLockReadOnly
   Set cmbDept.RowSource = RcDept.DBRecordset
End Sub

Private Sub Form_Load()

CenterForm Picture2, Me
HiasForm Picture1, Me

LoadDept

mVarDetailPOClose = False
Set mCall = New frmCaller
DTPicker1.value = dDateBegin
With MyDDE
     .EditModeReplace = False
     Set .BindForm = Me
     Set .ActiveConnection = CNN
      If lParams = "" Then
         MyDDE.SetReadOnlyMode = False
         .PrepareQuery = " SELECT * from SPP_Header where status = 0 "
      Else
         MyDDE.SetReadOnlyMode = True
         .PrepareQuery = " select * from SPP_Header where SPPID = '" & lParams & "'"
      End If
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
Set mCall = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
   lParams = ""
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
               DGPurchase.Col = 3
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
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
Select Case TagForm:
            
       Case "Inventory List":
            MyDDE.ChildRecordset.Fields("SPPID") = MyDDE.GetFieldByName("SPPID")
            MyDDE.ChildRecordset.Fields("NoItem") = mCall.GetFieldByName("No barang")
            MyDDE.ChildRecordset.Fields("ItemName") = mCall.GetFieldByName("Nama Barang")
            MyDDE.ChildRecordset.Fields("UOM") = mCall.GetFieldByName("UOM")
            MyDDE.ChildRecordset.Fields("Keperluan") = "-"
            MyDDE.ChildRecordset.Fields("Note") = "-"
            MyDDE.ChildRecordset.Fields("QTY_SPP") = 1
            MyDDE.ChildRecordset.Fields("status") = 0
End Select
End If
End Sub

Private Sub MergeDoubleItem()
   
End Sub


Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If mEdit = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Exit Sub
End If
With DGPurchase
     Select Case .Col
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
                  MyDDE.GetFieldByName("SPP_Date") = DTPicker1.value
                  PrepareQuery
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
Dim PrintPreview As New Utility
txtBox(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            mEdit = True
            mEditPO = True
            Call DGPurchase_RowColChange(DGPurchase.Row, DGPurchase.Col)
       Case tmbAddNew:
            mEdit = True
            DTPicker1.value = Now
            MyDDE.GetFieldByName("Date_SPP") = DTPicker1.value
            MyDDE.GetFieldByName("Note") = ""
            MyDDE.GetFieldByName("SPPID") = IDGen.GetID("PP")
            MyDDE.GetFieldByName("status") = False
            DTPicker1.SetFocus
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               mEdit = False
               mEditPO = False
               OpenDetail txtBox(0)
               mVarDetailPOClose = False
            Else
               MessageBox "Detail Item  belum ada datanya.", "Peringatan", msgOkOnly
            End If

       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               mEdit = False
               mVarDetailPOClose = False
             End If
       Case tmbDetail:
             OpenPartner 3
             mEdit = True
             mVarDetailPOClose = False
       Case tmbPrint:
         Dim aReport As New Utility
         aReport.CallReportView "Select * From QueryListSPP where Status =0", "ListSPP.rpt", ReportPath, "Daftar Permintaan Pembelian"
         Set aReport = Nothing
       Case tmbQuit:
            'Unload Me
            'Set MyDDE.BindForm = Nothing
End Select

Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("SPPID")
mEdit = False
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen MyData.UploadQuery("Supplier"), CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen MyData.UploadQuery("BANK", MyDDE.GetFieldByName("PartnerID")), CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT [Remainder PO].NoItem, Inventory.ItemName, Inventory.[Serial Supplier], [Remainder PO].QTYOrder, Inventory.PPn, Inventory.PriceIn * (Inventory.Markup / 100)   + Inventory.PriceIn AS Harga, [Remainder PO].SCNo FROM [Remainder PO] INNER JOIN Inventory ON [Remainder PO].NoItem = Inventory.NoItem ORDER BY [Remainder PO].NoItem", CNN, lckLockReadOnly
       Case 3:
            RcPartner.DBOpen "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn,PriceIn AS Harga FROM         Inventory WHERE     (Manufacture = 0) ORDER BY NoItem", CNN, lckLockReadOnly
            'mFirstCaller = True
       Case 4:
            RcPartner.DBOpen "Select Code as Kode, Description as Keterangan,  [Bal_ Account Type], [Bal_ Account No_] from TermMethod ", CNN, lckLockReadOnly
       Case 5:
            RcPartner.DBOpen "Select No_ as Kode, Description as Keterangan, [Gen_ Prod_ Posting Group],  [Tax Group Code], [VAT Prod_ Posting Group], [Search Description], [Global Dimension 1 Code], [Global Dimension 2 Code] from item_charge ", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Supplier List"
            mCall.CaptionLink = "Supplier"
          Case 1:
            mCall.FromTagActive = "Bank List"
          Case 2:
            mCall.FromTagActive = "Remindier"
          Case 3:
            mCall.FromTagActive = "Inventory List"
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
RcDetail.DBOpen " SELECT * FROM QuerySPP where SPPID='" & ParameterString & "'", CNN, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
End Sub

Private Sub SimpanDetail()
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM SPP_line WHERE SPPID = '" & txtBox(0) & "'") = True Then
           Do
              If .EOF = True Then Exit Do
              SendDataToServer " INSERT INTO SPP_Line ( SPPID, NoItem, QTY_SPP, Keperluan, Note) " & _
                               " VALUES (N'" & txtBox(0) & "', N'" & .Fields("NoItem") & "', " & .Fields("QTY_SPP") & ", N'" & .Fields("Keperluan") & "','" & .Fields("Note") & "')"
              .MoveNext
           Loop
           End If
           .MoveLast
           DGPurchase.Refresh
     End If
End With
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub




Private Sub PrepareQuery()

On Error Resume Next
Dim strSQl As String
With MyDDE
   strSQl = " INSERT INTO  SPP_Header ( SPPID," & _
                                       "SPP_Date," & _
                                       "Note," & _
                                       "Ordered_by," & _
                                       "Status," & _
                                       "kode_dept) " & _
            " Values ('" & .GetFieldByName("SPPID") & _
                   "','" & Format(.GetFieldByName("SPP_Date"), "yyyy-MM-dd") & _
                   "','" & .GetFieldByName("Note") & _
                   "','" & MainMenu.StatusBar1.Panels(1).Text & _
                   "', 0,'" & .GetFieldByName("kode_dept") & "')"
    .PrepareAppend = strSQl
                     
                     
    strSQl = " UPDATE SPP_Header set SPP_Date = '" & Format(.GetFieldByName("SPP_Date"), "yyyy-MM-dd") & _
                               "', Note = '" & .GetFieldByName("Note") & _
                               "', EmpID = '" & MainMenu.StatusBar1.Panels(1).Text & _
                               "', status = " & CDbl(.GetFieldByName("Status")) & _
                               " , kode_dept = '" & .GetFieldByName("kode_dept") & _
             "' where SPPID ='" & .GetFieldByName("SPPID") & "'"
    Debug.Print strSQl
    .PrepareUpdate = strSQl
                     
    .PrepareDelete = " DELETE FROM  SPP_Header WHERE (SPPID = '" & .GetFieldByName("SPPID") & "')"
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
            If Temp <> "" Then
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


