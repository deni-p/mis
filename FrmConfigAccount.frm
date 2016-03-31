VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F2DD8007-5788-48C8-839C-E57EEDFCBFC6}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmConfigAccount 
   Caption         =   "Konfigurasi Jurnal"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConfigAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11880
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
      Height          =   6705
      Left            =   0
      ScaleHeight     =   6645
      ScaleWidth      =   11640
      TabIndex        =   4
      Top             =   -15
      Width           =   11700
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   5760
         Left            =   105
         ScaleHeight     =   5730
         ScaleWidth      =   11430
         TabIndex        =   5
         Top             =   780
         Width           =   11460
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   960
            Left            =   8445
            TabIndex        =   9
            Top             =   2385
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   1693
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
            BorderStyle     =   0
            Enabled         =   0   'False
            ColumnHeaders   =   0   'False
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "List Field"
               Caption         =   "List Field"
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
               DataField       =   "Sub Total"
               Caption         =   "Sub Total"
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
               ScrollBars      =   2
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   2324.977
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   2069.858
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   3120
            Left            =   3345
            TabIndex        =   2
            Top             =   1680
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   5503
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   15
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
               DataField       =   "No Akun"
               Caption         =   "No Akun"
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
               DataField       =   "Nama Akun"
               Caption         =   "Nama Akun"
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
               DataField       =   "Tipe"
               Caption         =   "Tipe Akun"
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
               DataField       =   "Posisi"
               Caption         =   "Posisi"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "DR"
                  FalseValue      =   "CR"
                  NullValue       =   "CR"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Value Data"
               Caption         =   "Link Data"
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
                  ColumnWidth     =   1409.953
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2745.071
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1964.976
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2550.047
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   1860
            Top             =   4290
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConfigAccount.frx":08CA
                  Key             =   "Kelompok Journal"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConfigAccount.frx":11A4
                  Key             =   "BKK"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConfigAccount.frx":1A7E
                  Key             =   "List"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConfigAccount.frx":2358
                  Key             =   "BKM"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConfigAccount.frx":2C32
                  Key             =   "Penjualan"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConfigAccount.frx":350C
                  Key             =   "Pembelian"
               EndProperty
            EndProperty
         End
         Begin SemeruDC.SemeruTree Menuku 
            Height          =   5490
            Left            =   105
            TabIndex        =   8
            Top             =   105
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   9684
            BackColorTree   =   7159830
            BackColorBackground=   -2147483633
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Nama Konfigurasi"
            Height          =   315
            Index           =   1
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Partner"
            Top             =   900
            Width           =   4710
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "Kode Konfigurasi"
            Height          =   315
            Index           =   0
            Left            =   5040
            MaxLength       =   16
            TabIndex        =   0
            Tag             =   "Partner"
            Top             =   555
            Width           =   1935
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   3345
            X2              =   5070
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3345
            X2              =   5070
            Y1              =   855
            Y2              =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Journal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   3360
            TabIndex        =   7
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Journal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   3360
            TabIndex        =   6
            Top             =   585
            Width           =   1050
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   7185
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   873
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmConfigAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcJournal As New DBQuick
Private RcPartner As New DBQuick
Private WithEvents LookLink As Recordset
Attribute LookLink.VB_VarHelpID = -1
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private mEdit As Boolean
Private mVarNodeKey As String

Private Sub DataGrid1_DblClick()
   DataGrid1.Visible = False
   DataGrid1.Enabled = False
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 27 Then
   DataGrid1.Visible = False
   DataGrid1.Enabled = False
End If
End Sub

Private Sub DataGrid1_LostFocus()
   DataGrid1.Visible = False
   DataGrid1.Enabled = False
End Sub

Private Sub DataGrid2_ButtonClick(ByVal ColIndex As Integer)
If mEdit = False Then Exit Sub
Select Case ColIndex
       Case 3:
            If DataGrid2.Columns(ColIndex).Value = 0 Then
               DataGrid2.Columns(ColIndex).Value = 1
            Else
               DataGrid2.Columns(ColIndex).Value = 0
            End If
       Case 4:
            If DataGrid1.Visible = False Then
               CallGroupData
            Else
               DataGrid1.Visible = False
               DataGrid1.Enabled = False
            End If
End Select
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mEdit = True Then
   If DataGrid2.Col = 3 Then DataGrid2.AllowUpdate = True
Else
   DataGrid2.AllowUpdate = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
Set mCall = New frmCaller
CreateList
OpenDB "XXX"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If MyDDE.CheckRecordPendinged = True Then
'   ScanKey vbKeyF5, 0, MyDDE
'   If MyDDE.IsSucces = True Then
'      Cancel = False
'      MyDDE.ClearRecordset
'   Else
'      Cancel = True
'   End If
'Else
'   MyDDE.ClearRecordset
'End If
End Sub

Private Sub Form_Resize()

HiasForm Picture1, Me
CenterForm Picture2, Me
Menuku.BackColorBackground = Picture2.BackColor
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmConfigAccount = Nothing
End Sub

Private Sub LookLink_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If LookLink.Recordcount <> 0 Then
   If mEdit = True Then
       MyDDE.ChildRecordset.Fields("Value Data") = LookLink.Fields(0)
       MyDDE.ChildRecordset.Fields("Nama Field") = LookLink.Fields(0)
       'MyDDE.ChildRecordset.Fields("Posisi") = mVarSetingJournalData.Posisijournal
   End If
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "MASTER ACCOUNT":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = mCall.GetFieldByName(2)
                 .Fields(3) = 0
            End With
End Select
End Sub

Private Sub Menuku_NodeClick(ByVal Node As MSComctlLib.INode)
mVarNodeKey = Node.Key
Menuku.Tag = Node.Text
OpenDB mVarNodeKey

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            MyDDE.GetFieldByName("Kode Konfigurasi") = mVarNodeKey
            MyDDE.GetFieldByName("Nama Konfigurasi") = Menuku.Tag
            txtBox(0).SetFocus
            txtBox(0).Enabled = False
            txtBox(1).Enabled = False
            mEdit = True
       Case tmbEdit:
            txtBox(0).Enabled = False
            txtBox(1).SetFocus
            txtBox(0).Enabled = False
            txtBox(1).Enabled = False
            mEdit = True
       Case tmbDetail:
            OpenPartner 0
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               mEdit = False
            End If
       Case tmbPrint:
            CallRPTReport "Konfigurasi journal.rpt"
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount <> 0 Then
               mEdit = True
            Else
               mEdit = False
            End If
       Case Else:
End Select
Menuku.MenuTreeView.Enabled = Not mEdit
DataGrid2.Columns(3).Button = mEdit
DataGrid2.Columns(4).Button = mEdit
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Opendetail MyDDE.GetFieldByName("Kode Konfigurasi")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            Select Case mVarNodeKey
                   Case "Kelompok Journal", "BKK", "BKM", "JUAL", "BELI", "":
                        MyDDE.CancelTrans = True
                        MessageBox "Bukan Kelompok Transaksi", "Peringatan", msgOkOnly
                   Case Else: MyDDE.CancelTrans = False
                        
            End Select
       Case tmbEdit:
            '
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False And MyDDE.ChildRecordset.Recordcount <> 0 And CheckEmptyGrid(MyDDE.ChildRecordset, "Value Data") = False Then
               MyDDE.IsChildMemberReady = True
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

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO  [Tabel Configurasi] ([Kode Konfigurasi], [Nama Konfigurasi], Keterangan) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(1)) & "')"

    .PrepareUpdate = " UPDATE [Tabel Configurasi] Set [Nama Konfigurasi] = N'" & ValidString(txtBox(1)) & "', Keterangan =N'" & ValidString(txtBox(1)) & "'  WHERE     ([Kode Konfigurasi] = N'" & ValidString(txtBox(0)) & "')"
    .PrepareDelete = " DELETE FROM [Tabel Configurasi] WHERE   ([Kode Konfigurasi] = N'" & ValidString(txtBox(0)) & "') "
End With
Err.Clear
End Sub

Private Sub Opendetail(ByVal NoKonfig As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen "SELECT  [Daftar Configurasi].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun], GlAccount.Type AS Tipe, [Daftar Configurasi].Posisi,[Daftar Configurasi].[Nama Form] as [Nama Field],[Daftar Configurasi].[Value Data] FROM         GlAccount INNER JOIN                       [Daftar Configurasi] ON GlAccount.NoAccount = [Daftar Configurasi].NoAccount WHERE     ([Daftar Configurasi].[Kode Konfigurasi] = N'" & NoKonfig & "') ORDER BY [Daftar Configurasi].NoAccount", Cnn, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid2.DataSource = MyDDE.ChildRecordset

RcDetail.CloseDB
Set RcDetail = Nothing
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:

Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun], Type AS [Tipe Akun] FROM         GlAccount WHERE     ([Group]  = N'Detail List Account') ORDER BY NoAccount", Cnn, lckLockReadOnly
            mCall.FromTagActive = "MASTER ACCOUNT"
            'mCall.txtCari = lblBank(0)
End Select
If RcPartner.Recordcount <> 0 Then
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
'    If FindOwnRecordset(MyDDE.ChildRecordset, "[No Akun] = '" & MyDDE.ChildRecordset.Fields("No Akun") & "'") = True Then
'       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("No Akun") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
'       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
'       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
'       DataGrid2.SetFocus
'    End If
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub
Private Sub SimpanDetail()
Dim I As Integer
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        SendDataToServer ("DELETE FROM [Daftar Configurasi] WHERE     ([Kode Konfigurasi] = N'" & ValidString(txtBox(0)) & "')")
        .MoveFirst
        I = 0
        Do
         I = I + 1
         If .EOF Then Exit Do
         If CBool(.Fields("Posisi")) = True Then
            SendDataToServer ("INSERT INTO [Daftar Configurasi]  ([Kode Konfigurasi], NoAccount, Posisi,[Value Data],[Nama Form],[No Index]) VALUES     (N'" & ValidString(txtBox(0)) & "', N'" & .Fields("No Akun") & "',1,N'" & .Fields("Value Data") & "',N'" & .Fields("Nama Field") & "'," & I & ")")
         Else
            SendDataToServer ("INSERT INTO [Daftar Configurasi]  ([Kode Konfigurasi], NoAccount, Posisi,[Value Data],[Nama Form],[No Index]) VALUES     (N'" & ValidString(txtBox(0)) & "', N'" & .Fields("No Akun") & "', 0,N'" & .Fields("Value Data") & "',N'" & .Fields("Nama Field") & "'," & I & ")")
         End If
        .MoveNext
        Loop
        .MoveLast
     End If
End With
End Sub

Private Sub CreateList()
With Menuku
     Set .MenuTreeView.ImageList = ImageList1
     .NodeAdd , , "Kelompok Journal", "Seting Journal", "Kelompok Journal", , , True, , True, True, , &HFCF1ED, &H6D4016
     .NodeAdd "Kelompok Journal", tvwChild, "BKK", "Konfigurasi Jurnal BKK", "BKK", , , True, , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "BKK", tvwChild, "BKKPTP", "Pembelian Tunai Persediaan", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "BKK", tvwChild, "BKKPTPKK", "Pengeluaran Tunai Piutang Ke Karyawan", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        '.NodeAdd "BKK", tvwChild, "BKKPHDU", "Pembayaran Hutang Dagang/Usaha", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        '.NodeAdd "BKK", tvwChild, "BKKPTAT", "Pembelian Tunai Aktiva Tetap", "List", , , , , True, True, , &HFCF1ED, &H6D4016
     .NodeAdd "Kelompok Journal", tvwChild, "BKM", "Konfigurasi Jurnal BKM", "BKM", , , True, , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "BKM", tvwChild, "BKMPTP", "Penjualan Tunai Persediaan", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        '.NodeAdd "BKM", tvwChild, "BKMPTPKK", "Pembayaran Piutang Karyawan", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        '.NodeAdd "BKM", tvwChild, "BKMPHDU", "Penerimaan Piutang Dagang/Usaha", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        '.NodeAdd "BKM", tvwChild, "BKMPTAT", "Penjualan Tunai Aktiva Tetap", "List", , , , , True, True, , &HFCF1ED, &H6D4016
     .NodeAdd "Kelompok Journal", tvwChild, "JUAL", "Konfigurasi jurnal Penjualan", "Penjualan", , , True, , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "JUAL", tvwChild, "BPJK", "Penjualan Persediaan (Kredit)", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "JUAL", tvwChild, "BRPJ", "Retur Penjualan Persediaan", "List", , , , , True, True, , &HFCF1ED, &H6D4016
     .NodeAdd "Kelompok Journal", tvwChild, "BELI", "Konfigurasi jurnal Pembelian", "Pembelian", , , True, , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "BELI", tvwChild, "BPBK", "Pembelian Persediaan (Kredit)", "List", , , , , True, True, , &HFCF1ED, &H6D4016
        .NodeAdd "BELI", tvwChild, "BRPB", "Retur Pembelian Persediaan", "List", , , , , True, True, , &HFCF1ED, &H6D4016
'     .NodeAdd "Kelompok Journal", tvwChild, "Depresiasi", "Depresiasi Aktiva", "Pembelian", , , True, , True, True, , &HFCF1ED, &H6D4016
'        .NodeAdd "Depresiasi", tvwChild, "DEPAT", "Akumulasi Penyusutan Aktiva", "List", , , , , True, True, , &HFCF1ED, &H6D4016
End With
End Sub

Private Sub OpenDB(ByVal ParamString As String)
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmConfigAccount
    .BindFormTAG = "Partner"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT * FROM [Tabel Configurasi] where [Kode Konfigurasi] = N'" & ParamString & "' ORDER BY [Kode Konfigurasi]"
End With

End Sub

Private Sub CallGroupData()
Dim Fld As Field
Dim I As Integer
If RcJournal.DBOpen(" Select * From [Journal " & mVarNodeKey & "]", Cnn, lckLockReadOnly) = True Then
   CloseDB LookLink
   Set LookLink = New Recordset
   With LookLink
       .Fields.Append "List Field", adBSTR
       .Fields.Append "Sub Total", adBSTR
       .Open
       Set DataGrid1.DataSource = LookLink
       I = 0
       For Each Fld In RcJournal.DBRecordset.Fields
           If InStr(UCase(RcJournal.DBRecordset.Fields(I).Name), "KODE") Then
              .AddNew "List Field", RcJournal.DBRecordset.Fields(I).Name
              .Fields("Sub Total").Value = RcJournal.DBRecordset.Fields(I + 1).Name
           End If
           I = I + 1
       Next
       If .Recordcount <> 0 Then
          DataGrid1.Enabled = True
          DataGrid1.Visible = True
          DataGrid1.Move DataGrid2.Columns(4).Left + DataGrid2.Columns(4).Width + 800, DataGrid2.Columns(4).Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight + DataGrid1.Height + 310, DataGrid2.Columns(4).Width
       End If
   End With
Else
   MessageBox "Data Tabel Journal belum ada.", "Peringatan", msgOkOnly
End If
End Sub
