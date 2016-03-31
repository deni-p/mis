VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmBKM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penerimaan Kas"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBKM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10170
   Tag             =   "Cash Receipt"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   10275
      TabIndex        =   10
      Top             =   0
      Width           =   10275
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         DataField       =   "note"
         Height          =   330
         Left            =   1410
         TabIndex        =   6
         Tag             =   "ASM"
         Text            =   " - Keterangan -"
         Top             =   1522
         Width           =   8535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal Bukti"
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Tag             =   "ASM"
         Top             =   440
         Width           =   2670
         _ExtentX        =   4710
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71630851
         CurrentDate     =   38272
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4515
         Picture         =   "FrmBKM.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   818
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   3105
         Left            =   135
         TabIndex        =   7
         Top             =   1920
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   5477
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Kode Akun"
            Caption         =   "Kode Akun"
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
            DataField       =   "Keterangan"
            Caption         =   "Keterangan"
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
            DataField       =   "Doc Reff"
            Caption         =   "Doc Reff"
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
         BeginProperty Column04 
            DataField       =   "Jumlah Transaksi"
            Caption         =   "Jumlah Transaksi"
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
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   17
         Top             =   1590
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   135
         X2              =   2000
         Y1              =   1837
         Y2              =   1837
      End
      Begin VB.Label lblFixAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "No Bukti"
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
         Height          =   330
         Index           =   0
         Left            =   1410
         TabIndex        =   0
         Tag             =   "ASM"
         Top             =   75
         Width           =   2670
      End
      Begin VB.Label NoVoucher 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Kode Kas"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1410
         TabIndex        =   4
         Tag             =   "ASM"
         Top             =   1155
         Width           =   3105
      End
      Begin VB.Label lblAlamatBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Nama Kas"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1410
         TabIndex        =   2
         Tag             =   "ASM"
         Top             =   790
         Width           =   3105
      End
      Begin VB.Label lblTotalKas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   330
         Left            =   6840
         TabIndex        =   5
         Top             =   1155
         Width           =   3105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   16
         Top             =   1230
         Width           =   765
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
         Top             =   500
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   143
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   6615
         TabIndex        =   13
         Top             =   5160
         Width           =   555
      End
      Begin VB.Label LblAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   315
         Left            =   7545
         TabIndex        =   8
         Top             =   5100
         Width           =   2400
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6795
         X2              =   8100
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Akun Kas"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   12
         Top             =   855
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   6165
         TabIndex        =   11
         Top             =   1223
         Width           =   390
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   1970
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1985
         Y1              =   740
         Y2              =   740
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   135
         X2              =   2015
         Y1              =   1105
         Y2              =   1105
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   6120
         X2              =   7425
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   135
         X2              =   2000
         Y1              =   390
         Y2              =   390
      End
   End
End
Attribute VB_Name = "FrmBKM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcGroup As New DBQuick
Private MyData As New clsTransaksi
Private mVarAdd, mFirstCaller As Boolean
Private mBook As Variant
Private RcPartner As New DBQuick
Dim IDGen As New IDGenerator
Dim kode_kas As String
Dim nama_kas As String
Dim Tmp As Boolean



Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 4 Then
   If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = "0"
   HitungTotal
Else
   If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = "-"
End If
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If mVarAdd = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mVarAdd = True Then
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Select Case DGPurchase.col
          Case 2, 3, 4:
            DGPurchase.AllowUpdate = True
            DGPurchase.SetFocus
          Case Else: DGPurchase.AllowUpdate = False
   End Select
Else
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub DTPicker1_Change()
Dim I As Long

    If MyData.GetPeriodeStatus(DTPicker1.Value) Then
        MsgBox "Periode Transaksi " & Format(DTPicker1.Value, "mmmm") & " telah closing", vbExclamation, "Period Closing Control"
        DTPicker1.Value = Now 'Format(Now, "dd-mmm-yyyy HH:MM")
        Exit Sub
    Else
        Select Case DTPicker1.Value
            Case Is > EndTgl    'OUT OF CURRENT PERIOD DATE
                MsgBox "Tanggal Transaksi tidak dapat melebihi Periode yang telah ditentukan " & Chr(13) & _
                " ( " & Format(StartTgl, "dd mmm yyyy") & " - " & Format(EndTgl, "dd mmm yyyy") & " ) ", vbExclamation, "Period based Transaction"
                DTPicker1.Value = Now
            Case Is < StartTgl  'SUSULAN
                MsgBox "Tanggal Transaksi tidak dapat kurang dari Periode yang telah ditentukan " & Chr(13) & _
                " ( " & Format(StartTgl, "dd mmm yyyy") & " - " & Format(EndTgl, "dd mmm yyyy") & " ) ", vbExclamation, "Period based Transaction"
                DTPicker1.Value = Now
            Case Else
'                ChkSusulan.Value = 0
'                TXTEntry(4).Text = GetNoFSR
        End Select
    End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

'Private Sub Form_Activate()
'If Me.WindowState = 0 Then If Me.WindowState = 0 Then Me.WindowState = 2
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmBKM
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     [Table Journal].JournalID AS [No Bukti], [Table Journal].DateTrans AS [Tanggal Bukti], [Table Journal].Periode, [Table Journal].NoAccount AS [Kode Kas], [Table Journal].note , GLAccount.AccountName AS [Nama Kas] FROM         [Table Journal] INNER JOIN GLAccount ON [Table Journal].NoAccount = GLAccount.NoAccount WHERE     ([Table Journal].TypeTrans = N'BKM') AND ([Table Journal].Periode = " & mVarPeriode & ") AND ([Table Journal].Status = 0) ORDER BY [Table Journal].JournalID"
End With
MyData.GetTglPeriode
Tmp = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGroup.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
Set mBook = Nothing
End Sub

Private Sub Form_Resize()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub



Private Sub mCall_BeforeUnload()
If UCase(mCall.FromTagActive) = "MASTER KAS" Then
   If IsNull(MyDDE.GetFieldByName("Kode Kas")) = True Or MyDDE.GetFieldByName("Kode Kas") = "" Then MyDDE.CallButtonActive tmbCancel
   mVarAdd = False
   If DGPurchase.Enabled = True Then DGPurchase.SetFocus
ElseIf UCase(mCall.FromTagActive) = "MASTER PERKIRAAN" Then
    If IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields(0) = "" Then
       MyDDE.ChildRecordset.CancelBatch adAffectCurrent
       If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
    End If
    If DGPurchase.Enabled = True Then DGPurchase.SetFocus
    mFirstCaller = False
End If
mVarAdd = DTPicker1.Enabled
cmdLink(1).Enabled = mVarAdd
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case UCase(TagForm)
       Case "MASTER PARTNER":
            With MyDDE.ActiveRecordset
                 .Fields("Kode Partner") = mCall.GetFieldByName(0)
                 .Fields("Nama Partner") = mCall.GetFieldByName(1)
            End With
       Case "MASTER KAS":
            With MyDDE.ActiveRecordset
                 .Fields("KOde Kas") = mCall.GetFieldByName(0)
                 .Fields("Nama Kas") = mCall.GetFieldByName(1)
                 .Fields("note") = mCall.GetFieldByName("note")
            End With
            lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
       Case "MASTER PERKIRAAN":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = "-"
                 .Fields(3) = "-"
                 .Fields(4) = 0
            End With
'            DGPurchase.SetFocus
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            mVarAdd = True
            DTPicker1.SetFocus
       Case tmbAddNew:
            
            NoVoucher(1).Caption = kode_kas
            lblAlamatBank = nama_kas
            lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            With MyDDE
                 .GetFieldByName("Tanggal Bukti") = DTPicker1.Value
                 '.GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiBKM, 5, "", TglIndex)
                 .GetFieldByName("No Bukti") = IDGen.GetID("KM")
            End With
            mVarAdd = True
            DTPicker1.SetFocus
            lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(kode_kas), kode_kas, "XXXXX")), 0)
            cmdLink(1).Enabled = True
            txtNote.Text = "-"
       Case tmbDetail:
'            mVarAdd = DTPicker1.Enabled
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(2) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
'       Case tmbDetail:
'            If MyDDE.IsChildMemberReady = True Then If MyDDE.ChildRecordset.Recordcount > 1 Then mBook = MyDDE.ChildRecordset.AbsolutePosition
       Case tmbDelete:
       
            If MyDDE.IsChildMemberReady = True Then
               'SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & 'txtBox(0) & "') ")
'               mVarAdd = False
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               kode_kas = NoVoucher(1)
               nama_kas = lblAlamatBank
               SimpanDetail
               TotalKas IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")
               mVarAdd = False
               Tmp = True
               cmdLink(1).Enabled = False
            End If
            
       Case tmbPrint:
            HitungTotal
            CallRPTReport "Bukti BKM.Rpt", "Select * from [Bukti BKM] where [No Bukti]=N'" & lblFixAssets(0) & "'", True, CCur(LblAmount)
End Select
mVarAdd = DTPicker1.Enabled

If Tmp = False Then
   lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
   Else
   lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(kode_kas), kode_kas, "XXXXX")), 0)
End If
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "XXXXX")
lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
HitungTotal
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGrid = True And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada/belum siap. Harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
            'mVarAdd = cmdLink(1).Enabled
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
            'mVarAdd = cmdLink(1).Enabled
       Case tmbCancel: 'mVarAdd = cmdLink(1).Enabled
       Case tmbDetail:
'            DGPurchase.SetFocus
            MyDDE.CancelTrans = mFirstCaller
            If MyDDE.CancelTrans = True Then Exit Sub
            If NoVoucher(1) <> "" Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If MyDDE.ChildRecordset.Fields(4) = 0 Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly, msgCrtical
                     
                     DGPurchase.SetFocus
                  Else
                     MyDDE.IsChildMemberReady = True
                     MyDDE.CancelTrans = False
                  
                  End If
               Else
                  MyDDE.IsChildMemberReady = True
                  MyDDE.CancelTrans = False
               End If
            Else
                MessageBox "Data bank atau kas belum dipilih.", "Peringatan", msgOkOnly, msgCrtical
                MyDDE.IsChildMemberReady = False
                MyDDE.CancelTrans = True
            End If
'            DGPurchase.SetFocus
End Select
mVarAdd = DTPicker1.Enabled
End Sub

Private Sub PrepareQuery()
On Error GoTo xErr
With MyDDE
    .PrepareAppend = " INSERT INTO [Table Journal]" & _
                     " ( [JournalID], DateTrans, NoAccount,TypeTrans,periode,NoUrut,note)" & _
                     " VALUES  (N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & NoVoucher(1) & "','BKM'," & mVarPeriode & ",N'" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "', '" & txtNote.Text & "')"

    .PrepareUpdate = " UPDATE [Table Journal]" & _
                     " SET DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & _
                     " NoAccount=N'" & NoVoucher(1) & "', note= '" & txtNote.Text & "'" & _
                     " WHERE ([JournalID] = N'" & lblFixAssets(0) & "')"

    .PrepareDelete = " DELETE FROM [Table Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "') "
End With

Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub SimpanDetail()
Dim I As Integer
Dim strNotes As String
On Error GoTo xErr
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Detail Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "')") = True Then
           .MoveFirst
           I = 1
           strNotes = "Penerimaan Kas " & lblFixAssets(0)
           Do
             I = I + 1
             If .EOF = True Then Exit Do
             SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Credit,[No]) " & _
                               " VALUES (N'" & lblFixAssets(0) & "', N'" & .Fields("Kode Akun") & "', '" & ValidString(.Fields("Doc Reff")) & "',N'" & ValidString(.Fields("Keterangan")) & "', " & CCur(.Fields("Jumlah Transaksi")) & "," & I & ")")
             strNotes = strNotes & "-" & .Fields("Keterangan")
             .MoveNext
           Loop
           SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Debet,[NO]) " & _
                             " VALUES (N'" & lblFixAssets(0) & "', N'" & NoVoucher(1).Caption & "', 'xxx',N'" & ValidString(Left(strNotes, 250)) & "', " & CCur(LblAmount) & ",1)")
'           Debug.Print NoVoucher(1).Caption
           .MoveLast
           SendDataToServer (" UPDATE    [Table Journal] SET  RefNotes = N'" & ValidString(Left(strNotes, 244)) & "' WHERE     (JournalID = N'" & lblFixAssets(0) & "')")
        End If
     End If
End With
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen " SELECT     [Detail Journal].NoAccount AS [Kode Akun], GLAccount.AccountName AS [Nama Akun], [Detail Journal].Keterangan, [Detail Journal].[Doc Reff],                        [Detail Journal].Credit AS [Jumlah Transaksi] FROM         [Detail Journal] INNER JOIN                      GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([Detail Journal].Credit <> 0) AND ([Detail Journal].JournalID = N'" & ParamString & "') ORDER BY [Detail Journal].NoAccount", CNN
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
            RcPartner.DBOpen " select * from [Daftar PartnerBiayaKeluar]", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     GLAccount.NoAccount AS [Kode Kas], GLAccount.AccountName AS [Nama Kas] FROM         GLAccount INNER JOIN                       AccType ON GLAccount.Type = AccType.Tipe WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = 31 OR AccType.ID = 50)", CNN, lckLockReadOnly
       Case 2:
'            RcPartner.DBOpen " SELECT GLAccount.NoAccount AS [Kode Akun], " & _
'            " GLAccount.AccountName AS [Nama Akun] FROM GLAccount INNER JOIN   " & _
'            " AccType ON GLAccount.Type = AccType.Tipe WHERE (GLAccount.[Group] = N'Detail List Account') " & _
'            " AND (AccType.ID = 11 OR  AccType.ID = 38 OR  AccType.ID = 48 OR " & _
'            " AccType.ID = 10 OR  AccType.ID = 30 OR   AccType.ID = 3 OR   " & _
'            " AccType.ID = 29 OR   AccType.ID = 2 OR  AccType.ID = 27 OR   " & _
'            " AccType.ID = 24 OR   AccType.ID = 49 OR  AccType.ID = 26 OR    " & _
'            " AccType.ID = 25 OR    AccType.ID = 32 OR   AccType.ID = 39 OR   AccType.ID = 34) " & _
'            " ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly
'
'            mFirstCaller = True
            RcPartner.DBOpen " SELECT GlAccount.NoAccount AS [Kode Akun], GlAccount.AccountName AS [Nama Akun] FROM GlAccount INNER JOIN AccType ON GlAccount.ID = AccType.ID WHERE     (AccType.receipt_group = N'RECEIPT') AND (GlAccount.[Group] = N'Detail List Account')", CNN, lckLockReadOnly
'            mFirstCaller = True

End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "Master Partner"
           Case 1:
                mCall.FromTagActive = "Master Kas"
                mCall.txtCari = NoVoucher(1)
           Case 2:  mCall.FromTagActive = "Master Perkiraan"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
'    DGPurchase.SetFocus
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   
   OpenPartner = True
End If
End Function

Private Sub CancelDetailTrans()
On Error GoTo 1
If MyDDE.ChildRecordset.Recordcount <> 0 Then
  If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveNext
  If MyDDE.ChildRecordset.EOF And MyDDE.ChildRecordset.Recordcount > 0 Then MyDDE.ChildRecordset.MoveLast
End If
Exit Sub
1:
MessageBox Err.Description, "frmbkm_canceldetailtrans" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "KM-" & Format(Day(Date), "0#") & Format(Month(Date), "0#") & Right(Format(Year(Date), "0#"), 2) & "-"
End Function

Private Function CekGrid() As Boolean
On Error GoTo 2
Dim RcGrd As New DBQuick
Set RcGrd.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
RcGrd.DBRecordset.Filter = "[Jumlah Transaksi] = 0"
With RcGrd.DBRecordset
     If .Recordcount <> 0 Then
        CekGrid = False
     Else
        CekGrid = True
     End If
End With
Exit Function
2:
MessageBox Err.Description, "frmbkm_cekgrid" & Err.Number, msgOkOnly, msgExclamation
End Function

Private Sub HitungTotal()
On Error GoTo 3
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mTotal As Currency
Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotal = 0
With RcTotal.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            mTotal = mTotal + IIf(Not IsNull(Avdata(4, I)), Avdata(4, I), 0)
        Next I
     Else
        mTotal = 0
     End If
     LblAmount = FormatNumber(Abs(mTotal), 0)
End With
Set Avdata = Nothing
Exit Sub
3:
MessageBox Err.Description, "frmbkm_hitungtotal" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1379.906
'DGPurchase.Columns(1).width = 1814.74
DGPurchase.Columns(1).width = 5000
'DGPurchase.Columns(2).width = 3165.166
DGPurchase.Columns(3).width = 1514.835
DGPurchase.Columns(4).width = 1365.165
End Sub


