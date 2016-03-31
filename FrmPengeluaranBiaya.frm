VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmBKK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengeluaran Kas"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPengeluaranBiaya.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10155
   Tag             =   "Cash Payment"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   0
      ScaleHeight     =   5370
      ScaleWidth      =   10155
      TabIndex        =   9
      Top             =   0
      Width           =   10155
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         DataField       =   "note"
         Height          =   330
         Left            =   1440
         TabIndex        =   16
         Tag             =   "ASM"
         Text            =   " - Keterangan -"
         Top             =   1515
         Width           =   8520
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4545
         Picture         =   "FrmPengeluaranBiaya.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   818
         Width           =   405
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal Bukti"
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Tag             =   "ASM"
         Top             =   465
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
         Format          =   139329539
         CurrentDate     =   38272
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2970
         Left            =   135
         TabIndex        =   6
         Top             =   1920
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   5239
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
      Begin VB.Line Line1 
         Index           =   6
         X1              =   150
         X2              =   1455
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   17
         Top             =   1590
         Width           =   840
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
         Left            =   1440
         TabIndex        =   4
         Tag             =   "ASM"
         Top             =   1155
         Width           =   2670
      End
      Begin VB.Label lblAlamatBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Nama Kas"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Tag             =   "ASM"
         Top             =   810
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7275
         TabIndex        =   5
         Top             =   1155
         Width           =   2670
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   6450
         TabIndex        =   15
         Top             =   4995
         Width           =   465
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
         Left            =   7740
         TabIndex        =   7
         Top             =   4920
         Width           =   2205
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   6465
         X2              =   7815
         Y1              =   5220
         Y2              =   5220
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
         Left            =   1440
         TabIndex        =   0
         Tag             =   "ASM"
         Top             =   120
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   195
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kas"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   13
         Top             =   870
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   517
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kas"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1215
         Width           =   270
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   1485
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   4
         Left            =   5970
         TabIndex        =   10
         Top             =   1215
         Width           =   435
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   135
         X2              =   1485
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   135
         X2              =   1815
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1485
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5970
         X2              =   7320
         Y1              =   1470
         Y2              =   1470
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   5370
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmBKK"
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
On Error GoTo 1
If ColIndex = 4 Then
   If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = "0"
   HitungTotal
Else
   If DGPurchase.Columns(ColIndex) = "" Then DGPurchase.Columns(ColIndex) = "-"
End If
Exit Sub
1:
MessageBox Err.Description, "frmbkk_dgpurchase_aftercoledit" & Err.Number, msgOkOnly, msgExclamation
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
On Error GoTo 2
If mVarAdd = True Then
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Select Case DGPurchase.col
          Case 2, 3, 4: DGPurchase.AllowUpdate = True
          Case Else: DGPurchase.AllowUpdate = False
   End Select
Else
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRow
End If
Exit Sub
2:
MessageBox Err.Description, "frmbkk_dgpurchase_rowcolchange" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub DTPicker1_Change()
On Error GoTo 1
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
Exit Sub
1:
MessageBox Err.Description, "frmbkk_dtpicker1_change" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
On Error GoTo 1
GridLayout
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmBKK
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT [Table Journal].JournalID AS [No Bukti], [Table Journal].NoAccount AS [Kode Kas], [Table Journal].DateTrans AS [Tanggal Bukti], " & _
            " [Table Journal].Periode, [Table Journal].PartnerID AS [Kode Partner], ' ' AS [Nama Partner], GLAccount.AccountName AS [Nama Kas], [Table Journal].note " & _
            " FROM [Table Journal] INNER JOIN GLAccount ON [Table Journal].NoAccount = GLAccount.NoAccount " & _
            " WHERE ([Table Journal].TypeTrans = N'BKK BIAYA') AND ([Table Journal].Periode = " & mVarPeriode & ") AND ([Table Journal].Status = 0) " & _
            " ORDER BY [Table Journal].JournalID"
End With
MyData.GetTglPeriode
Tmp = False
Exit Sub
1:
MessageBox Err.Description, "frmbkk_form_load" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcGroup.CloseDB
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
Set mBook = Nothing
End Sub

Private Sub Form_Resize()
'
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmBKK = Nothing
End Sub

Private Sub mCall_BeforeUnload()
On Error GoTo 1
'   Select Case mCall.FromTagActive
'          Case "MASTER KAS": If DGPurchase.Enabled = True Then DGPurchase.SetFocus
'          Case "MASTER PERKIRAAN": If DGPurchase.Enabled = True Then DGPurchase.SetFocus
'                mFirstCaller = False
'   End Select
Select Case UCase(mCall.FromTagActive)
       Case "MASTER KAS":
            If IsNull(MyDDE.GetFieldByName("Kode Kas")) = True Or MyDDE.GetFieldByName("Kode Kas") = "" Then MyDDE.CallButtonActive tmbCancel
            mVarAdd = False
            cmdLink(1).Enabled = mVarAdd
            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
       Case "MASTER PERKIRAAN":
            If IsNull(MyDDE.ChildRecordset.Fields(0)) = True Or MyDDE.ChildRecordset.Fields(0) = "" Then
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            End If
            If DGPurchase.Enabled = True Then DGPurchase.SetFocus
            mFirstCaller = False
End Select
mVarAdd = DTPicker1.Enabled
cmdLink(1).Enabled = mVarAdd
Exit Sub
1:
MessageBox Err.Description, "frmbkk_mcall_beforeunload" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error GoTo 1
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
                 lblTotalKas = FormatNumber(mCall.GetFieldByName(2), 0)
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
End Select
Exit Sub
1:
MessageBox Err.Description, "frmbkk_mcall_rowcolchange" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error GoTo 1
Select Case AdReasonActiveDb
       Case tmbEdit:
            mVarAdd = True
            DTPicker1.SetFocus
       Case tmbAddNew:
            lblAlamatBank = nama_kas
            NoVoucher(1) = kode_kas
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            With MyDDE
                 .GetFieldByName("Tanggal Bukti") = DTPicker1.Value
                 '.GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiBiayaKeluarJournal, 5, "", TglIndex)
                 .GetFieldByName("No Bukti") = IDGen.GetID("KK")
            End With
            mVarAdd = True
            DTPicker1.SetFocus
            lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(kode_kas), kode_kas, "XXXXX")), 0)
            cmdLink(1).Enabled = True
       Case tmbDetail:
            mVarAdd = True
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(2) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
            
       Case tmbDetail:
            If MyDDE.IsChildMemberReady = True Then If MyDDE.ChildRecordset.Recordcount > 1 Then mBook = MyDDE.ChildRecordset.AbsolutePosition
       Case tmbDelete:
       
            If MyDDE.IsChildMemberReady = True Then
               'SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & 'txtBox(0) & "') ")
               mVarAdd = False
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
                nama_kas = lblAlamatBank
                kode_kas = NoVoucher(1)
               SimpanDetail
               'lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX"), True), 0)
               mVarAdd = False
               Tmp = True
               cmdLink(1).Enabled = False
            End If
       Case tmbPrint:
            CallRPTReport "Bukti BKK.Rpt", "Select * from [Bukti BKK] Where [No Bukti]='" & lblFixAssets(0) & "'"
End Select
mVarAdd = DTPicker1.Enabled
'cmdLink(1).Enabled = mVarAdd
If Tmp = False Then
   lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
Else
   lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(kode_kas), kode_kas, "XXXXX")), 0)
End If

Exit Sub
1:
MessageBox Err.Description, "frmbkk_mydde_afterprepareactivedb" & Err.Number, msgOkOnly, msgExclamation
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
On Error GoTo 2
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGrid = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If CekKeterangan = False Then
                    If CCur(LblAmount) > CCur(lblTotalKas) Then
                       MyDDE.IsChildMemberReady = False
                       MyDDE.CancelTrans = True
                       MessageBox "Jumlah kas tidak cukup untuk melakukan transaksi.", "Peringatan", msgOkOnly, msgCrtical
                       DGPurchase.SetFocus
                   Else
                       MyDDE.IsChildMemberReady = True
                    End If
                  Else
                       MyDDE.IsChildMemberReady = False
                       MyDDE.CancelTrans = True
                  End If
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada/belum siap. Harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
                  DGPurchase.SetFocus
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
            'mVarAdd = False
       
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
            Else
               MyDDE.IsChildMemberReady = False
            End If
            'mVarAdd = False
       Case tmbCancel: 'mVarAdd = False
       Case tmbDetail:
            MyDDE.CancelTrans = mFirstCaller
            If mFirstCaller = True Then Exit Sub
            If NoVoucher(1) <> "" Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If MyDDE.ChildRecordset.Fields(4) = 0 Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly, msgCrtical
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

End Select
mVarAdd = DTPicker1.Enabled
Exit Sub
2:
MessageBox Err.Description, "frmbkk_mydde_prepareactivedb" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub PrepareQuery()
On Error GoTo 6
With MyDDE

    .PrepareAppend = " INSERT INTO [Table Journal]" & _
                     " ( [JournalID],   DateTrans,  NoAccount,TypeTrans,Periode,NoUrut,note)" & _
                     " VALUES  (N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & NoVoucher(1) & "','BKK BIAYA', " & mVarPeriode & ",N'" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "', '" & txtNote.Text & "')"

    .PrepareUpdate = " UPDATE [Table Journal]" & _
                     " SET DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & _
                     " NoAccount=N'" & NoVoucher(1) & "', note = " & FNumText(txtNote.Text) & _
                     " WHERE ([JournalID] = N'" & lblFixAssets(0) & "')"

    .PrepareDelete = " DELETE FROM [Table Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "') "
End With
Err.Clear
Exit Sub
6:
MessageBox Err.Description, "frmbkk_preparequery" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub SimpanDetail()
On Error GoTo 7
Dim I As Integer
Dim strNotes As String
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Detail Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "')") = True Then
           strNotes = "Pengeluaran Kas " & lblFixAssets(0)
           .MoveFirst
           I = 0
           Do
             I = I + 1
             If .EOF = True Then Exit Do
             SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Debet,No) " & _
                               " VALUES (N'" & lblFixAssets(0) & "', N'" & ValidString(.Fields("Kode Akun")) & "', '" & ValidString(.Fields("Doc Reff")) & "',N'" & ValidString(.Fields("Keterangan")) & "', " & CCur(.Fields("Jumlah Transaksi")) & "," & I & ")")
             strNotes = ValidString(strNotes & "," & .Fields("Keterangan"))
             .MoveNext
           Loop
           SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Credit,No) " & _
                            " VALUES (N'" & lblFixAssets(0) & "', N'" & NoVoucher(1) & "', 'xxx',N'" & ValidString(Right(strNotes, 244)) & "', " & CCur(LblAmount) & "," & I & ")")
                            
           .MoveLast
           SendDataToServer (" UPDATE    [Table Journal] SET  RefNotes = N'" & ValidString(Left(strNotes, 244)) & "' WHERE     (JournalID = N'" & lblFixAssets(0) & "')")
        End If
     End If
End With
Exit Sub
7:
MessageBox Err.Description, "frmbkk_simpan detail" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen " SELECT     [Detail Journal].NoAccount AS [Kode Akun], GLAccount.AccountName AS [Nama Akun], [Detail Journal].Keterangan, [Detail Journal].[Doc Reff],                        [Detail Journal].Debet AS [Jumlah Transaksi] FROM         [Detail Journal] INNER JOIN                      GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([Detail Journal].Debet <> 0) AND ([Detail Journal].JournalID = N'" & ParamString & "') ORDER BY [Detail Journal].NoAccount", CNN
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset

End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
On Error GoTo 5
Select Case Index
       Case 0:
            RcPartner.DBOpen " select * from [Daftar PartnerBiayaKeluar]", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     GLAccount.NoAccount AS [Kode Kas], GLAccount.AccountName AS [Nama Kas], ISNULL(ABS(SUM(ISNULL([Tabel Pembantu].CurrentDR" & PeriodeFilter & ", 0) + [ListMaster Kas].Debet) - SUM(ISNULL([Tabel Pembantu].CurrentCR" & PeriodeFilter & ", 0) + [ListMaster Kas].Credit)), 0) AS Saldo FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type LEFT OUTER JOIN                       [ListMaster Kas] ON GLAccount.NoAccount = [ListMaster Kas].NoAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GLAccount.[Group] = N'Detail List Account') AND ([ListMaster Kas].Periode = " & mVarPeriode & " OR                       [ListMaster Kas].Periode IS NULL) GROUP BY GLAccount.NoAccount, GLAccount.AccountName, AccType.ID HAVING      (AccType.ID = 31 OR AccType.ID = 50 OR AccType.ID = 51) ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly
       Case 2:
'            RcPartner.DBOpen " SELECT     GLAccount.NoAccount AS [Kode Akun], GLAccount.AccountName AS [Nama Akun] FROM         GLAccount INNER JOIN                      AccType ON GLAccount.Type = AccType.Tipe" & _
                             " WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = 11 OR " & _
                             " AccType.ID = 38 OR  AccType.ID = 48 OR AccType.ID = 10 OR AccType.ID = 30 OR " & _
                             " AccType.ID = 3 OR AccType.ID = 29 OR AccType.ID = 2 OR AccType.ID = 49 OR " & _
                             " AccType.ID = 26 OR AccType.ID = 25 OR AccType.ID = 28 OR AccType.ID = 32 OR AccType.ID = 17 OR " & _
                             " AccType.ID = 45 OR AccType.ID = 59 OR AccType.ID = 60 OR AccType.ID = 61 OR " & _
                             " AccType.ID = 15 OR AccType.ID = 13 OR AccType.ID = 16 OR AccType.ID = 14) " & _
                             " ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly
             RcPartner.DBOpen " SELECT GlAccount.NoAccount AS [Kode Akun], GlAccount.AccountName AS [Nama Akun] FROM GlAccount INNER JOIN AccType ON GlAccount.ID = AccType.ID WHERE     (AccType.payment_group = N'PAYMENT') AND (GlAccount.[Group] = N'Detail List Account')", CNN, lckLockReadOnly
'            mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0: mCall.FromTagActive = "Master Partner"
          Case 1:
               mCall.FromTagActive = "Master Kas"
               mCall.txtCari = NoVoucher(1)
          Case 2: mCall.FromTagActive = "Master Perkiraan"
   End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
Exit Function
5:
MessageBox Err.Description, "frmbkk_openpartner" & Err.Number, msgOkOnly, msgExclamation

End Function

Private Sub CancelDetailTrans()
On Error GoTo 1
If MyDDE.ChildRecordset.Recordcount <> 0 Then
  If Not MyDDE.ChildRecordset.EOF Then MyDDE.ChildRecordset.MoveNext
  If MyDDE.ChildRecordset.EOF And MyDDE.ChildRecordset.Recordcount > 0 Then MyDDE.ChildRecordset.MoveLast
End If
Exit Sub
1:
MessageBox Err.Description, "frmbkk_canceldetailtrans" & Err.Number, msgOkOnly, msgExclamation

End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "KK-" & Format(Day(Date), "0#") & Format(Month(Date), "0#") & Right(Format(Year(Date), "0#"), 2) & "-"
End Function

Private Function CekGrid() As Boolean
On Error GoTo 2
Dim RcGrd As New DBQuick
Set RcGrd.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
RcGrd.DBRecordset.Filter = "[Jumlah Transaksi] = 0"
With RcGrd.DBRecordset
     If .Recordcount <> 0 Then
        CekGrid = True
     Else
        CekGrid = False
     End If
End With
RcGrd.CloseDB
Exit Function
2:
MessageBox Err.Description, "frmbkk_cekgrid" & Err.Number, msgOkOnly, msgExclamation
End Function

'edit

Private Function CekKeterangan() As Boolean
On Error GoTo 3
Dim RcGrd As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Avdata = RcGrd.Getrows(MyDDE.ChildRecordset, "Keterangan")
For I = 0 To UBound(Avdata, 2)
    'If IsNull(Avdata(0, I)) = True Or Avdata(0, I) = "" Or Avdata(0, I) = "-" Then
       If IsNull(Avdata(0, I)) = True Then
       CekKeterangan = True
       MessageBox "Keterangan Jurnal Kosong, harap diisi dulu.", "Peringatan", msgOkOnly, msgCrtical
       Exit For
    End If
Next I
Set RcGrd = Nothing
Exit Function
3:
MessageBox Err.Description, "frmbkk_cekketerangan" & Err.Number, msgOkOnly, msgExclamation

End Function

Private Sub HitungTotal()
On Error GoTo 4
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
4:
MessageBox Err.Description, "frmbkk_hitungtotal" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1379.906
'DGPurchase.Columns(1).width = 1814.74
DGPurchase.Columns(1).width = 5000
'DGPurchase.Columns(2).width = 3165.166
DGPurchase.Columns(3).width = 1514.835
DGPurchase.Columns(4).width = 1365.165
End Sub

