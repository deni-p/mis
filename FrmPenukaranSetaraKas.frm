VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPenukaranSetaraKas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penukaran Setara Kas"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPenukaranSetaraKas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10140
   Tag             =   "Penukaran Setara Kas Ke Kas"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
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
      Height          =   5550
      Left            =   0
      ScaleHeight     =   5550
      ScaleWidth      =   10140
      TabIndex        =   1
      Top             =   0
      Width           =   10140
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Tanggal Bukti"
         Height          =   315
         Left            =   1425
         TabIndex        =   3
         Tag             =   "ASM"
         Top             =   435
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
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4095
         Picture         =   "FrmPenukaranSetaraKas.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   803
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   3105
         Left            =   135
         TabIndex        =   2
         Top             =   1935
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
         EndProperty
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
         TabIndex        =   15
         Top             =   5100
         Width           =   2205
      End
      Begin VB.Label LBLTukar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1425
         TabIndex        =   9
         Top             =   1515
         Width           =   2670
      End
      Begin VB.Label LBLTukar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Nama Kas"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1425
         TabIndex        =   8
         Tag             =   "ASM"
         Top             =   1155
         Width           =   3000
      End
      Begin VB.Label LBLTukar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Kode Kas"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   1425
         TabIndex        =   7
         Tag             =   "ASM"
         Top             =   795
         Width           =   2670
      End
      Begin VB.Label LBLTukar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Bukti"
         DataField       =   "No Bukti"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1425
         TabIndex        =   6
         Tag             =   "ASM"
         Top             =   75
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kas"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   16
         Top             =   1230
         Width           =   705
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   180
         X2              =   1485
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6510
         X2              =   7815
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   180
         X2              =   1485
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   1485
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   6510
         TabIndex        =   14
         Top             =   5145
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   150
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   495
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Kas"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   11
         Top             =   870
         Width           =   660
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   180
         X2              =   1485
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   10
         Top             =   1590
         Width           =   390
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   180
         X2              =   1485
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label LBLTukar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "ID"
         Enabled         =   0   'False
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
         Height          =   330
         Index           =   4
         Left            =   4575
         TabIndex        =   5
         Tag             =   "ASM"
         Top             =   825
         Visible         =   0   'False
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmPenukaranSetaraKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcGroup As New DBQuick
Private MyData As New clsTransaksi
Private mVarAdd, mFirstCaller As Boolean
Private mBook As Variant
Private RcPartner As New DBQuick
Dim IDGen As New IDGenerator

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
          Case 2, 3, 4: DGPurchase.AllowUpdate = True
          Case Else: DGPurchase.AllowUpdate = False
   End Select
Else
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

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
    Set .BindForm = FrmPenukaranSetaraKas
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT [Table Journal].JournalID AS [No Bukti], [Table Journal].DateTrans AS [Tanggal Bukti], " & _
            " [Table Journal].Periode, [Table Journal].NoAccount AS [Kode Kas], GLAccount.AccountName AS [Nama Kas],GLAccount.ID " & _
            " FROM [Table Journal] INNER JOIN GLAccount ON [Table Journal].NoAccount = GLAccount.NoAccount " & _
            " WHERE ([Table Journal].TypeTrans = N'CHANGE') AND ([Table Journal].Periode = " & mVarPeriode & ")  " & _
            " AND ([Table Journal].Status = 0) ORDER BY [Table Journal].JournalID"
            
'    Debug.Print .PrepareQuery
End With
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

Private Sub Form_Unload(Cancel As Integer)
Set FrmPenukaranSetaraKas = Nothing
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
Select Case mCall.FromTagActive
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
                 .Fields("ID") = mCall.GetFieldByName(2)
            End With
            LBLTukar(2) = Format(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), AngkaForm)
       Case "MASTER PERKIRAAN":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(0)
                 .Fields(1) = mCall.GetFieldByName(1)
                 .Fields(2) = "-"
                 .Fields(3) = "-"
                 .Fields(4) = 0
            End With
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            mVarAdd = True
            DTPicker1.SetFocus
       Case tmbAddNew:
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            With MyDDE
                 .GetFieldByName("Tanggal Bukti") = DTPicker1.Value
'                 .GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiChange, 5, "", TglIndex)
                 .GetFieldByName("No Bukti") = IDGen.GetID("TK")
                 
            End With
            mVarAdd = True
            DTPicker1.SetFocus
       Case tmbDetail:
            mVarAdd = True
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(2) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               'SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & 'txtBox(0) & "') ")
               mVarAdd = False
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               TotalKas IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")
               mVarAdd = False
            End If
       Case tmbPrint:
            CallRPTReport "PenukaranKas.Rpt", "Select * from [PenukaranKas] where [No Bukti]=N'" & LBLTukar(0) & "'"
End Select
mVarAdd = DTPicker1.Enabled
cmdLink(1).Enabled = mVarAdd
'LBLTukar(2) = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
LBLTukar(2) = Format(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), AngkaForm)
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "XXXXX")
'LBLTukar(2) = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
LBLTukar(2) = Format(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), AngkaForm)
HitungTotal
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGrid = False And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If CDbl(LblAmount) > CDbl(LBLTukar(2)) Then
                     MessageBox "Total transaksi lebih besar dari total " & LBLTukar(1), "Peringatan", msgOkOnly, msgCrtical
                     MyDDE.CancelTrans = True
                  Else
                     MyDDE.IsChildMemberReady = True
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
            If LBLTukar(3) <> "" Then
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

End Select
mVarAdd = DTPicker1.Enabled
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [Table Journal]" & _
                     " ( [JournalID], DateTrans, NoAccount,TypeTrans,periode,NoUrut)" & _
                     " VALUES  (N'" & LBLTukar(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & LBLTukar(3) & "','CHANGE'," & mVarPeriode & ",'" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "')"

    .PrepareUpdate = " UPDATE [Table Journal]" & _
                     " SET DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & _
                     " NoAccount=N'" & LBLTukar(3) & "'" & _
                     " WHERE ([JournalID] = N'" & LBLTukar(0) & "')"

    .PrepareDelete = " DELETE FROM [Table Journal] WHERE     ([JournalID] = N'" & LBLTukar(0) & "') "
End With
Err.Clear
End Sub

Private Sub SimpanDetail()
Dim I As Integer
Dim strParticKredit As String
Dim strParticDebet As String
StrPartic = ""
With MyDDE.ChildRecordset
    If .Recordcount <> 0 Then
       If SendDataToServer("DELETE FROM [Detail Journal] WHERE ([JournalID] = N'" & LBLTukar(0).Caption & "')") = True Then
           strParticDebet = "Transferan dari " & LBLTukar(1).Caption
           strParticKredit = "Transfer " & LBLTukar(1).Caption & " ke "
          
           .MoveFirst
          I = 1
          Do
            I = I + 1
            If .EOF = True Then Exit Do
            SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Debet,[No]) " & _
                              " VALUES (N'" & LBLTukar(0) & "', N'" & .Fields("Kode Akun") & "', '" & _
                              ValidString(.Fields("Doc Reff")) & "',N'" & ValidString(Left(strParticDebet, 244)) & "', " & _
                              CCur(.Fields("Jumlah Transaksi")) & "," & I & ")")
            strParticKredit = strParticKredit & "," & .Fields("Nama Akun").Value
            .MoveNext
          Loop
          SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Credit ,[NO]) " & _
                            " VALUES (N'" & LBLTukar(0) & "', N'" & LBLTukar(3) & "', 'xxx',N'" & ValidString(Left(strParticKredit, 244)) & "', " & CCur(LblAmount) & ",1)")
                           
          .MoveLast
          SendDataToServer ("Update [Table Journal] Set RefNotes ='" & ValidString(Left(StrPartic, 249)) & "' where JournalID='" & LBLTukar(0) & "'")
       End If
    End If
End With
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen " SELECT [Detail Journal].NoAccount AS [Kode Akun], GLAccount.AccountName AS [Nama Akun], " & _
        " [Detail Journal].Keterangan, [Detail Journal].[Doc Reff], [Detail Journal].Debet AS [Jumlah Transaksi] " & _
        " FROM [Detail Journal] INNER JOIN GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount " & _
        " WHERE ([Detail Journal].Debet <> 0) AND ([Detail Journal].JournalID = N'" & ParamString & "') " & _
        " ORDER BY [Detail Journal].NoAccount", CNN
'Debug.Print RcDetail.DBRecordset.Source
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Dim TypeRange As String

'11  Bank
'31  Kas
'51  Kas Kecil

Select Case Index
    Case 0:
         RcPartner.DBOpen " select * from [Daftar PartnerBiayaKeluar]", CNN, lckLockReadOnly
    Case 1:  'HEADER TRANSFER DEBET
         
         RcPartner.DBOpen "SELECT  GLAccount.NoAccount AS [Kode Kas], GLAccount.AccountName AS [Nama Kas], AccType.ID " & _
                 " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe " & _
                 " WHERE (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID IN (51,31,11))", CNN, lckLockReadOnly
    Case 2:  'DETAIL CREDIT
        
'        Select Case Val(LBLTukar(4).Caption)
'            Case 11:
'                TypeRange = "(31,51,11)"
'            Case 31:
'                TypeRange = "(11,51)"
'            Case 51:
'                TypeRange = "(31,11)"
'        End Select
        TypeRange = "(31,51,11)"
        RcPartner.DBOpen "SELECT GLAccount.NoAccount AS [Kode Akun], GLAccount.AccountName AS [Nama Akun] " & _
                " FROM GLAccount INNER JOIN AccType ON GLAccount.Type = AccType.Tipe " & _
                " WHERE (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID IN " & TypeRange & ")  AND (dbo.GLAccount.NoAccount <> N'" & LBLTukar(3).Caption & "')" & _
                " ORDER BY GLAccount.NoAccount", CNN, lckLockReadOnly
'        Debug.Print RcPartner.DBRecordset.Source
'        mFirstCaller = True

End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "Master Partner"
           Case 1:
                mCall.FromTagActive = "Master Kas"
                mCall.txtCari = LBLTukar(3)
           Case 2: mCall.FromTagActive = "Master Perkiraan"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   OpenPartner = True
End If
End Function

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
TglIndex = "TK-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Function CekGrid() As Boolean
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
End Function

Private Sub HitungTotal()
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
RcTotal.CloseDB
End Sub

'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
'End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1379.906
DGPurchase.Columns(1).width = 1814.74
DGPurchase.Columns(2).width = 3165.166
DGPurchase.Columns(3).width = 1514.835
DGPurchase.Columns(4).width = 1365.165
End Sub




