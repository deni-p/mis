VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F2DD8007-5788-48C8-839C-E57EEDFCBFC6}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmBKK 
   AutoRedraw      =   -1  'True
   Caption         =   "Pengeluaran Tunai Lainnya"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBKK.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   10605
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
      ScaleWidth      =   10395
      TabIndex        =   7
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5370
         Left            =   105
         ScaleHeight     =   5310
         ScaleWidth      =   10065
         TabIndex        =   8
         Top             =   630
         Width           =   10125
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   5430
            Picture         =   "frmBKK.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   795
            Width           =   405
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Height          =   3105
            Left            =   135
            TabIndex        =   5
            Top             =   1785
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
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3105.071
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1530.142
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Tanggal Bukti"
            Height          =   315
            Left            =   1815
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   450
            Width           =   4020
            _ExtentX        =   7091
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
            Format          =   22806531
            CurrentDate     =   38272
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   660
            X2              =   1965
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Kas"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   705
            TabIndex        =   14
            Top             =   1185
            Width           =   780
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   6810
            X2              =   8115
            Y1              =   5220
            Y2              =   5220
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   675
            X2              =   1980
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   675
            X2              =   1980
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Kode Kas"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   1830
            TabIndex        =   2
            Tag             =   "ASM"
            Top             =   795
            Width           =   3570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Kas"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   705
            TabIndex        =   13
            Top             =   840
            Width           =   750
         End
         Begin VB.Label lblAlamatBank 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Kas"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1830
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   1140
            Width           =   4005
         End
         Begin VB.Label lblTotalKas 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   210
            Left            =   1830
            TabIndex        =   4
            Top             =   1470
            Width           =   120
         End
         Begin VB.Label lblFixAssets 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   1830
            TabIndex        =   0
            Tag             =   "ASM"
            Top             =   165
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   705
            TabIndex        =   12
            Top             =   510
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   705
            TabIndex        =   11
            Top             =   165
            Width           =   690
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
            Left            =   6780
            TabIndex        =   10
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
            Height          =   240
            Left            =   7815
            TabIndex        =   9
            Top             =   4995
            Width           =   2100
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   6
      Top             =   6330
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   873
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmBKK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcGroup As New DBQuick
Private MyData As New clsTransaksi
Private mVarAdd As Boolean
Private mBook As Variant
Private RcPartner As New DBQuick

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 4 Then HitungTotal
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mVarAdd = True Then
   Select Case DGPurchase.Col
          Case 2, 3, 4: DGPurchase.AllowUpdate = True
          Case Else: DGPurchase.AllowUpdate = False
   End Select
Else
   DGPurchase.AllowUpdate = False
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
With MyDDE
    .EditModeReplace = False
    .SetPermissions = UserEditDenied
    Set .BindForm = frmBKK
    .BindFormTAG = "ASM"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT     [Table Journal].JournalID AS [No Bukti], [Table Journal].NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas], [Table Journal].DateTrans AS [Tanggal Bukti], [Table Journal].Periode FROM         [Table Journal] INNER JOIN GlAccount ON [Table Journal].NoAccount = GlAccount.NoAccount WHERE     ([Table Journal].TypeTrans = N'BKK') AND ([Table Journal].Periode = " & mVarPeriode & ") ORDER BY [Table Journal].JournalID"
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

HiasForm Picture1, Me
CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "MASTER PARTNER":
            With MyDDE.ActiveRecordset
                 .Fields("Kode Partner") = mCall.GetFieldByName(0)
                 .Fields("Nama Partner") = mCall.GetFieldByName(1)
            End With
       Case "MASTER KAS":
            With MyDDE.ActiveRecordset
                 .Fields("KOde Kas") = mCall.GetFieldByName(0)
                 .Fields("Nama Kas") = mCall.GetFieldByName(1)
                 lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
            End With
            TotalKas NoVoucher(1)
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
       Case tmbAddNew:
            With MyDDE
                 .GetFieldByName("No Bukti") = MyData.PrepareIndex(tmbTransaksiBKK, 5, "", TglIndex)
            End With
            mVarAdd = True
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
               SimpanDetail
               TotalKas IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")
               mVarAdd = False
            End If
       Case tmbPrint:
            CallRPTReport "Bukti BKK.Rpt", "Select * from [Bukti BKK] Where [No Bukti]='" & lblFixAssets(0) & "'"
End Select
'cmdLink(0).Enabled = mVarAdd
cmdLink(1).Enabled = mVarAdd
lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Opendetail IIf(Not IsNull(MyDDE.GetFieldByName("No Bukti")), MyDDE.GetFieldByName("No Bukti"), "XXXXX")
lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
HitungTotal
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If CekGrid = True And MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If CCur(LblAmount) > CCur(lblTotalKas) Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah kas tidak cukup untuk melakukan transaksi.", "Peringatan", msgOkOnly
                  Else
                     MyDDE.IsChildMemberReady = True
                  End If
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada/belum siap. Harap diisi dulu.", "Peringatan", msgOkOnly
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
       Case tmbDetail:
            If NoVoucher(1) <> "" Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If MyDDE.ChildRecordset.Fields(4) = 0 Then
                     MyDDE.IsChildMemberReady = False
                     MyDDE.CancelTrans = True
                     MessageBox "Jumlah transaksi harus isi.", "Peringatan", msgOkOnly
                  Else
                     MyDDE.IsChildMemberReady = True
                     MyDDE.CancelTrans = False
                  
                  End If
               Else
                  MyDDE.IsChildMemberReady = True
                  MyDDE.CancelTrans = False
               End If
            Else
                MessageBox "Data bank atau kas belum dipilih.", "Peringatan", msgOkOnly
                MyDDE.IsChildMemberReady = False
                MyDDE.CancelTrans = True
            End If

End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [Table Journal]" & _
                     " ( [JournalID],   DateTrans,  NoAccount,TypeTrans,Periode,NoUrut)" & _
                     " VALUES  (N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & NoVoucher(1) & "','BKK'," & mVarPeriode & ", N'" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13) & "')"

    .PrepareUpdate = " UPDATE [Table Journal]" & _
                     " SET DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & _
                     " NoAccount=N'" & NoVoucher(1) & "'" & _
                     " WHERE ([JournalID] = N'" & lblFixAssets(0) & "')"

    .PrepareDelete = " DELETE FROM [Table Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "') "
End With
Err.Clear
End Sub

Private Sub SimpanDetail()
Dim I As Integer
Dim strNotes As String
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [Detail Journal] WHERE     ([JournalID] = N'" & lblFixAssets(0) & "')") = True Then
           .MoveFirst
           strNotes = "Pengeluaran Kas " & lblFixAssets(0)
           I = 1
           Do
             I = I + 1
             If .EOF = True Then Exit Do
             SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Debet,[No]) " & _
                               " VALUES (N'" & lblFixAssets(0) & "', N'" & .Fields("Kode Akun") & "', '" & ValidString(.Fields("Doc Reff")) & "',N'" & .Fields("Keterangan") & "', " & CCur(.Fields("Jumlah Transaksi")) & "," & I & ")")
             strNotes = strNotes & "," & .Fields("Keterangan")
             .MoveNext
           Loop
           SendDataToServer (" INSERT INTO [Detail Journal]  ([JournalID], [NoAccount],[Doc Reff],[Keterangan], Credit,[NO]) " & _
                             " VALUES (N'" & lblFixAssets(0) & "', N'" & NoVoucher(1) & "', '" & NoVoucher(1) & "',N'Pengeluaran Kas', " & CCur(LblAmount) & ",1)")
                                       
           .MoveLast
           SendDataToServer (" UPDATE    [Table Journal] SET  RefNotes = N'" & Left(strNotes, 244) & "' WHERE     (JournalID = N'" & lblFixAssets(0) & "')")
        End If
     End If
End With
End Sub

Private Sub Opendetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen "SELECT     [Detail Journal].NoAccount AS [Kode Akun], GlAccount.AccountName AS [Nama Akun], [Detail Journal].Keterangan, [Detail Journal].[Doc Reff],                       [Detail Journal].Debet AS [Jumlah Transaksi] FROM         [Detail Journal] INNER JOIN                      GlAccount ON [Detail Journal].NoAccount = GlAccount.NoAccount INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID WHERE     ([Detail Journal].JournalID = N'" & ParamString & "') AND ([Detail Journal].Debet <> 0) AND ([Table Journal].TypeTrans = N'BKK') AND                       ([Table Journal].Periode = " & mVarPeriode & ") ORDER BY [Detail Journal].NoAccount", Cnn
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
            RcPartner.DBOpen " select * from [Daftar PartnerBiayaKeluar]", Cnn, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     GlAccount.NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas], ISNULL(ABS(SUM(ISNULL([Tabel Pembantu].CurrentDR" & PeriodeFilter & ", 0)                        + [ListMaster Kas].Debet) - SUM(ISNULL([Tabel Pembantu].CurrentCR" & PeriodeFilter & ", 0) + [ListMaster Kas].Credit)), 0) AS Saldo FROM         [ListMaster Kas] RIGHT OUTER JOIN                       GlAccount ON [ListMaster Kas].NoAccount = GlAccount.NoAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GlAccount.[Group] = N'Detail List Account') AND (GlAccount.Type = N'Kas' OR                       GlAccount.Type = N'Setara Kas') AND ([ListMaster Kas].Periode = " & mVarPeriode & " OR                       [ListMaster Kas].Periode IS NULL) GROUP BY GlAccount.NoAccount, GlAccount.AccountName", Cnn, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT     NoAccount AS [Kode Akun], AccountName AS [Nama Akun] FROM         GlAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'BANK' OR                       Type = N'Piutang Lancar Lain-Lain' OR                       Type = N'Uang Muka Dikeluarkan' OR                       Type = N'Angsuran Dimuka Pajak' OR                       Type = N'Investasi Sementara' OR                       Type = N'Aktiva Lancar Lain-Lain' OR                       Type = N'Investasi Jangka Panjang' OR                       Type = N'Aktiva Lain-Lain' OR                       Type = N'Hutang Lancar Lain-Lain' OR                       Type = N'Modal Usaha' OR                       Type = N'Hutang Jangka Panjang' OR                       Type = N'Biaya Lain-Lain' OR                       Type = N'Uang Muka Diterima') ORDER BY NoAccount", Cnn, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "MASTER PARTNER"
           Case 1:
                mCall.FromTagActive = "MASTER KAS"
                mCall.txtCari = NoVoucher(1)
           Case 2:
            mCall.FromTagActive = "MASTER PERKIRAAN"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
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
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "KK/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Function CekGrid() As Boolean
Dim RcGrd As New DBQuick
Set RcGrd.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
RcGrd.DBRecordset.Filter = "[Jumlah Transaksi] <> 0"
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
            mTotal = mTotal + Avdata(4, I)
        Next I
     Else
        mTotal = 0
     End If
     LblAmount = FormatNumber(Abs(mTotal), 0)
End With
Set Avdata = Nothing
RcTotal.CloseDB
End Sub

Private Function KodeAccount() As String
Dim RcAcc As New DBQuick
RcAcc.DBOpen "SELECT     NoAccount FROM         [Temp Bank] WHERE     (BankID = N'" & NoVoucher(1) & "')", Cnn, lckLockReadOnly
KodeAccount = ""
With RcAcc.DBRecordset
     If .Recordcount <> 0 Then
        KodeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
RcAcc.CloseDB
End Function



