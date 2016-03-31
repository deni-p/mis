VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPiutangKaryawan 
   Caption         =   "Transaksi Piutang Karyawan"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPiutangKaryawanssss.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
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
      Height          =   6630
      Left            =   90
      ScaleHeight     =   6570
      ScaleWidth      =   11025
      TabIndex        =   13
      Top             =   0
      Width           =   11085
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         Height          =   5625
         Left            =   105
         ScaleHeight     =   5565
         ScaleWidth      =   10740
         TabIndex        =   14
         Top             =   270
         Width           =   10800
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            DataField       =   "Jumlah"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0;(#,##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   330
            Left            =   1245
            MaxLength       =   15
            TabIndex        =   4
            Tag             =   "Voucher"
            Top             =   2040
            Width           =   2355
         End
         Begin VB.TextBox Text1 
            DataField       =   "Term"
            Height          =   330
            Left            =   6465
            MaxLength       =   5
            TabIndex        =   7
            Tag             =   "Voucher"
            Top             =   915
            Width           =   1590
         End
         Begin VB.TextBox txtNotes 
            DataField       =   "Keterangan"
            Height          =   330
            Left            =   1245
            TabIndex        =   3
            Tag             =   "Voucher"
            Top             =   1695
            Width           =   3750
         End
         Begin SemeruDC.SemeruButton cmdLink 
            Height          =   315
            Index           =   0
            Left            =   4545
            TabIndex        =   1
            ToolTipText     =   "VIEW MASTER KAS"
            Top             =   285
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   3
            FOCUSR          =   0   'False
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmPiutangKaryawanssss.frx":08CA
            PICN            =   "FrmPiutangKaryawanssss.frx":08E6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SemeruDC.SemeruButton cmdLink 
            Height          =   315
            Index           =   1
            Left            =   10020
            TabIndex        =   9
            ToolTipText     =   "VIEW MASTER PARTNER"
            Top             =   1470
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   3
            FOCUSR          =   0   'False
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmPiutangKaryawanssss.frx":1B68
            PICN            =   "FrmPiutangKaryawanssss.frx":1B84
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker dtDate 
            DataField       =   "Tanggal"
            Height          =   315
            Index           =   0
            Left            =   6465
            TabIndex        =   6
            Tag             =   "Voucher"
            Top             =   555
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd, dd/MMMM/yyyy"
            Format          =   58327043
            CurrentDate     =   38139
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Height          =   2580
            Left            =   165
            TabIndex        =   11
            Tag             =   "Partner"
            Top             =   2640
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   4551
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "Angsuran"
               Caption         =   "Angsuran"
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
            BeginProperty Column01 
               DataField       =   "Cash Reff"
               Caption         =   "Cash Reff"
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
               DataField       =   "Cash"
               Caption         =   "Cash"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Tunai"
                  FalseValue      =   "Check"
                  NullValue       =   "Check"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               AllowRowSizing  =   -1  'True
               AllowSizing     =   -1  'True
               BeginProperty Column00 
                  Alignment       =   1
                  ColumnWidth     =   3314.835
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3899.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1665.071
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
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
            Left            =   495
            TabIndex        =   22
            Top             =   2085
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Term                                Bulan"
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
            Left            =   5895
            TabIndex        =   21
            Top             =   1005
            Width           =   2880
         End
         Begin VB.Label NoVoucher 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Karyawan"
            Enabled         =   0   'False
            Height          =   1035
            Index           =   3
            Left            =   1245
            TabIndex        =   2
            Tag             =   "Voucher"
            Top             =   615
            Width           =   3750
         End
         Begin VB.Label NoVoucher 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Kode Karyawan"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   1245
            TabIndex        =   0
            Tag             =   "Voucher"
            Top             =   270
            Width           =   3255
         End
         Begin VB.Label NoVoucher 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Kode Kas"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   6465
            TabIndex        =   8
            Tag             =   "Voucher"
            Top             =   1470
            Width           =   3510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Transaksi"
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
            Left            =   5115
            TabIndex        =   20
            Top             =   615
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Bukti"
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
            Left            =   5505
            TabIndex        =   19
            Top             =   255
            Width           =   840
         End
         Begin VB.Label Label2 
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
            Index           =   0
            Left            =   75
            TabIndex        =   18
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label NoVoucher 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "No Piutang"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6465
            TabIndex        =   5
            Tag             =   "Voucher"
            Top             =   210
            Width           =   4020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kas/Bank"
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
            Left            =   5475
            TabIndex        =   17
            Top             =   1500
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Karyawan"
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
            Left            =   240
            TabIndex        =   16
            Top             =   315
            Width           =   900
         End
         Begin VB.Label lblAlamatBank 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Kas"
            ForeColor       =   &H80000010&
            Height          =   315
            Left            =   6465
            TabIndex        =   10
            Tag             =   "Voucher"
            Top             =   1815
            Width           =   4020
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
            Left            =   6465
            TabIndex        =   15
            Top             =   2205
            Width           =   120
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   12
      Top             =   6780
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1217
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPiutangKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner As New DBQuick
Private mEdit As Boolean
Private MyData As New clsTransaksi
Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mEdit = True Then
   DGPurchase.AllowUpdate = True
Else
   DGPurchase.AllowUpdate = False
End If
End Sub

Private Sub dtDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
dtDate(0).Value = dDateBegin
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmPiutangKaryawan
    .BindFormTAG = "Voucher"
    '.SetPermissions = UserEditDeleteDenied
    Set .ActiveConnection = Cnn
    .PrepareQuery = " SELECT     [BKK Karyawan].[No Piutang], [BKK Karyawan].BankID AS [Kode Kas], [Temp Bank].NamaBank AS [Nama Kas], [BKK Karyawan].DateTrans AS Tanggal,                        [BKK Karyawan].EmpID AS [Kode Karyawan], Employees.FullName AS [Nama Karyawan], [BKK Karyawan].Term, [BKK Karyawan].Notes AS Keterangan,                        [BKK Karyawan].Jumlah FROM         [BKK Karyawan] INNER JOIN                       Employees ON [BKK Karyawan].EmpID = Employees.EmpID INNER JOIN                       [Temp Bank] ON [BKK Karyawan].BankID = [Temp Bank].BankID ORDER BY [BKK Karyawan].[No Piutang]"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MyDDE.ClearRecordset
RcPartner.CloseDB
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMaximized Then
   Me.Height = MainMenu.ScaleHeight
   Me.Width = MainMenu.ScaleWidth
End If
HiasForm Picture1, Me
CenterForm Picture2
Err.Clear
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm:
       Case "MASTER KARYAWAN":
            With MyDDE
                 .GetFieldByName("Kode Karyawan") = mCall.GetFieldByName(0)
                 .GetFieldByName("Nama Karyawan") = mCall.GetFieldByName(1)
            End With
       Case "MASTER KAS":
            With MyDDE
                 .GetFieldByName("Kode Kas") = mCall.GetFieldByName(0)
                 .GetFieldByName("Nama Kas") = mCall.GetFieldByName(1)
            End With
            TotalKas NoVoucher(1)
End Select
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     EmpID AS [Kode Karyawan], FullName AS [Nama Karyawan] FROM         Employees ORDER BY EmpID", Cnn, lckLockReadOnly
            mCall.FromTagActive = "MASTER KARYAWAN"
            mCall.txtCari = NoVoucher(2)
       Case 1:
            RcPartner.DBOpen " SELECT BankID as [Kode Kas], NamaBank as [Nama Kas], Amount as [Saldo Kas/Bank]  FROM  [Temp Bank] ORDER BY NamaBank", Cnn, lckLockReadOnly
            mCall.FromTagActive = "MASTER KAS"
            mCall.txtCari = NoVoucher(1)
End Select
If RcPartner.Recordcount <> 0 Then
    Set mCall.FormData = RcPartner.DBRecordset
'    If Index = 0 Then
'       mCall.SetFormat(2) = "#,##0"
'       mCall.SetAlignmentFormat(2) = 1
'    Else
'       OpenDetail NoVoucher(2)
'    End If
    mCall.Show vbModal
    
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If
Set mCall = Nothing
Exit Sub
Hell:
    'MsgBox Err.Description
    Err.Clear
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("No Piutang")
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit, tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = False
            Else
               MyDDE.CancelTrans = True
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            mEdit = True
       Case tmbAddNew:
            mEdit = True
            MyDDE.GetFieldByName("No Piutang") = MyData.PrepareIndex(tmbTransaksiPiutangKarywan, 5, "", TglIndex)
            MyDDE.GetFieldByName("Keterangan") = "-"
            MyDDE.GetFieldByName("Term") = 0
            MyDDE.GetFieldByName("Jumlah") = 0
       Case tmbDetail:
            MyDDE.ChildRecordset.Fields("Cash Reff") = "-"
            MyDDE.ChildRecordset.Fields("Cash") = 0
            MyDDE.ChildRecordset.Fields("Angsuran") = 0
            mEdit = True
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               mEdit = False
            End If
       Case tmbCancel:
            mEdit = False
       Case tmbPrint:
End Select
cmdLink(0).Enabled = mEdit
cmdLink(1).Enabled = mEdit
DGPurchase.Columns(2).Button = mEdit
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
     .PrepareAppend = " INSERT INTO [BKK Karyawan]" & _
                      " ([No Piutang], BankID, DateTrans, EmpID, Term, Notes, Jumlah)" & _
                      " VALUES (N'" & NoVoucher(0) & "', N'" & NoVoucher(1) & "', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', " & CDbl(Text1) & ", N'" & ValidString(txtNotes) & "', " & CCur(Text2) & ")"
     .PrepareUpdate = " UPDATE    [BKK Karyawan]" & _
                      " Set BankID = N'" & NoVoucher(1) & "', DateTrans = CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), EmpID = N'" & NoVoucher(2) & "', Term = " & CDbl(Text1) & ", Notes = N'" & ValidString(txtNotes) & "', Jumlah = " & CCur(Text2) & _
                      " WHERE     ([No Piutang] = N'" & NoVoucher(0) & "') "
     .PrepareDelete = " DELETE FROM [BKK Karyawan] WHERE ([No Piutang] = N'" & NoVoucher(0) & "')"
End With
Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
Dim RcDetail As New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
RcDetail.DBOpen "SELECT     Angsuran, [Cash Reff], Cash FROM         [Detail PKaryawan] WHERE     ([No Piutang] = N'" & ParameterString & "')", Cnn, lckLockBatch
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "PK/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub TotalKas(ByVal NoBankID As String)
Dim RcKas As New DBQuick
RcKas.DBOpen "SELECT Amount FROM [Temp Bank] WHERE (BankID = N'" & NoBankID & "')", Cnn, lckLockReadOnly
lblTotalKas = 0
With RcKas
     If .Recordcount <> 0 Then
        lblTotalKas = FormatNumber(.Fields(0), 0)
     End If
End With
RcKas.CloseDB
End Sub

Private Sub SimpanDetail()
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        SendDataToServer ("DELETE FROM  [Detail PKaryawan] WHERE     ([No Piutang] = N'" & ValidString(NoVoucher(0)) & "')")
        .MoveFirst
        Do
         If .EOF Then Exit Do
         If CBool(.Fields("Cash")) = True Then
            SendDataToServer ("INSERT INTO [Detail PKaryawan]  ([No Piutang], Angsuran, [Cash Reff], Cash) VALUES (N'" & ValidString(NoVoucher(0)) & "'," & .Fields("Angsuran") & ", N'" & .Fields("Cash Reff") & "',1)")
         Else
            SendDataToServer ("INSERT INTO [Detail PKaryawan]  ([No Piutang], Angsuran, [Cash Reff], Cash) VALUES (N'" & ValidString(NoVoucher(0)) & "'," & .Fields("Angsuran") & ", N'" & .Fields("Cash Reff") & "', 0)")
         End If
        .MoveNext
        Loop
        .MoveLast
     End If
End With
End Sub

Private Sub Text1_GotFocus()
Block Text1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Text2_GotFocus()
Block Text2
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub txtNotes_GotFocus()
Block txtNotes
End Sub

Private Sub txtNotes_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub
