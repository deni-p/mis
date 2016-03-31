VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C531F5F8-C7B5-4A23-BE73-45A21BBBD9DF}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPiutangKaryawan 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPiutangKaryawan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Tag             =   "Pengeluaran Piutang Ke Karyawan"
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
      Height          =   6105
      Left            =   90
      ScaleHeight     =   6075
      ScaleWidth      =   11625
      TabIndex        =   16
      Top             =   0
      Width           =   11655
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5685
         Left            =   120
         ScaleHeight     =   5655
         ScaleWidth      =   11355
         TabIndex        =   17
         Top             =   165
         Width           =   11385
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   4860
            Picture         =   "FrmPiutangKaryawan.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1170
            Width           =   405
         End
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   9975
            Picture         =   "FrmPiutangKaryawan.frx":6BDC
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   150
            Width           =   405
         End
         Begin MSComCtl2.DTPicker dtDate 
            DataField       =   "Jatuh Tempo"
            Height          =   315
            Index           =   1
            Left            =   1545
            TabIndex        =   2
            Tag             =   "Voucher"
            Top             =   825
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MMMM/yyyy"
            Format          =   45547523
            CurrentDate     =   38139
         End
         Begin MSComCtl2.DTPicker dtDate 
            DataField       =   "Tanggal"
            Height          =   315
            Index           =   0
            Left            =   1545
            TabIndex        =   1
            Tag             =   "Voucher"
            Top             =   480
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MMMM/yyyy"
            Format          =   45547523
            CurrentDate     =   38139
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1935
            Left            =   105
            TabIndex        =   14
            Tag             =   "Voucher"
            Top             =   3285
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
               DataField       =   "No Piutang"
               Caption         =   "No Piutang"
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
               DataField       =   "Kode Karyawan"
               Caption         =   "Kode Karyawan"
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
               DataField       =   "Nama Karyawan"
               Caption         =   "Nama Karyawan"
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
            BeginProperty Column04 
               DataField       =   "JmlTemp"
               Caption         =   "Jumlah"
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
                  Alignment       =   1
                  ColumnWidth     =   1830,047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1769,953
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1904,882
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3225,26
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1860,095
               EndProperty
            EndProperty
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "JmlTemp"
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
            Left            =   1545
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "Voucher"
            Top             =   2205
            Width           =   2355
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "Term"
            Height          =   330
            Left            =   1545
            MaxLength       =   5
            TabIndex        =   8
            Tag             =   "Voucher"
            Top             =   2565
            Width           =   2355
         End
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            DataField       =   "Keterangan"
            Height          =   330
            Left            =   1545
            TabIndex        =   6
            Tag             =   "Voucher"
            Top             =   1845
            Width           =   3750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   9
            Left            =   180
            TabIndex        =   31
            Top             =   1575
            Width           =   1290
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   180
            X2              =   1875
            Y1              =   1815
            Y2              =   1815
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   150
            X2              =   1845
            Y1              =   3225
            Y2              =   3225
         End
         Begin VB.Label LblAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Kode Kas"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8025
            TabIndex        =   30
            Top             =   5250
            Width           =   3255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   8
            Left            =   6855
            TabIndex        =   29
            Top             =   5295
            Width           =   420
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   6855
            X2              =   8550
            Y1              =   5550
            Y2              =   5550
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   5
            Left            =   6675
            TabIndex        =   13
            Top             =   795
            Width           =   3750
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   5490
            X2              =   7185
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   7
            Left            =   5550
            TabIndex        =   28
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Kas"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   6
            Left            =   5550
            TabIndex        =   27
            Top             =   510
            Width           =   780
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   5520
            X2              =   7215
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Kas"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   4
            Left            =   6675
            TabIndex        =   12
            Tag             =   "Voucher"
            Top             =   480
            Width           =   3750
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   5520
            X2              =   7215
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Kas"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   5550
            TabIndex        =   26
            Top             =   195
            Width           =   750
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
            Left            =   6675
            TabIndex        =   10
            Tag             =   "Voucher"
            Top             =   150
            Width           =   3255
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   195
            X2              =   1890
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   195
            X2              =   1890
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   165
            X2              =   1860
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   180
            X2              =   1875
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   165
            X2              =   1860
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   150
            X2              =   1845
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jatuh Tempo"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   195
            TabIndex        =   25
            Top             =   885
            Width           =   1095
         End
         Begin VB.Label Angsuran 
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
            Left            =   1545
            TabIndex        =   9
            Top             =   2910
            Width           =   2355
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Angsuran"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   195
            TabIndex        =   24
            Top             =   2970
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   195
            TabIndex        =   23
            Top             =   2250
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jangka Waktu                                             Bulan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   195
            TabIndex        =   22
            Top             =   2610
            Width           =   4290
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Karyawan"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   3
            Left            =   1545
            TabIndex        =   5
            Tag             =   "Voucher"
            Top             =   1530
            Width           =   3750
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Kode Karyawan"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   1545
            TabIndex        =   3
            Tag             =   "Voucher"
            Top             =   1185
            Width           =   3255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Transaksi"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   195
            TabIndex        =   21
            Top             =   540
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Bukti"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   195
            TabIndex        =   20
            Top             =   195
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   195
            TabIndex        =   19
            Top             =   1890
            Width           =   945
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "No Piutang"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1545
            TabIndex        =   0
            Tag             =   "Voucher"
            Top             =   150
            Width           =   3750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Karyawan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   5
            Left            =   195
            TabIndex        =   18
            Top             =   1230
            Width           =   1260
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   150
            X2              =   1845
            Y1              =   1140
            Y2              =   1140
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   15
      Top             =   6375
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   1005
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

Private Sub dtDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
Set mCall = New frmCaller
dtDate(0).Value = Date
dtDate(1).Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmPiutangKaryawan
        .BindFormTAG = "Voucher"
    Set .ActiveConnection = Cnn
    .PrepareQuery = " SELECT     [BKK Karyawan].[No Piutang], [BKK Karyawan].DateTerm AS [Jatuh Tempo], [BKK Karyawan].DateTrans AS Tanggal,                        [BKK Karyawan].EmpID AS [Kode Karyawan], Employees.FullName AS [Nama Karyawan], [BKK Karyawan].Term, [BKK Karyawan].Notes AS Keterangan,                        [BKK Karyawan].JmlTemp, [BKK Karyawan].NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas] FROM         [BKK Karyawan] INNER JOIN                       Employees ON [BKK Karyawan].EmpID = Employees.EmpID INNER JOIN                       GlAccount ON [BKK Karyawan].NoAccount = GlAccount.NoAccount WHERE     ([BKK Karyawan].TypeTrans = N'PPK') ORDER BY [BKK Karyawan].[No Piutang]"
End With
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

'HiasForm Picture1, Me
'CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPiutangKaryawan = Nothing
End Sub



Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mykey As String
Select Case AdReasonActiveDb
       Case tmbAddNew:
            mEdit = True
            dtDate(0).Value = CDate(Format(Date, "dd/mm/yyyy"))
            MyDDE.GetFieldByName("No Piutang") = MyData.PrepareIndex(tmbTransaksiPiutangKaryawan, 5, "", TglIndex)
            MyDDE.GetFieldByName("Keterangan") = "Pengeluaran Piutang Karyawan"
            MyDDE.GetFieldByName("Tanggal") = dtDate(0).Value
            MyDDE.GetFieldByName("Jatuh Tempo") = dtDate(1).Value
            MyDDE.GetFieldByName("Term") = 1
            MyDDE.GetFieldByName("Jumlah") = 0
            MyDDE.GetFieldByName("JmlTemp") = 0
            dtDate(0).SetFocus
       Case tmbEdit:
            mEdit = True
            dtDate(0).SetFocus
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               mykey = IdxAuto
               'SendVoucher NoVoucher(0), NoVoucher(2), ValidString(txtNotes), dtDate(0).Value, CCur(Text2), 0, "", "PK"
               If SendDataToServer(" INSERT INTO [Table Journal]" & _
                                   " (JournalID, TransID,  NoAccount, PartnerID, Currency, DateTrans,  Periode, TypeTrans,refnotes,[NoUrut])" & _
                                   " VALUES     (N'" & mykey & "', N'" & NoVoucher(0) & "',  N'" & NoVoucher(1) & "',N'" & NoVoucher(2) & "',  N'IDR', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'PPK','Pengeluaran piutang karyawan ke " & NoVoucher(3) & "','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "/") & "')") = True Then
                  SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                    " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan,[No]) " & _
                                    " VALUES   (N'" & mykey & "', N'" & CariTypeAccount(66) & "', N'" & NoVoucher(2) & "', " & CCur(Text2) & ", 0,N'Pengeluaran piutang karyawan ke " & NoVoucher(3) & "',1)")
                                    
                  SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                    " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan,[No]) " & _
                                    " VALUES   (N'" & mykey & "', N'" & NoVoucher(1) & "', N'xxx', 0, " & CCur(Text2) & ",N'Pengeluaran piutang karyawan ke " & NoVoucher(3) & "',2)")
               End If
               mEdit = False
            End If
       Case tmbCancel:
            mEdit = False
       Case tmbDelete:
            mEdit = False
       Case tmbPrint:
            CallRPTReport "BKK PiutangKaryawan.rpt", "select * from [Bkk piutangkaryawan] where [no Piutang] =N'" & NoVoucher(0) & "'"
       Case Else: 'mVarDataDc = False
End Select
cmdLink(0).Enabled = mEdit
cmdLink(1).Enabled = mEdit
SaldoKas IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "xxxx")
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If IIf(Not IsNull(MyDDE.GetFieldByName("JMLTEMP")), MyDDE.GetFieldByName("JMLTEMP"), 0) <> 0 And IIf(Not IsNull(MyDDE.GetFieldByName("Term")), MyDDE.GetFieldByName("Term"), 0) <> 0 Then
   Angsuran = FormatNumber(CDbl(MyDDE.GetFieldByName("JMLTEMP")) / CDbl(MyDDE.GetFieldByName("Term")), 0)
Else
   Angsuran = FormatNumber(MyDDE.GetFieldByName("JMLTEMP"), 0)
End If
SaldoKas IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "xxxx")

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               'If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
              '    PrepareQuery
              ' Else
              '    MyDDE.CancelTrans = True
              '    MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
              '    MyDDE.IsChildMemberReady = False
              ' End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False And Val(Text2) <> 0 And Val(NoVoucher(5)) <> 0 Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MessageBox "Data transaksi belum lengkap. Harap diperiksa dulu", "Peringatan", msgOkOnly
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
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
            SaldoKas IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "xxxx")
End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
     .PrepareAppend = " INSERT INTO [BKK Karyawan]" & _
                      " ([No Piutang],NoAccount, Dateterm, DateTrans, EmpID, Term, Notes, Jumlah,JmlTemp)" & _
                      " VALUES (N'" & NoVoucher(0) & "',N'" & NoVoucher(1) & "', CONVERT(DATETIME, '" & Format(dtDate(1).Value, "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', " & CDbl(Text1) & ", N'" & ValidString(txtNotes) & "', " & CCur(Text2) & ", " & CCur(Text2) & ")"
                      
     .PrepareUpdate = " UPDATE    [BKK Karyawan]" & _
                      " Set  NoAccount=N'" & NoVoucher(1) & "',dateterm=CONVERT(DATETIME, '" & Format(dtDate(1).Value, "dd/mm/yy") & "', 3),DateTrans = CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), EmpID = N'" & NoVoucher(2) & "', Term = " & CDbl(Text1) & ", Notes = N'" & ValidString(txtNotes) & "',JmlTemp = " & CCur(Text2) & ", Jumlah = " & CCur(Text2) & _
                      " WHERE     ([No Piutang] = N'" & NoVoucher(0) & "') "
          ' MessageBox .PrepareAppend
     .PrepareDelete = " DELETE FROM [BKK Karyawan] WHERE ([No Piutang] = N'" & NoVoucher(0) & "')"
End With
Err.Clear
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
Select Case mCall.FromTagActive
       Case "MASTER KARYAWAN": If txtNotes.Enabled = True Then txtNotes.SetFocus
       Case "MASTER KAS": MyDDE.SetFocus
End Select
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     EmpID AS [Kode Karyawan], FullName AS [Nama Karyawan] FROM         Employees ORDER BY EmpID", Cnn, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     GlAccount.NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas], ISNULL(ABS(SUM(ISNULL([Tabel Pembantu].CurrentDR" & PeriodeFilter & ", 0)                        + [ListMaster Kas].Debet) - SUM(ISNULL([Tabel Pembantu].CurrentCR" & PeriodeFilter & ", 0) + [ListMaster Kas].Credit)), 0) AS Saldo FROM         [ListMaster Kas] RIGHT OUTER JOIN                       GlAccount ON [ListMaster Kas].NoAccount = GlAccount.NoAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GlAccount.[Group] = N'Detail List Account') AND (GlAccount.Type = N'Kas' OR                       GlAccount.Type = N'Setara Kas') AND ([ListMaster Kas].Periode = " & mVarPeriode & " OR                       [ListMaster Kas].Periode IS NULL) GROUP BY GlAccount.NoAccount, GlAccount.AccountName", Cnn, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0:
                mCall.FromTagActive = "MASTER KARYAWAN"
                mCall.txtCari = NoVoucher(2)
           Case 1:
                mCall.FromTagActive = "MASTER KAS"
                mCall.txtCari = NoVoucher(1)
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If
Exit Sub
Hell:
    Err.Clear
End Sub
Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "PK/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub Text1_Change()
On Error Resume Next
If mEdit = True Then
   If Text1 = "" Or Text1 = "0" Then Text1 = 1
   If Val(Text1) <> 0 Then MyDDE.GetFieldByName("Jatuh Tempo") = DateAdd("m", Val(Text1), dtDate(0).Value)
    If CDbl(Text2) <> 0 And CDbl(Text1) <> 0 Then
       Angsuran = FormatNumber(CDbl(Text2) / CDbl(Text1), 0)
    Else
       Angsuran = FormatNumber(CDbl(Text2), 0)
    End If
End If
Err.Clear
End Sub

Private Sub Text1_GotFocus()
Block Text1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
ValidNum KeyAscii
End Sub

Private Sub Text2_Change()
If mEdit = True Then
    If Text2 <> "" Then
       If CDbl(Text2) <> 0 And CDbl(Text1) <> 0 Then
          Angsuran = FormatNumber(CDbl(Text2) / CDbl(Text1), 0)
       Else
          Angsuran = FormatNumber(CDbl(Text2), 0)
       End If
       MyDDE.GetFieldByName("Jumlah") = CCur(Text2)
    Else
       Text2 = 0
       MyDDE.GetFieldByName("Jumlah") = 0
    End If
End If
End Sub

Private Sub Text2_GotFocus()
Block Text2
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
ValidNum KeyAscii
End Sub

Private Sub txtNotes_GotFocus()
Block txtNotes
End Sub

Private Sub txtNotes_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Function IdxAuto() As String
IdxAuto = MyData.PrepareIndex(tmbTransaksiBKKKARYAWAN, 5, "", "PK/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-")
End Function

Private Sub SaldoKas(ByVal NoAccount As String)
Dim mTotalDR, mTotalCR As Variant
Dim RcKas As New DBQuick
RcKas.DBOpen " SELECT  SUM([Tabel Pembantu].CurrentDR" & PeriodeFilter & ") AS [Saldo Awal DR], SUM([Tabel Pembantu].CurrentCR" & PeriodeFilter & ") AS [Saldo Awal CR], SUM([Detail Journal].Debet) AS Debet,  SUM([Detail Journal].Credit) AS Kredit FROM GlAccount INNER JOIN" & _
             " [Detail Journal] ON GlAccount.NoAccount = [Detail Journal].NoAccount INNER JOIN [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID LEFT OUTER JOIN [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     ([Detail Journal].NoAccount = N'" & NoAccount & "') AND ([Table Journal].Periode = " & mVarPeriode & ")", Cnn, lckLockReadOnly
With RcKas
     If .Recordcount <> 0 Then
        mTotalDR = IIf(Not IsNull(.Fields(0)), .Fields(0), 0) + IIf(Not IsNull(.Fields(2)), .Fields(2), 0)
        mTotalCR = IIf(Not IsNull(.Fields(1)), .Fields(1), 0) + IIf(Not IsNull(.Fields(3)), .Fields(3), 0)
     Else
        mTotalDR = 0
        mTotalCR = 0
     End If
     If mTotalDR > mTotalCR Then
        NoVoucher(5) = FormatNumber(mTotalDR - mTotalCR, 0)
     Else
        NoVoucher(5) = FormatNumber(mTotalCR - mTotalDR, 0)
     End If
End With
HitungTotal
End Sub

Private Sub HitungTotal()
Dim RcTotal As New Recordset
Dim Avdata As Variant
Dim mTotal As Currency
Dim I As Long
RcTotal.CursorLocation = adUseClient
Set RcTotal = MyDDE.ActiveRecordset.Clone(adLockReadOnly)
mTotal = 0
With RcTotal
        If .Recordcount <> 0 Then
           Avdata = .Getrows(.Recordcount, adBookmarkFirst, "JmlTemp")
           For I = 0 To UBound(Avdata, 2)
               mTotal = mTotal + IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), 0)
           Next I
        Else
           mTotal = 0
        End If
        LblAmount = FormatNumber(mTotal, 0)
End With
Set Avdata = Nothing
End Sub

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GlAccount.NoAccount, AccType.ID, GlAccount.AccountName FROM         AccType INNER JOIN                       GlAccount ON AccType.Tipe = GlAccount.Type WHERE     (GlAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", Cnn, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function
