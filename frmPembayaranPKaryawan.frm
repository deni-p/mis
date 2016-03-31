VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C531F5F8-C7B5-4A23-BE73-45A21BBBD9DF}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmPembayaranPKaryawan 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11790
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
   Icon            =   "frmPembayaranPKaryawan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Tag             =   "Pelunasan Piutang Karyawan"
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
      Height          =   6315
      Left            =   90
      ScaleHeight     =   6285
      ScaleWidth      =   11625
      TabIndex        =   7
      Top             =   0
      Width           =   11655
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5685
         Left            =   135
         ScaleHeight     =   5655
         ScaleWidth      =   11355
         TabIndex        =   8
         Top             =   180
         Width           =   11385
         Begin VB.CommandButton cmdLink 
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   10080
            Picture         =   "frmPembayaranPKaryawan.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   105
            Width           =   405
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2760
            Left            =   120
            TabIndex        =   5
            Tag             =   "Voucher"
            Top             =   2310
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   4868
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
               DataField       =   "Jumlah"
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
                  ColumnWidth     =   1590,236
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2340,284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3135,118
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1860,095
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtDate 
            DataField       =   "Tanggal"
            Height          =   315
            Index           =   0
            Left            =   1560
            TabIndex        =   1
            Tag             =   "Voucher"
            Top             =   450
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dddd, dd/MMMM/yyyy"
            Format          =   45678595
            CurrentDate     =   38139
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
            Index           =   7
            Left            =   7470
            TabIndex        =   28
            Top             =   5205
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
            Height          =   255
            Left            =   8715
            TabIndex        =   27
            Top             =   5190
            Width           =   2535
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   7485
            X2              =   8835
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Label LblAngsuran 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
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
            Index           =   6
            Left            =   6780
            TabIndex        =   26
            Top             =   1635
            Width           =   600
         End
         Begin VB.Label LblAngsuran 
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
            DataField       =   "Nama Kas"
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
            Height          =   270
            Index           =   5
            Left            =   6780
            TabIndex        =   25
            Tag             =   "Voucher"
            Top             =   1305
            Width           =   1515
         End
         Begin VB.Label LblAngsuran 
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
            DataField       =   "Kode Kas"
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
            Height          =   270
            Index           =   4
            Left            =   6780
            TabIndex        =   24
            Tag             =   "Voucher"
            Top             =   1005
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   5430
            TabIndex        =   23
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Kas"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   5430
            TabIndex        =   22
            Top             =   1290
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Kas"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   6
            Left            =   5430
            TabIndex        =   21
            Top             =   990
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   5430
            TabIndex        =   20
            Top             =   510
            Width           =   1290
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   5415
            X2              =   6990
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   5430
            X2              =   7005
            Y1              =   405
            Y2              =   405
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   195
            X2              =   1770
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   210
            X2              =   1785
            Y1              =   405
            Y2              =   405
         End
         Begin VB.Label LblAngsuran 
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
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
            Height          =   270
            Index           =   3
            Left            =   1575
            TabIndex        =   19
            Top             =   1965
            Width           =   1515
         End
         Begin VB.Label LblAngsuran 
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
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
            Height          =   270
            Index           =   2
            Left            =   1575
            TabIndex        =   18
            Top             =   1635
            Width           =   1515
         End
         Begin VB.Label LblAngsuran 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
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
            Left            =   1575
            TabIndex        =   17
            Top             =   1305
            Width           =   600
         End
         Begin VB.Label LblAngsuran 
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
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
            Height          =   270
            Index           =   0
            Left            =   1575
            TabIndex        =   16
            Top             =   1005
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Karyawan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   5
            Left            =   5430
            TabIndex        =   15
            Top             =   165
            Width           =   780
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
            Left            =   1575
            TabIndex        =   0
            Tag             =   "Voucher"
            Top             =   105
            Width           =   3705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   14
            Top             =   1290
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Bukti"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   210
            TabIndex        =   13
            Top             =   157
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl. Transaksi"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   12
            Top             =   502
            Width           =   1110
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
            Left            =   6780
            TabIndex        =   2
            Tag             =   "Voucher"
            Top             =   105
            Width           =   3225
         End
         Begin VB.Label NoVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Karyawan"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   6780
            TabIndex        =   4
            Tag             =   "Voucher"
            Top             =   465
            Width           =   3705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jangka Waktu                     Bulan"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   210
            TabIndex        =   11
            Top             =   1950
            Width           =   2850
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Angsuran"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   210
            TabIndex        =   10
            Top             =   1620
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jatuh Tempo"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   210
            TabIndex        =   9
            Top             =   990
            Width           =   1095
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   6480
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmPembayaranPKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private RcPartner As New DBQuick
Private RcDetail As New DBQuick
Private MyData As New clsTransaksi
Private mEdit As Boolean

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
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmPembayaranPKaryawan
        .BindFormTAG = "Voucher"
    Set .ActiveConnection = Cnn
    .PrepareQuery = " SELECT [BKK Karyawan].[No Piutang], [BKK Karyawan].DateTerm AS [Jatuh Tempo], [BKK Karyawan].DateTrans AS Tanggal, [BKK Karyawan].EmpID AS [Kode Karyawan], Employees.FullName AS [Nama Karyawan], [BKK Karyawan].Term, [BKK Karyawan].Notes AS Keterangan,[BKK Karyawan].Jumlah, [BKK Karyawan].NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas] FROM [BKK Karyawan] INNER JOIN Employees ON [BKK Karyawan].EmpID = Employees.EmpID INNER JOIN GlAccount ON [BKK Karyawan].NoAccount = GlAccount.NoAccount WHERE ([BKK Karyawan].TypeTrans = N'BPK') ORDER BY [BKK Karyawan].[No Piutang]"
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
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
'
'HiasForm Picture1, Me
'CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPembayaranPKaryawan = Nothing
End Sub



Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mykey As String
Select Case AdReasonActiveDb
       Case tmbAddNew:
            dtDate(0).Value = CDate(Format(Date, "dd/mm/yyyy"))
            mEdit = True
            MyDDE.GetFieldByName("No Piutang") = MyData.PrepareIndex(tmbTransaksiBayarPiutangKaryawan, 5, "", TglIndex)
            MyDDE.GetFieldByName("Keterangan") = "-"
            MyDDE.GetFieldByName("Tanggal") = dtDate(0).Value
            'MyDDE.GetFieldByName("Jatuh Tempo") = dtDate(1).Value
            MyDDE.GetFieldByName("Term") = 0
            MyDDE.GetFieldByName("Jumlah") = 0
            dtDate(0).SetFocus
       Case tmbEdit:
            mEdit = True
            dtDate(0).SetFocus
       Case tmbSave:
            mEdit = False
            mykey = IdxAuto
            'MsgBox NoVoucher(1)
            If SendDataToServer(" INSERT INTO [Table Journal]" & _
                                " (JournalID, TransID,  NoAccount, PartnerID, Currency, DateTrans,  Periode, TypeTrans,RefNotes,[NoUrut])" & _
                                " VALUES     (N'" & mykey & "', N'" & NoVoucher(0) & "',  N'" & LblAngsuran(4) & "',N'" & NoVoucher(2) & "',  N'IDR', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BPK',N'Pelunasan piutang karyawan dari " & NoVoucher(3) & "','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "/") & "')") = True Then
                                
               SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                 " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan,[No]) " & _
                                 " VALUES   (N'" & mykey & "', N'" & CariTypeAccount(66) & "', N'" & NoVoucher(2) & "', 0, " & CCur(LblAngsuran(2)) & ",N'Pelunasan piutang karyawan dari " & NoVoucher(3) & "',2)")
                                 
               SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                 " (JournalID, NoAccount, [Doc Reff], Debet, Credit,Keterangan,[No]) " & _
                                 " VALUES   (N'" & mykey & "', N'" & LblAngsuran(4) & "', N'xxx', " & CCur(LblAngsuran(2)) & ", 0,N'Pelunasan piutang karyawan dari " & NoVoucher(3) & "',1)")
            End If
       Case tmbCancel:
            mEdit = False
       Case tmbDelete:
            mEdit = False
       Case tmbPrint:
            CallRPTReport "BKM PiutangKaryawan.rpt", "select * from [BkM piutangkaryawan] where [no Piutang] =N'" & NoVoucher(0) & "'"
       Case Else: 'mVarDataDc = False
End Select
cmdLink(0).Enabled = mEdit
LblAngsuran(6) = FormatNumber(TotalKas(LblAngsuran(4)), 0)
HitungTotal
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenPiutangPartner IIf(Not IsNull(MyDDE.GetFieldByName("Kode Karyawan")), MyDDE.GetFieldByName("Kode Karyawan"), 0)

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
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm:
       Case "PIUTANG KARYAWAN":
            With MyDDE
                 .GetFieldByName("Kode Karyawan") = mCall.GetFieldByName("Kode Karyawan")
                 .GetFieldByName("Nama Karyawan") = mCall.GetFieldByName("Nama Karyawan")
                 .GetFieldByName("Kode Kas") = mCall.GetFieldByName("Kode Kas")
                 .GetFieldByName("Nama Kas") = mCall.GetFieldByName("Nama Kas")
                 LblAngsuran(0) = Format(mCall.GetFieldByName(4), "dd/mm/yyyy")
                 LblAngsuran(1) = mCall.GetFieldByName("Keterangan")
                 LblAngsuran(2) = FormatNumber(mCall.GetFieldByName("Angsuran"), 0)
                 LblAngsuran(3) = mCall.GetFieldByName("Jangka Waktu")
                 .GetFieldByName("Jumlah") = mCall.GetFieldByName("Angsuran")
                 LblAngsuran(6) = FormatNumber(TotalKas(mCall.GetFieldByName("Kode Kas")), 0)
            End With
       Case "MASTER KAS":
            With MyDDE
                 .GetFieldByName("Kode Kas") = mCall.GetFieldByName(0)
                 .GetFieldByName("Nama Kas") = mCall.GetFieldByName(1)

            End With '
End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
     .PrepareAppend = " INSERT INTO [BKK Karyawan]" & _
                      " ([No Piutang],NoAccount, Dateterm, DateTrans, EmpID, Term, Notes, Jumlah,typetrans)" & _
                      " VALUES (N'" & NoVoucher(0) & "',N'" & LblAngsuran(4) & "', CONVERT(DATETIME, '" & Format(LblAngsuran(0), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', " & CDbl(LblAngsuran(3)) & ", N'Pelunasan Piutang Karyawan', " & CCur(LblAngsuran(2)) & ",N'BPK')"
     .PrepareUpdate = " UPDATE    [BKK Karyawan]" & _
                      " Set  noAccount =N'" & LblAngsuran(4) & "', dateterm=CONVERT(DATETIME, '" & Format(LblAngsuran(0), "dd/mm/yy") & "', 3),DateTrans = CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), EmpID = N'" & NoVoucher(2) & "', Term = " & CDbl(LblAngsuran(3)) & ", Notes = N'Pelunasan Piutang Karyawan', Jumlah = " & CCur(LblAngsuran(2)) & _
                      " WHERE     ([No Piutang] = N'" & NoVoucher(0) & "') "
     .PrepareDelete = " DELETE FROM [BKK Karyawan] WHERE ([No Piutang] = N'" & NoVoucher(0) & "')"
End With
Err.Clear
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
Select Case mCall.FromTagActive
       Case "PIUTANG KARYAWAN": MyDDE.SetFocus
       Case "MASTER KAS": MyDDE.SetFocus
End Select
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell

Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     [BKK Karyawan].[No Piutang], [BKK Karyawan].EmpID AS [Kode Karyawan], Employees.FullName AS [Nama Karyawan],                       [BKK Karyawan].Notes AS Keterangan, [BKK Karyawan].DateTerm AS [Jatuh Tempo], [BKK Karyawan].Term AS [Jangka Waktu],                       [BKK Karyawan].Jumlah / [BKK Karyawan].Term AS Angsuran, [BKK Karyawan].NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas] FROM         [BKK Karyawan] INNER JOIN                      Employees ON [BKK Karyawan].EmpID = Employees.EmpID INNER JOIN                      GlAccount ON [BKK Karyawan].NoAccount = GlAccount.NoAccount WHERE     ([BKK Karyawan].TypeTrans = N'PPK') AND ([BKK Karyawan].Status = 0)", Cnn, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT BankID as [Kode Kas], NamaBank as [Nama Kas], Amount as [Saldo Kas/Bank]  FROM  [Temp Bank] ORDER BY NamaBank", Cnn, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    
    Select Case Index
           Case 0:
                mCall.FromTagActive = "PIUTANG KARYAWAN"
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
TglIndex = "BK/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub OpenPiutangPartner(ByVal Kodekaryawan As String)
On Error Resume Next
Dim RcCari As New DBQuick
RcCari.DBOpen "SELECT     [BKK Karyawan].Notes, [BKK Karyawan].DateTerm, [BKK Karyawan].Term, [BKK Karyawan].Jumlah / [BKK Karyawan].Term AS Angsuran,                       [BKK Karyawan].NoAccount AS [Kode Kas], GlAccount.AccountName AS [Nama Kas] FROM         [BKK Karyawan] INNER JOIN                      Employees ON [BKK Karyawan].EmpID = Employees.EmpID INNER JOIN                       GlAccount ON [BKK Karyawan].NoAccount = GlAccount.NoAccount WHERE     ([BKK Karyawan].TypeTrans = N'PPK') AND (Employees.EmpID = N'" & Kodekaryawan & "')", Cnn, lckLockReadOnly
With RcCari.DBRecordset
     If .Recordcount <> 0 Then
        LblAngsuran(0) = Format(.Fields("DateTerm"), "dd/mm/yyyy")
        LblAngsuran(1) = .Fields("Notes")
        LblAngsuran(2) = FormatNumber(.Fields("Angsuran"), 0)
        LblAngsuran(3) = .Fields("Term")
'        LblAngsuran(4) = .Fields("Kode kas")
'        LblAngsuran(5) = .Fields("Nama Kas")
     Else
        LblAngsuran(0) = ""
        LblAngsuran(1) = "-"
        LblAngsuran(2) = 0
        LblAngsuran(3) = 0
'        LblAngsuran(0) = ""
'        LblAngsuran(1) = ""
     End If
End With
LblAngsuran(6) = FormatNumber(TotalKas(LblAngsuran(4)), 0)
HitungTotal
Err.Clear
End Sub

Private Sub HitungTotal()
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mTotal As Currency
Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ActiveRecordset.Clone(adLockReadOnly)
mTotal = 0
With RcTotal.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst, "Jumlah")
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

Private Function IdxAuto() As String
IdxAuto = MyData.PrepareIndex(tmbTransaksiBKMKARYAWAN, 5, "", "BK/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-")
End Function

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GlAccount.NoAccount, AccType.ID, GlAccount.AccountName FROM         AccType INNER JOIN                       GlAccount ON AccType.Tipe = GlAccount.Type WHERE     (GlAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", Cnn, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub
