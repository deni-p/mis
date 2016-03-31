VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmClosing 
   Caption         =   "Tutup Buku/Periode"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmClosing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Keluar"
      Height          =   405
      Index           =   2
      Left            =   4905
      TabIndex        =   7
      Top             =   6750
      Width           =   2055
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Proses Closing"
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   2707
      TabIndex        =   6
      Top             =   6750
      Width           =   2055
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Mulai Proses"
      Height          =   405
      Index           =   0
      Left            =   510
      TabIndex        =   5
      Top             =   6750
      Width           =   2055
   End
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
      Height          =   6405
      Left            =   105
      ScaleHeight     =   6375
      ScaleWidth      =   10725
      TabIndex        =   8
      Top             =   180
      Width           =   10755
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   375
         ScaleHeight     =   5505
         ScaleWidth      =   10125
         TabIndex        =   9
         Top             =   660
         Width           =   10155
         Begin MSDataGridLib.DataGrid DGDetail 
            Height          =   2235
            Left            =   120
            TabIndex        =   4
            Tag             =   "Partner"
            Top             =   2865
            Width           =   9900
            _ExtentX        =   17463
            _ExtentY        =   3942
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
            HeadLines       =   2
            RowHeight       =   16
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
            ColumnCount     =   7
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
               Caption         =   "Akun Ref"
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
               DataField       =   "Nama"
               Caption         =   "Nama"
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
            BeginProperty Column05 
               DataField       =   "Debet"
               Caption         =   "Debet"
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
            BeginProperty Column06 
               DataField       =   "Kredit"
               Caption         =   "Kredit"
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
               MarqueeStyle    =   4
               Locked          =   -1  'True
               BeginProperty Column00 
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1590.236
               EndProperty
               BeginProperty Column04 
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  DividerStyle    =   6
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DGHeader 
            Bindings        =   "FrmClosing.frx":08CA
            Height          =   2235
            Left            =   120
            TabIndex        =   3
            Tag             =   "Partner"
            Top             =   585
            Width           =   9900
            _ExtentX        =   17463
            _ExtentY        =   3942
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
            HeadLines       =   2
            RowHeight       =   16
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
               DataField       =   "Tanggal Journal"
               Caption         =   "Tanggal Journal"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "d/MMM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "No Journal"
               Caption         =   "No Journal"
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
               DataField       =   "Doc Reff"
               Caption         =   "Doc Reff"
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
               DataField       =   "PO/SC"
               Caption         =   "PO/SC"
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
               DataField       =   "Status"
               Caption         =   "Status"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Ya"
                  FalseValue      =   "Tidak"
                  NullValue       =   "Tidak"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1679.811
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1950.236
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1814.74
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1440
               EndProperty
            EndProperty
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Unmark All"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   5385
            TabIndex        =   2
            Top             =   195
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Mark All"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3690
            TabIndex        =   1
            Top             =   195
            Width           =   1590
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   135
            Width           =   2445
         End
         Begin VB.Line Line1 
            X1              =   6105
            X2              =   7530
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   7365
            TabIndex        =   12
            Top             =   5145
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   6135
            TabIndex        =   11
            Top             =   5190
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periode"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   165
            TabIndex        =   10
            Top             =   180
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "FrmClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents RcHeader As DBQuick
Attribute RcHeader.VB_VarHelpID = -1
Private RcDetail As New DBQuick
Private mVarOpen As Boolean
Private mVarJournal As Boolean
Private mVarJrl As String
Private LastNoLink As String

Private Sub CmdOK_Click(Index As Integer)
Dim I As Integer
Select Case Index
       Case 0:
            If RcHeader.DBRecordset.Recordcount <> 0 Then
                Option1(0).Enabled = True
                Option1(1).Enabled = True
                CmdOk(0).Enabled = False
                CmdOk(1).Enabled = True
                DGHeader.Columns(4).Button = True
            Else
                MessageBox "Data belum siap untuk melakukan proses closing", "Peringatan", msgOkOnly
            End If
       Case 1:
            If Option1(0).Value = True Then
                I = MessageBox("Anda yakin untuk melakukan proses Closing.", "Closing", msgYesNo)
                If I = 1 Then
                   If AccountLink <> "xxx" Then
                      PrepareJournalFixAssets
                      Closing Combo2.ListIndex + 1
                      DGHeader.Columns(4).Button = False
                      CmdOk(0).Enabled = True
                      CmdOk(1).Enabled = False
                      Option1(0).Enabled = False
                      Option1(1).Enabled = False
                      MessageBox "Proses closing telah selesai.", "Closing", msgOkOnly
                      If PeriodeBerjalan = False Then FrmSetingPeriode.SetFocus
                      Unload Me
                    Else
                      MessageBox "Seting kode akun tampungan rugi laba belum tersedia.", "Kode Akun", msgOkOnly
                    End If
                End If
            Else
                MessageBox "Proses (Mark All) belum dipilih.", "Closing", msgOkOnly
            End If
       Case 2: Unload Me
End Select
End Sub

Private Sub Combo2_Click()
If mVarOpen = False Then OpenDataHeader Combo2.ListIndex + 1
End Sub



Private Sub DgHeader_ButtonClick(ByVal ColIndex As Integer)
If ColIndex = 4 Then
   If CmdOk(1).Enabled = True Then
      If RcHeader.DBRecordset.Recordcount <> 0 Then
         If RcHeader.DBRecordset.Fields(4) = False Then
            RcHeader.DBRecordset.Fields(4) = True
         Else
            RcHeader.DBRecordset.Fields(4) = False
         End If
         DGHeader.Columns(4).Value = CBool(RcHeader.DBRecordset.Fields(4))
      End If
   End If
End If
End Sub

Private Sub DgHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGHeader.Col = 4 Then
   DGHeader.MarqueeStyle = dbgFloatingEditor
Else
   DGHeader.MarqueeStyle = dbgHighlightRowRaiseCell
End If
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
Set RcHeader = New DBQuick
mVarOpen = True
OpenPeriode
mVarOpen = False
OpenDataHeader Combo2.ListIndex + 1
PisahGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

RcHeader.CloseDB
RcDetail.CloseDB
End Sub

Private Sub Form_Resize()

HiasForm Picture1, Me
CenterForm Picture2, Me
CmdOk(0).Left = Picture1.Left + 40
CmdOk(0).Top = Picture1.Height + (CmdOk(0).Height - 350)

CmdOk(1).Left = CmdOk(0).Left + CmdOk(0).Width + 75
CmdOk(1).Top = CmdOk(0).Top

CmdOk(2).Left = CmdOk(1).Left + CmdOk(0).Width + 75
CmdOk(2).Top = CmdOk(0).Top
GridLayout
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmClosing = Nothing
End Sub

Private Sub OpenPeriode()
Dim I As Integer
Combo2.Clear
For I = 1 To 12
    Combo2.AddItem Format(DateSerial(Year(Date), I, 1), "MMMM")
Next I
Combo2.ListIndex = mVarPeriode - 1

End Sub

Private Sub OpenDataHeader(ByVal Periode As Byte)
RcHeader.DBOpen " SELECT     [Table Journal].DateTrans AS [Tanggal Journal], [Table Journal].JournalID AS [No Journal], [Table Journal].TransID AS [Doc Reff],                        [Table Journal].PurchaseID AS [PO/SC], [Table Journal].NoAccount AS [Kode Bank], [Table Journal].PartnerID AS [Kode Partner],                        PartnerDB.CompanyName AS [Nama Partner], [Table Journal].EmpID AS [Kode Karyawan], Employees.FullName AS [Nama Karyawan],                        [Table Journal].Status, [Table Journal].TypeTrans, GlAccount.AccountName AS [Nama Kas]" & _
                " FROM         [Table Journal] INNER JOIN                       GlAccount ON [Table Journal].NoAccount = GlAccount.NoAccount LEFT OUTER JOIN                       Employees ON [Table Journal].EmpID = Employees.EmpID LEFT OUTER JOIN                       PartnerDB ON [Table Journal].PartnerID = PartnerDB.PartnerID WHERE     ([Table Journal].Periode = " & Periode & ") AND ([Table Journal].Status = 0) ORDER BY [Table Journal].DateTrans, [Table Journal].JournalID", Cnn, lckLockBatch
Set DGHeader.DataSource = RcHeader.DBRecordset
End Sub

Private Sub OpenDetail(ByVal NoJournal As String)
RcDetail.DBOpen "SELECT     [Detail Journal].NoAccount AS [No Akun], GlAccount.AccountName AS [Nama Akun], [Detail Journal].Keterangan, [Detail Journal].[Doc Reff],                        [Union SaveMaster].Nama, [Detail Journal].Debet AS Debet, [Detail Journal].Credit AS Kredit FROM         [Detail Journal] INNER JOIN                       GlAccount ON [Detail Journal].NoAccount = GlAccount.NoAccount INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID LEFT OUTER JOIN                       [Union SaveMaster] ON [Detail Journal].[Doc Reff] = [Union SaveMaster].[Doc Reff] WHERE     ([Table Journal].JournalID = N'" & NoJournal & "') ORDER BY [Detail Journal].Debet DESC", Cnn, lckLockReadOnly
Set DGDetail.DataSource = RcDetail.DBRecordset

End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
   MarkAll True
Else
   MarkAll False
End If
End Sub

Private Sub RcHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail RcHeader.Fields("No Journal")
Label2 = FormatNumber(RcDetail.Summary("Debet"), 0)
End Sub

Private Sub PisahGrid()
Dim ukur As Double
Dim Count As Integer
Dim Spl As Splits
Set Spl = DGDetail.Splits
DGDetail.Splits.Add (1)
Spl(1).AllowSizing = False
'Spl(1).LeftCol = 3
Spl(1).MarqueeStyle = dbgHighlightCell
Spl(0).Columns(5).Visible = False
Spl(0).Columns(6).Visible = False
Spl(1).Columns(0).Visible = False
Spl(1).Columns(1).Visible = False
Spl(1).Columns(2).Visible = False
Spl(1).Columns(3).Visible = False
Spl(1).Columns(4).Visible = False
Spl(0).Columns(3).Width = 1250
Spl(1).Columns(5).Width = 1500
Spl(1).Columns(6).Width = 1500
Spl(1).RecordSelectors = False
Spl(0).Size = 2
Spl(1).ScrollBars = dbgBoth
Spl(0).ScrollBars = dbgHorizontal
Spl(0).MarqueeStyle = dbgHighlightRow
Set Spl = Nothing
End Sub

Private Sub MarkAll(ByVal Tipical As Boolean)
Dim rcSet As New DBQuick
Dim I As Long
Dim mBook As Long
Dim mVarData As Variant
Set rcSet.DBRecordset = RcHeader.DBRecordset.Clone(adLockBatchOptimistic)
With rcSet.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            .AbsolutePosition = I + 1
            .Fields("Status") = Tipical
        Next I
     End If
End With
rcSet.CloseDB
End Sub

Private Sub Closing(ByVal PeriodeActive As Integer)
Dim rcSet As New DBQuick
Dim I As Long
Dim mPer As Integer
Dim mBook As Long
Dim mVarData As Variant
Select Case PeriodeActive
       Case 1: mPer = 12
       Case 2: mPer = 1
       Case 3: mPer = 2
       Case 4: mPer = 3
       Case 5: mPer = 4
       Case 6: mPer = 5
       Case 7: mPer = 6
       Case 8: mPer = 7
       Case 9: mPer = 8
       Case 10: mPer = 9
       Case 11: mPer = 10
       Case 12: mPer = 1
End Select
IsiLabaRugi PeriodeActive
'MsgBox "SELECT  [Detail Journal].NoAccount, GlAccount.Period" & mPer & " AS [Saldo Awal], SUM([Detail Journal].Debet) AS Debet, SUM([Detail Journal].Credit) AS Kredit,  ABS(GlAccount.Period" & mPer & " + SUM([Detail Journal].Debet) - SUM([Detail Journal].Credit)) AS Balance FROM         [Detail Journal] INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN                       GlAccount ON [Detail Journal].NoAccount = GlAccount.NoAccount WHERE     ([Table Journal].Periode = " & PeriodeActive & ") GROUP BY [Detail Journal].NoAccount, GlAccount.Period" & mPer & " ORDER BY [Detail Journal].NoAccoun"
rcSet.DBOpen "SELECT  [Detail Journal].NoAccount, GlAccount.Period" & mPer & " AS [Saldo Awal], SUM([Detail Journal].Debet) AS Debet, SUM([Detail Journal].Credit) AS Kredit,  ABS(GlAccount.Period" & mPer & " + SUM([Detail Journal].Debet) - SUM([Detail Journal].Credit)) AS Balance FROM         [Detail Journal] INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN                       GlAccount ON [Detail Journal].NoAccount = GlAccount.NoAccount WHERE     ([Table Journal].Periode = " & PeriodeActive & ") GROUP BY [Detail Journal].NoAccount, GlAccount.Period" & mPer & " ORDER BY [Detail Journal].NoAccount", Cnn, lckLockReadOnly
With rcSet.DBRecordset
     If .Recordcount <> 0 Then
        CompareAccount
        IsiListDataDetail
     End If
End With
SendDataToServer ("Delete from [Table Journal] where JournalID ='LINK'")
'rcSet.CloseDB
End Sub

Private Sub CompareAccount()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData As Variant
RcCom.DBOpen "SELECT GlAccount.NoAccount, [Tabel Pembantu].NoAccount AS NoAccountB FROM         GlAccount LEFT OUTER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount ORDER BY GlAccount.NoAccount", Cnn, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            If IsNull(mVarData(1, I)) Then SendDataToServer ("INSERT INTO [Tabel Pembantu] (NoAccount) VALUES (N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiGroupDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData, mSaldo, mTotalDR, mTotalCR, mCdr, mCcr As Variant
RcCom.DBOpen "SELECT     GlAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GlAccount INNER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GlAccount.[Group] = N'Group Account') GROUP BY GlAccount.GroupAccount", Cnn, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData, mSaldo, mTotalDR, mTotalCR, mCdr, mCcr As Variant
RcCom.DBOpen "SELECT     GlAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GlAccount INNER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GlAccount.[Group] = N'Sub Account') GROUP BY GlAccount.GroupAccount", Cnn, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiListData()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData, mSaldo, mTotalDR, mTotalCR, mCdr, mCcr As Variant
RcCom.DBOpen "SELECT  GlAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GlAccount INNER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GlAccount.[Group] = N'Detail List Account') GROUP BY GlAccount.GroupAccount", Cnn, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
End Sub


Private Sub IsiSubDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData, mSaldo, mTotalDR, mTotalCR, mCdr, mCcr As Variant
RcCom.DBOpen "SELECT  GlAccount.GroupAccount, SUM([Tabel Pembantu].CurrentDR" & mVarPeriode & ") AS DR, SUM([Tabel Pembantu].CurrentCR" & mVarPeriode & ") AS CR FROM         GlAccount INNER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GlAccount.[Group] = N'List Account') GROUP BY GlAccount.GroupAccount", Cnn, lckLockReadOnly
'MessageBox RcCom.PrepareQuery
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(1, I) ' Debet
            mCcr = mVarData(2, I) ' Credit
            mSaldo = mCdr - mCcr
            If mCdr > mCcr Then
               If mSaldo < 0 Then mTotalDR = mSaldo * (-1)
               mTotalCR = 0
            Else
               If mSaldo < 0 Then mTotalCR = mSaldo * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
End Sub

Private Sub IsiListDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData, mTotalDR, mTotalCR, mCdr, mCcr As Variant
Dim mPer As Integer
Select Case mVarPeriode
       Case 1: mPer = 12
       Case 2: mPer = 1
       Case 3: mPer = 2
       Case 4: mPer = 3
       Case 5: mPer = 4
       Case 6: mPer = 5
       Case 7: mPer = 6
       Case 8: mPer = 7
       Case 9: mPer = 8
       Case 10: mPer = 9
       Case 11: mPer = 10
       Case 12: mPer = 1
End Select
RcCom.DBOpen "SELECT  GlAccount.NoAccount, ABS([Tabel Pembantu].SaldoAwalDR" & mPer & " - [Tabel Pembantu].SaldoAwalCR" & mPer & ") AS [Saldo Awal], [Detail Journal].Debet,  [Detail Journal].Credit, [Table Journal].Periode, GlAccount.[Group] FROM GlAccount INNER JOIN [Detail Journal] ON GlAccount.NoAccount = [Detail Journal].NoAccount INNER JOIN [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     ([Table Journal].Periode = " & mVarPeriode & ")", Cnn, lckLockReadOnly
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            mCdr = mVarData(2, I)
            mCcr = mVarData(3, I)
            If mCdr > mCcr Then 'Saldo      Debet             Credit
               mTotalDR = ((mVarData(1, I) + mVarData(2, I)) - mVarData(3, I))
               If mTotalDR < 0 Then mTotalDR = mTotalDR * (-1)
               mTotalCR = 0
            Else                'Saldo      Credit            Debet
               mTotalCR = ((mVarData(1, I) + mVarData(3, I)) - mVarData(2, I))
               If mTotalCR < 0 Then mTotalCR = mTotalCR * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = CurrentDR" & mVarPeriode & " + " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = CurrentCR" & mVarPeriode & " + " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
RcCom.CloseDB
End Sub

Private Sub IsiListDataDetail()
Dim RcCom As New DBQuick
Dim I As Long
Dim mVarData, mSaldo, mTotalDR, mTotalCR, mCdr, mCcr As Variant
Dim mPer As Integer
Select Case mVarPeriode
       Case 1: mPer = 12
       Case 2: mPer = 1
       Case 3: mPer = 2
       Case 4: mPer = 3
       Case 5: mPer = 4
       Case 6: mPer = 5
       Case 7: mPer = 6
       Case 8: mPer = 7
       Case 9: mPer = 8
       Case 10: mPer = 9
       Case 11: mPer = 10
       Case 12: mPer = 1
End Select
RcCom.DBOpen "SELECT GlAccount.NoAccount, [Tabel Pembantu].CurrentDR" & mPer & " - [Tabel Pembantu].CurrentCR" & mPer & " AS [Saldo Awal], SUM([Detail Journal].Debet) AS Debet,SUM([Detail Journal].Credit) AS Credit, [Table Journal].Periode, GlAccount.[Group] FROM         GlAccount INNER JOIN [Detail Journal] ON GlAccount.NoAccount = [Detail Journal].NoAccount INNER JOIN [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID INNER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount GROUP BY GlAccount.NoAccount, [Table Journal].Periode, GlAccount.[Group],[Tabel Pembantu].CurrentDR" & mPer & " , [Tabel Pembantu].CurrentCR" & mPer & " HAVING      ([Table Journal].Periode = " & mVarPeriode & ")", Cnn, lckLockReadOnly
'MsgBox RcCom.PrepareQuery
With RcCom.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(mVarData, 2)
            'Awal Variabel
            mCdr = mVarData(2, I)
            mCcr = mVarData(3, I)
            mSaldo = mVarData(1, I)
            If mSaldo < 0 Then mSaldo = mSaldo * (-1)
            If mCdr > mCcr Then 'Saldo      Debet             Credit
               mTotalDR = ((mSaldo + mVarData(2, I)) - mVarData(3, I))
               'If mTotalDR < 0 Then mTotalDR = mTotalDR * (-1)
               mTotalCR = 0
            Else                'Saldo      Credit            Debet
               mTotalCR = ((mSaldo + mVarData(3, I)) - mVarData(2, I))
               'If mTotalCR < 0 Then mTotalCR = mTotalCR * (-1)
               mTotalDR = 0
            End If
            SendDataToServer ("UPDATE    [Tabel Pembantu] SET  CurrentDR" & mVarPeriode & " = " & CCur(mTotalDR) & ", CurrentCR" & mVarPeriode & " = " & CCur(mTotalCR) & " WHERE     (NoAccount = N'" & mVarData(0, I) & "')")
        Next I
     End If
End With
'RcCom.CloseDB
IsiListData
IsiSubDetail
IsiDetail
IsiGroupDetail
'SendDataToServer ("UPDATE SettingPeriod SET Closed = 1 WHERE (Periode=" & mVarPeriode & ") AND Left([GlFile],4)='" & TahunFiskalYear & "'")
End Sub

Private Sub PrepareJournalFixAssets()
Dim AccJournal As New DBQuick
Dim mVarData As Variant
Dim mVarI As Integer
Dim Kodeku As String
AccJournal.DBOpen " SELECT [TR Aktiva Tetap].DateTrans AS [Tanggal Bukti], [TR Aktiva Tetap].[No FA] AS [No Bukti], [DTR Aktiva Tetap].[No Aktiva] AS [Kode Aktiva],                       [TR Aktiva Tetap].DepAktiva, [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga / [TR Aktiva Tetap].Umur AS [Perolehan Aktiva],                       [TR Aktiva Tetap].BankID AS [Kode Kas], [TR Aktiva Tetap].AccDep,                       [DTR Aktiva Tetap].[Aktiva Beli] * [DTR Aktiva Tetap].Harga / [TR Aktiva Tetap].Umur AS [Kas Keluar]" & _
                  " FROM [TR Aktiva Tetap] INNER JOIN                       [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].Periode = " & mVarPeriode & ")", Cnn, lckLockReadOnly
With AccJournal.DBRecordset
     If .Recordcount <> 0 Then
        mVarData = .Getrows(.Recordcount, adBookmarkFirst)
        For mVarI = 0 To UBound(mVarData, 2)
            Kodeku = IdxAuto
            If SendDataToServer(" INSERT INTO [Table Journal]" & _
                                " (JournalID,TransID, DateTrans, Periode, TypeTrans, RefNotes,Status) " & _
                                " VALUES     (N'" & Kodeku & "',N'" & mVarData(1, mVarI) & "', CONVERT(DATETIME, '" & Format(mVarData(0, mVarI), "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'AKDEP', N'Akumulasi Penyusutan Aktiva',1)") = True Then
                                
               SendDataToServer " INSERT INTO [Detail Journal]" & _
                                " (JournalID, NoAccount,  Debet, Credit, Keterangan) " & _
                                " VALUES     (N'" & Kodeku & "', N'" & mVarData(6, mVarI) & "'," & mVarData(4, mVarI) & ",0,N'Akumulasi Depresiasi " & mVarData(4, mVarI) & "')"
                                
               SendDataToServer " INSERT INTO [Detail Journal]" & _
                                " (JournalID, [Doc Reff], NoAccount, Debet, Credit, Keterangan) " & _
                                " VALUES     (N'" & Kodeku & "',N'" & mVarData(2, mVarI) & "', N'" & mVarData(3, mVarI) & "',0," & mVarData(4, mVarI) & ",N'Akumulasi Depresiasi " & mVarData(4, mVarI) & "')"
                                
            End If
            mVarJournal = True
        Next mVarI
     End If
End With
End Sub

Private Function TglIndex() As String
Dim MyData As New clsTransaksi
Dim TglHari, TglBulan, TglTahun As String
TglIndex = MyData.PrepareIndex(tmbTransaksiAkumDepre, 5, "", "AD/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-")
End Function

Private Function IdxAuto() As String
Dim mNo As Double
Dim mStr As String
If mVarJournal = False Then
   IdxAuto = TglIndex
   mVarJrl = IdxAuto
Else
   mNo = Val(Right(mVarJrl, 5))
   mNo = mNo + 1
   mStr = Left(mVarJrl, 10) & KirimNull(5 - Len(Trim(Str(mNo)))) & Trim(Str(mNo))
   IdxAuto = mStr
End If
End Function

Private Sub GridLayout()
DGHeader.Height = 2235
DGHeader.Width = 9900
DGDetail.Height = 2235
DGDetail.Width = 9900
End Sub

Private Sub IsiLabaRugi(ByVal PeriodeData As Integer)
Dim Rc As New DBQuick
Dim mVarJrl As New clsJournal
Rc.DBOpen "SELECT ABS(SUM([Detail Journal].Debet - [Detail Journal].Credit)) AS Debet FROM  [Detail Journal] INNER JOIN GlAccount ON [Detail Journal].NoAccount = GlAccount.NoAccount INNER JOIN                       [Tabel Pembantu] ON GlAccount.NoAccount = [Tabel Pembantu].NoAccount INNER JOIN                       [Table Journal] ON [Detail Journal].JournalID = [Table Journal].JournalID WHERE     ([Tabel Pembantu].[Kelompok Perkiraan] = 0) AND ([Table Journal].Periode = " & PeriodeData & ")", Cnn, lckLockReadOnly
With Rc
     If .Recordcount <> 0 Then
        If mVarJrl.CiptaKaryaHeaderJournal("LINK", "", "", "", "", "", "IDR", Now(), Trim(Str(PeriodeData)), "LINK") = True Then
           mVarJrl.CiptaKaryaDetailJournal "", AccountLink, "xxx", 0, IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        End If
     Else
     End If
End With
End Sub

Private Function AccountLink() As String
Dim Rc As New DBQuick
Rc.DBOpen "SELECT     NoAccount FROM         [Tabel Pembantu] WHERE     ([Seting Relasi] = 1)", Cnn, lckLockReadOnly
AccountLink = "xxx"
With Rc
     If .Recordcount <> 0 Then
        AccountLink = IIf(Not IsNull(.Fields(0)), .Fields(0), "xxx")
     End If
End With

End Function
