VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPenjualanFixAssets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penjualan Fixed Asset"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPenjualanFixAssets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Tag             =   "Asset Sales"
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
      Height          =   6045
      Left            =   90
      ScaleHeight     =   6015
      ScaleWidth      =   10425
      TabIndex        =   14
      Top             =   15
      Width           =   10455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5370
         Left            =   105
         ScaleHeight     =   5340
         ScaleWidth      =   10095
         TabIndex        =   15
         Top             =   225
         Width           =   10125
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "NoCheque"
            Height          =   330
            Index           =   1
            Left            =   6705
            MaxLength       =   15
            TabIndex        =   11
            Tag             =   "ASM"
            Top             =   1710
            Width           =   2835
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            DataField       =   "DP"
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
            Index           =   0
            Left            =   6705
            MaxLength       =   12
            TabIndex        =   10
            Tag             =   "ASM"
            Top             =   1365
            Width           =   2835
         End
         Begin MSDataGridLib.DataGrid DGPurchase 
            Bindings        =   "FrmPenjualanFixAssets.frx":6852
            Height          =   2295
            Left            =   105
            TabIndex        =   12
            Top             =   2670
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   4048
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "Doc Reff"
               Caption         =   "No Bukti"
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
               DataField       =   "No Aktiva"
               Caption         =   "No Aktiva"
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
               DataField       =   "Nama Aktiva"
               Caption         =   "Nama Aktiva"
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
               DataField       =   "Aktiva Jual"
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
               DataField       =   "Ppn"
               Caption         =   "Ppn"
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
               DataField       =   "Harga"
               Caption         =   "Harga Perolehan"
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
               DataField       =   "Harga Jual"
               Caption         =   "Harga Jual"
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
                  Alignment       =   1
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "Tanggal"
            Height          =   315
            Left            =   1800
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   450
            Width           =   3525
            _ExtentX        =   6218
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
            CustomFormat    =   "dddd dd/MMMM/yyyy"
            Format          =   61276163
            CurrentDate     =   38272
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   9135
            Picture         =   "FrmPenjualanFixAssets.frx":6867
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   690
            Width           =   405
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   4920
            Picture         =   "FrmPenjualanFixAssets.frx":6BF1
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   780
            Width           =   405
         End
         Begin VB.CommandButton cmdLInk 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   4920
            Picture         =   "FrmPenjualanFixAssets.frx":6F7B
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1470
            Width           =   405
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   5505
            X2              =   7185
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   5505
            X2              =   7185
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Cek/BG"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   5
            Left            =   5520
            TabIndex        =   23
            Top             =   1770
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uang Muka"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   4
            Left            =   5520
            TabIndex        =   22
            Top             =   1425
            Width           =   900
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   270
            X2              =   1950
            Y1              =   1785
            Y2              =   1785
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   270
            X2              =   1950
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   270
            X2              =   1950
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Group"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   300
            TabIndex        =   21
            Top             =   1515
            Width           =   1005
         End
         Begin VB.Label LblDep 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            DataField       =   "Nama Group"
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
            Left            =   1800
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   1470
            Width           =   3105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   20
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   19
            Top             =   510
            Width           =   645
         End
         Begin VB.Label lblFixAssets 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
            DataField       =   "No FA"
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
            Left            =   1800
            TabIndex        =   0
            Tag             =   "ASM"
            Top             =   150
            Width           =   780
         End
         Begin VB.Label lblFixAssets 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No Bukti"
            DataField       =   "Kode Customer"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   2
            Tag             =   "ASM"
            Top             =   780
            Width           =   3105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Customer"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   6
            Left            =   300
            TabIndex        =   18
            Top             =   825
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   7
            Left            =   300
            TabIndex        =   17
            Top             =   1185
            Width           =   1290
         End
         Begin VB.Label lblFixAssets 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No Bukti"
            DataField       =   "Nama Perusahaan"
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
            Index           =   2
            Left            =   1800
            TabIndex        =   4
            Tag             =   "ASM"
            Top             =   1125
            Width           =   3510
         End
         Begin VB.Label lblAlamatBank 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Nama Kas"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5520
            TabIndex        =   9
            Tag             =   "ASM"
            Top             =   1035
            Width           =   4020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kas/Bank"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   5520
            TabIndex        =   16
            Top             =   465
            Width           =   735
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
            Left            =   5520
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   705
            Width           =   3585
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   270
            X2              =   1950
            Y1              =   750
            Y2              =   750
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   6240
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
      LimitRecordData =   "1"
   End
End
Attribute VB_Name = "FrmPenjualanFixAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData As New clsTransaksi
Private mVarAdd As Boolean
Private TotalAkum As Variant
Private AkumAccount, mVarAccDep, mVarDepre As String
Private RcPartner As New DBQuick

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mVarAdd = True Then
   Select Case DGPurchase.Col
          Case 4, 6: DGPurchase.AllowUpdate = True
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
'
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
Set mCall = New frmCaller
DTPicker1.Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmPenjualanFixAssets
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     [TR Aktiva Tetap].[No FA], [TR Aktiva Tetap].DateTrans, [TR Aktiva Tetap].PartnerID AS [Kode Customer],                        PartnerDB.CompanyName AS [Nama Perusahaan], [TR Aktiva Tetap].[Id Group], [TR Aktiva Tetap].BankID AS [Kode Kas],                        GLAccount.AccountName AS [Nama Kas], [TR Aktiva Tetap].Periode, [TR Aktiva Tetap].DP, [TR Aktiva Tetap].NoCheque,                        GLAccount_1.AccountName AS [Nama Group] FROM         [TR Aktiva Tetap] INNER JOIN                       PartnerDB ON [TR Aktiva Tetap].PartnerID = PartnerDB.PartnerID INNER JOIN                       GLAccount ON [TR Aktiva Tetap].BankID = GLAccount.NoAccount INNER JOIN                       GLAccount GLAccount_1 ON [TR Aktiva Tetap].[Id Group] = GLAccount_1.NoAccount WHERE     ([TR Aktiva Tetap].TypeTrans = N'FJ') AND ([TR Aktiva Tetap].Periode = " & mVarPeriode & ") ORDER BY [TR Aktiva Tetap].[No FA]"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MyDDE.ClearRecordset
Set MyData = Nothing
Set mCall = Nothing
End Sub

Private Sub Form_Resize()


Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPenjualanFixAssets = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case TagForm
       Case "MASTER CUSTOMER":
            With MyDDE.ActiveRecordset
                 .Fields("Kode Customer") = mCall.GetFieldByName(0)
                 .Fields("Nama Perusahaan") = mCall.GetFieldByName(1)
            End With
       Case "MASTER KAS":
            With MyDDE
                 .GetFieldByName("Kode Kas") = mCall.GetFieldByName(0)
                 .GetFieldByName("Nama Kas") = mCall.GetFieldByName(1)
            End With
            'lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
       Case "RUGI PENJUALAN AKTIVA":
            With MyDDE.ActiveRecordset
                 .Fields("AccDep") = mCall.GetFieldByName(0)
            End With
       Case "LABA PENJUALAN AKTIVA":
            With MyDDE.ActiveRecordset
                 .Fields("DepAktiva") = mCall.GetFieldByName(0)
            End With
       Case "KELOMPOK AKTIVA":
            With MyDDE.ActiveRecordset
                 .Fields("ID Group") = mCall.GetFieldByName(0)
                 .Fields("Nama Group") = mCall.GetFieldByName(1)
            End With
       Case "TRANSAKSI AKTIVA":
            With MyDDE.ChildRecordset
                 .Fields(0) = mCall.GetFieldByName(1)
                 .Fields(1) = mCall.GetFieldByName(3)
                 .Fields(2) = mCall.GetFieldByName(4)
                 .Fields(3) = mCall.GetFieldByName(5)
                 .Fields(4) = mCall.GetFieldByName(6)
                 .Fields(5) = 0
                 .Fields("Ppn") = 0
                 .Fields("NoAccount") = mCall.GetFieldByName("No Akun")
            End With
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
            'txtBox(0).Enabled = False
       Case tmbAddNew:
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            With MyDDE
                 .GetFieldByName("No FA") = MyData.PrepareIndex(tmbTransaksiJualAktivaTetap, 5, "", TglIndex)
                 .GetFieldByName("Umur") = 0
                 .GetFieldByName("DP") = 0
                 .GetFieldByName("NoCheque") = "-"
                 .GetFieldByName("Tanggal") = DTPicker1.Value
            End With
            'txtBox(0).SetFocus
            mVarAdd = True
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               If OpenPartner(5) = True Then CancelDetailTrans
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.IsChildMemberReady = True Then
               'SendDataToServer ("DELETE FROM Inventory WHERE     (NoItem = N'" & 'txtBox(0) & "') ")
            End If
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
            End If
       Case tmbPrint:
            CallRPTReport "Penjualan Aktiva.Rpt", "Select * from [Penjualan Aktiva] Where [No Bukti]='" & lblFixAssets(0) & "'"
'       Case Else: 'mVarDataDc = False
End Select
cmdLInk(0).Enabled = mVarAdd
cmdLInk(1).Enabled = mVarAdd
cmdLInk(2).Enabled = mVarAdd
'cmdLInk(3).Enabled = mVarAdd
'cmdLInk(4).Enabled = mVarAdd
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail IIf(Not IsNull(MyDDE.GetFieldByName("No FA")), MyDDE.GetFieldByName("No FA"), "XXXXX")
''lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  MyDDE.IsChildMemberReady = True
               Else
                  MyDDE.IsChildMemberReady = False
                  MessageBox "Data detail belum ada. Harap diisi dulu.", "Peringatan", msgOkOnly
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
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then mVarAdd = False
'       Case tmbDetail:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               OpenPartner
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
End Select
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
    .PrepareAppend = " INSERT INTO [TR Aktiva Tetap]" & _
                     " ( [No FA], DateTrans, PartnerID, [Id Group], BankID, Periode,Typetrans,Disposal,Dp,NoCheque)" & _
                     " VALUES  (N'" & lblFixAssets(0) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & lblFixAssets(1) & "',N'" & MyDDE.GetFieldByName("Id Group") & "',N'" & NoVoucher(1) & "'," & mVarPeriode & ",'FJ',1," & CDbl(txtBox(0)) & ",'" & txtBox(1) & "')"
                     
    .PrepareUpdate = " UPDATE [TR Aktiva Tetap]" & _
                     " SET Dp=" & CDbl(txtBox(0)) & ",NoCheque= '" & txtBox(1) & "',DateTrans = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), PartnerID = N'" & lblFixAssets(1) & "'," & _
                     " [ID Group] =N'" & MyDDE.GetFieldByName("Id Group") & "' ,BankID=N'" & NoVoucher(1) & "'" & _
                     " WHERE ([No FA] = N'" & lblFixAssets(0) & "')"

    .PrepareDelete = " DELETE FROM [TR Aktiva Tetap] WHERE     ([No FA] = N'" & lblFixAssets(0) & "') "
End With
Err.Clear
End Sub

Private Sub SimpanDetail()
Dim mTot As Variant
Dim mVarNilaiBuku As Variant
Dim mVarPiutang As Variant
Dim MyKodeKu As String
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
        If SendDataToServer("DELETE FROM [DTR Aktiva Tetap] WHERE     ([No FA] = N'" & lblFixAssets(0) & "')") = True Then
           .MoveFirst
           Do
             mTot = 0
             MyKodeKu = MyData.PrepareIndex(tmbTransaksiBKMAT, 5, "", TglIndex2)
             If .EOF = True Then Exit Do
             If SendDataToServer(" INSERT INTO [DTR Aktiva Tetap]  ([No FA],[Doc Reff], [No Aktiva],[Aktiva Jual], Harga,[Harga Jual],ppn,NoAccount) " & _
                                 " VALUES (N'" & lblFixAssets(0) & "',N'" & .Fields("Doc Reff") & "', N'" & .Fields("No Aktiva") & "', " & .Fields("Aktiva Jual") & ", " & .Fields("Harga") & "," & .Fields("Harga Jual") & "," & .Fields("PPn") & ",'" & .Fields("NoAccount") & "')") = True Then
'                If SendDataToServer(" INSERT INTO [Table Journal] (JournalID, TransID,  PartnerID, DateTrans, Periode, TypeTrans,NoUrut) VALUES (N'" & MyKodeKu & "', N'" & lblFixAssets(0) & "', N'" & lblFixAssets(1) & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BKMAT','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "/") & "')") = True Then
'                    If txtBox(0) <> 0 Then
'                       SendDataToServer " INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES(N'" & MyKodeKu & "', N'" & NoVoucher(1) & "', N'" & txtBox(1) & "', " & CDbl(txtBox(0)) & ", 0, N'Penerimaan Penjualan Aktiva " & Left(.Fields("Nama Aktiva"), 112) & "')"
'                    End If
'                    mVarPiutang = CDbl((.Fields("Harga Jual") * (.Fields("PPn") / 100)) + .Fields("Harga Jual"))
'                    Select Case mVarPiutang
'                           Case Is = CDbl(txtBox(0)):
'                           Case Is < CDbl(txtBox(0)): SendDataToServer " INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES(N'" & MyKodeKu & "', N'" & CariTypeAccount(56) & "', N'" & txtBox(1) & "', " & CDbl((.Fields("Harga Jual") * (.Fields("PPn") / 100)) + .Fields("Harga Jual")) - CDbl(txtBox(0)) & ", 0, N'Penerimaan Penjualan Aktiva " & Left(.Fields("Nama Aktiva"), 112) & "')"
'                    End Select
'                    CariAkumulasi .Fields("No Aktiva"), .Fields("Doc Reff")
'                    CariNoAccountDepre .Fields("Doc Reff"), .Fields("No Aktiva")
                    
'                    SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & AkumAccount & "', N'" & .Fields("No Aktiva") & "', " & TotalAkum & ", 0, N'Penerimaan Penjualan Aktiva')")
'                    mVarNilaiBuku = .Fields("Harga") - TotalAkum
'                    mTot = .Fields("Harga Jual") - mVarNilaiBuku
'                    If mTot < 0 Then mTot = mTot * (-1)
'                    Select Case mVarNilaiBuku
'                           Case Is = .Fields("Harga Jual"): 'SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & CariTypeAccount("Rugi Penjualan/Penukaran Aktiva") & "', N'" & .Fields("No Aktiva") & "', 0, 0, N'Rugi Penjualan/Penukaran Aktiva" & Left(.Fields("Nama Aktiva"), 200) & "')")
'                           Case Is > .Fields("Harga Jual"): SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & CariTypeAccount(55) & "', N'" & .Fields("No Aktiva") & "', " & mTot & ", 0, N'Rugi Penjualan/Penukaran Aktiva" & Left(.Fields("Nama Aktiva"), 200) & "')")
'                           Case Is < .Fields("Harga Jual"): SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & CariTypeAccount(54) & "', N'" & .Fields("No Aktiva") & "', 0, " & mTot & ", N'Laba Penjualan/Penukaran Aktiva" & Left(.Fields("Nama Aktiva"), 200) & "')")
'                    End Select
'                    SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & .Fields("NoAccount") & "', N'" & .Fields("No Aktiva") & "', 0, " & .Fields("Harga") & ", N'Penerimaan Penjualan Aktiva')")
'                    SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & CariTypeAccount(41) & "', N'xxx', 0, " & .Fields("Harga Jual") * (.Fields("PPn") / 100) & ", N'Penerimaan Penjualan Aktiva')")
'                    SendDataToServer (" INSERT INTO [Detail Journal] (JournalID, NoAccount, [Doc Reff], Debet, Credit, Keterangan) VALUES (N'" & MyKodeKu & "', N'" & CariTypeAccount(56) & "', N'" & lblFixAssets(1) & "'," & mVarPiutang - CDbl(txtBox(0)) & ", 0, N'Piutang Penjualan Aktiva')")
                    SendDataToServer (" UPDATE    [TR Aktiva Tetap] SET              Disposal = 1 WHERE     ([No FA] = N'" & .Fields("Doc Reff") & "') ")
'                End If
             End If
             .MoveNext
           Loop
           .MoveLast
        End If
     End If
End With
End Sub

Private Sub OpenDetail(ByVal ParamString As String)
Dim RcDetail As New DBQuick
RcDetail.DBOpen " SELECT [DTR Aktiva Tetap].[Doc Reff], [DTR Aktiva Tetap].[No Aktiva], [Tabel Aktiva Tetap].[Nama Aktiva], [DTR Aktiva Tetap].[Aktiva Jual],                        [DTR Aktiva Tetap].Harga, [DTR Aktiva Tetap].[Harga Jual],[DTR Aktiva Tetap].[NoAccount],[DTR Aktiva Tetap].[Ppn] FROM         [DTR Aktiva Tetap] INNER JOIN                       [Tabel Aktiva Tetap] ON [DTR Aktiva Tetap].[No Aktiva] = [Tabel Aktiva Tetap].[No Aktiva] WHERE     ([DTR Aktiva Tetap].[No FA] = N'" & ParamString & "') ORDER BY [DTR Aktiva Tetap].[No Aktiva]", CNN
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
Set DGPurchase.DataSource = MyDDE.ChildRecordset
RcDetail.CloseDB
End Sub

Private Function OpenPartner(ByVal Index As Integer) As Boolean
Select Case Index
       Case 0:
            RcPartner.DBOpen "SELECT     PartnerID AS [Kode Supplier], CompanyName AS [Nama Perusahaan], Address AS Alamat, City AS Kota, PostalCode AS [Kode POS],                        Phone AS Telp FROM         PartnerDB WHERE     (PartnerType = N'CUSTOMER')", CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen " SELECT     NoAccount AS [Kode Kas], AccountName AS [Nama Kas] FROM         GLAccount WHERE     (Type = N'Kas' OR                      Type = N'Setara Kas' OR                      Type = N'Bank') AND ([Group] = N'Detail List Account') GROUP BY NoAccount, AccountName", CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen " SELECT     NoAccount AS [No Akun], AccountName AS [Nama Akun] FROM         GLAccount WHERE     ([Group] = N'Detail List Account') AND (Type = N'Aktiva Tetap Kantor' OR                      Type = N'Aktiva Tetap Produksi' OR                      Type = N'Aktiva Tetap Tak Berwujud') ORDER BY NoAccount", CNN, lckLockReadOnly
       Case 3:
            RcPartner.DBOpen " SELECT     NoAccount, AccountName FROM         GLAccount WHERE     ([Group] = N'List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
       Case 4:
            RcPartner.DBOpen " SELECT     NoAccount, AccountName FROM         GLAccount WHERE     ([Group] = N'List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
       Case 5:
            RcPartner.DBOpen "SELECT     [TR Aktiva Tetap].DateTrans AS [Tanggal Bukti], [TR Aktiva Tetap].[No FA] AS [No Bukti], [Tabel Aktiva Tetap].NoAccount AS [No Akun],                        [DTR Aktiva Tetap].[No Aktiva], [Tabel Aktiva Tetap].[Nama Aktiva], [DTR Aktiva Tetap].[Aktiva Beli] AS [Total Aktiva], [DTR Aktiva Tetap].Harga FROM         [TR Aktiva Tetap] INNER JOIN                       [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] INNER JOIN                       [Tabel Aktiva Tetap] ON [DTR Aktiva Tetap].[No Aktiva] = [Tabel Aktiva Tetap].[No Aktiva] WHERE     ([TR Aktiva Tetap].Disposal = 0) AND ([TR Aktiva Tetap].[Id Group] = N'" & MyDDE.GetFieldByName("Id Group") & "') GROUP BY [DTR Aktiva Tetap].[No Aktiva], [Tabel Aktiva Tetap].[Nama Aktiva], [DTR Aktiva Tetap].Harga, [TR Aktiva Tetap].[No FA], [TR Aktiva Tetap].DateTrans, [DTR Aktiva Tetap].[Aktiva Beli], [Tabel Aktiva Tetap].NoAccount ORDER BY [TR Aktiva Tetap].[No FA]", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
    Select Case Index
           Case 0: mCall.FromTagActive = "MASTER CUSTOMER"
           Case 1:
                mCall.FromTagActive = "MASTER KAS"
                mCall.txtCari = NoVoucher(1)
           Case 2:
                mCall.FromTagActive = "KELOMPOK AKTIVA"
                mCall.txtCari = NoVoucher(1)
           Case 3:
                mCall.FromTagActive = "RUGI PENJUALAN AKTIVA"
                mCall.txtCari = NoVoucher(1)
           Case 4:
                mCall.FromTagActive = "LABA PENJUALAN AKTIVA"
                mCall.txtCari = NoVoucher(1)
           Case 5:
                mCall.FromTagActive = "TRANSAKSI AKTIVA"
    End Select
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
    If FindOwnRecordset(MyDDE.ChildRecordset, "[No Aktiva] = '" & mCall.GetFieldByName("No Aktiva") & "'") = True Then
       MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("No Aktiva") & " Sudah Ada....! Silahkan Diulangi", "Peringatan"
       CancelDetailTrans
       DGPurchase.SetFocus
    End If
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
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = "FJ/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Function TglIndex2() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex2 = "JA/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub CariAkumulasi(ByVal KodeAktiva As String, ByVal NoFa As String)
Dim RcAkum As New DBQuick
'MsgBox "SELECT     SUM([Detail Journal].Debet) AS Akumulasi, [Detail Journal].NoAccount FROM         [Table Journal] INNER JOIN                       [TR Aktiva Tetap] ON [Table Journal].TransID = [TR Aktiva Tetap].[No FA] INNER JOIN                       [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([TR Aktiva Tetap].Disposal = 0) AND ([Table Journal].TypeTrans = N'AKDEP') AND ([TR Aktiva Tetap].[No FA] = N'" & NoFa & "') AND                        ([Detail Journal].[Doc Reff] = N'" & KodeAktiva & "') GROUP BY [Detail Journal].NoAccount HAVING      (SUM([Detail Journal].Debet) <> 0)"
RcAkum.DBOpen "SELECT     SUM([Detail Journal].Debet) AS Akumulasi, [Detail Journal].NoAccount FROM         [Table Journal] INNER JOIN                       [TR Aktiva Tetap] ON [Table Journal].TransID = [TR Aktiva Tetap].[No FA] INNER JOIN                       [Detail Journal] ON [Table Journal].JournalID = [Detail Journal].JournalID INNER JOIN                       GLAccount ON [Detail Journal].NoAccount = GLAccount.NoAccount WHERE     ([TR Aktiva Tetap].Disposal = 0) AND ([Table Journal].TypeTrans = N'AKDEP') AND ([TR Aktiva Tetap].[No FA] = N'" & NoFa & "') AND                        ([Detail Journal].[Doc Reff] = N'" & KodeAktiva & "') GROUP BY [Detail Journal].NoAccount HAVING      (SUM([Detail Journal].Debet) <> 0)", CNN, lckLockReadOnly
TotalAkum = 0
AkumAccount = ""
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        TotalAkum = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
        AkumAccount = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
     Else
        CariNoAccountDepre NoFa, KodeAktiva
        TotalAkum = 0
        AkumAccount = mVarAccDep
     End If
End With
End Sub

'Private Function CariNoAccount(ByVal Params As String) As String
'Dim RcAkum As New DBQuick
'RcAkum.DBOpen "SELECT     NoAccount, AccountName FROM         GLAccount WHERE     (Type = N'" & Params & "')", Cnn, lckLockReadOnly
'With RcAkum.DBRecordset
'     If .Recordcount <> 0 Then
'        CariNoAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
'     End If
'End With
'End Function

Private Sub CariNoAccountDepre(ByVal NoTranBeli As String, ByVal NoFixAssets As String)
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     [TR Aktiva Tetap].AccDep, [TR Aktiva Tetap].DepAktiva FROM         [TR Aktiva Tetap] INNER JOIN                       [DTR Aktiva Tetap] ON [TR Aktiva Tetap].[No FA] = [DTR Aktiva Tetap].[No FA] WHERE     ([TR Aktiva Tetap].[No FA] = N'" & NoTranBeli & "') AND ([DTR Aktiva Tetap].[No Aktiva] = N'" & NoFixAssets & "')", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        mVarAccDep = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
        mVarDepre = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
     End If
End With
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GLAccount.NoAccount, AccType.ID, GLAccount.AccountName FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function
