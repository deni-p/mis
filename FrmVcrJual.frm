VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmVcrJual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher Penjualan"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   Icon            =   "FrmVcrJual.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   0
      ScaleHeight     =   6180
      ScaleWidth      =   10815
      TabIndex        =   17
      Top             =   0
      Width           =   10815
      Begin MSComCtl2.DTPicker dtDate 
         DataField       =   "DateTrans"
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   3
         Tag             =   "Voucher"
         Top             =   490
         Width           =   2400
         _ExtentX        =   4233
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
         Format          =   65863683
         CurrentDate     =   38139
      End
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5145
         Picture         =   "FrmVcrJual.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   878
         Width           =   330
      End
      Begin VB.CommandButton CmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   10305
         Picture         =   "FrmVcrJual.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   878
         Width           =   330
      End
      Begin VB.ComboBox cboVoucher 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmVcrJual.frx":6F66
         Left            =   10860
         List            =   "FrmVcrJual.frx":6F6D
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         DataField       =   "RefNotes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   15
         Tag             =   "Voucher"
         Top             =   5745
         Width           =   4380
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2550
         Left            =   195
         TabIndex        =   13
         Tag             =   "Partner"
         Top             =   2700
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   4498
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
            DataField       =   "TransID"
            Caption         =   "No. Bukti"
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
            DataField       =   "DateTrans"
            Caption         =   "Tgl. Bukti"
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
         BeginProperty Column02 
            DataField       =   "RefNotes"
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
            DataField       =   "Debet"
            Caption         =   "Total Piutang"
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
            DataField       =   "Credit"
            Caption         =   "Total Bayar"
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
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotalKas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7125
         TabIndex        =   12
         Top             =   1657
         Width           =   3510
      End
      Begin VB.Label NoVoucher 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "PartnerID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1455
         TabIndex        =   6
         Top             =   877
         Width           =   3690
      End
      Begin VB.Label NoVoucher 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TransID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1455
         TabIndex        =   1
         Tag             =   "Voucher"
         Top             =   120
         Width           =   4020
      End
      Begin VB.Label NoVoucher 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Kode Kas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   7125
         TabIndex        =   10
         Tag             =   "Voucher"
         Top             =   870
         Width           =   3180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Kas/Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   5925
         TabIndex        =   8
         Top             =   945
         Width           =   1065
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   7440
         X2              =   5865
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   7440
         X2              =   5865
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1695
         X2              =   120
         Y1              =   1192
         Y2              =   1192
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1695
         X2              =   120
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1695
         X2              =   120
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Transaksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   0
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   7485
         TabIndex        =   24
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label lblABayar 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   7485
         TabIndex        =   23
         Top             =   5805
         Width           =   2580
      End
      Begin VB.Label lblAlamatBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Nama Kas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7125
         TabIndex        =   11
         Tag             =   "Voucher"
         Top             =   1245
         Width           =   3510
      End
      Begin VB.Label lblAlamat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   1455
         TabIndex        =   7
         Top             =   1245
         Width           =   4020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   4
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   5520
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Left            =   4815
         TabIndex        =   22
         Top             =   5520
         Width           =   435
      End
      Begin VB.Label LblAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   4815
         TabIndex        =   21
         Top             =   5805
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Transaksi"
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
         Left            =   10845
         TabIndex        =   20
         Top             =   2010
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Transaksi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   555
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   5925
         TabIndex        =   19
         Top             =   1695
         Width           =   390
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   6180
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmVcrJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents RcDetail As Recordset
Attribute RcDetail.VB_VarHelpID = -1
Private RcPartner As New DBQuick
Private MyData As New clsTransaksi
Private MEdit As Boolean
Private mVarTipe, mVarDetail, mVarKodeKas As String
Private mDebet, mCredit As Currency
Private mVarTempBayar As Variant
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim rsTrans As New ADODB.Recordset
Dim IDGen As New IDGenerator

Private Sub Form_Activate()
'If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'cboVoucher.ListIndex = 0
mVarTipe = "BR"
mVarDetail = "AR"
      
dtDate(0).Value = Date
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmVcrJual
    .BindFormTAG = "Voucher"
    .SetPermissions = UserEditDeleteDenied
    Set .ActiveConnection = CNN
    
     .PrepareQuery = " SELECT TransData.TransID, TransData.RefNotes, TransData.DateTrans, " & _
                    " TransData.bankid AS [Kode Kas], TransData.TypeTrans, TransData.PartnerId,  " & _
                    " TransData.EmpID, GLAccount.AccountName AS [Nama Kas] FROM TransData INNER JOIN " & _
                    " GLAccount ON TransData.bankid = GLAccount.NoAccount WHERE (TransData.TypeTrans = N'AR') or (TransData.TypeTrans = N'BR')" & _
                    " ORDER BY TransData.TransID"
                    
'DGPurchase.Columns(3).Caption = "Total Piutang"
End With
GridLayout
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set mCall = Nothing
Set MyData = Nothing
CloseDB RcDetail
RcPartner.CloseDB
MyDDE.ClearRecordset
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmVcrJual = Nothing
End Sub
'==============================================================================
'Private Sub cboVoucher_Click()
'Select Case cboVoucher
       'Case "Pelunasan Hutang":
            'mVarTipe = "BP"
            'mVarDetail = "AP"
       '     DGPurchase.Columns(3).Caption = "Total Hutang"
       '     Label1(5).Caption = "Supplier"
      ' Case "Pelunasan Piutang":
      '      mVarTipe = "BR"
      '      mVarDetail = "AR"
      '      DGPurchase.Columns(3).Caption = "Total Piutang"
      '      Label1(5).Caption = "Customer"
'       Case "Petty Cash":
'            mVarTipe = "PP"
'            DGPurchase.Columns(3).Caption = "Total Petty Cash"
'       Case "Expenses":
'            mVarTipe = "BE"
'            DGPurchase.Columns(3).Caption = "Total Expenses"
'       Case "Piutang Karyawan"
'            mVarTipe = "BK"
'            DGPurchase.Columns(3).Caption = "Piutang Karyawan"
'       Case "Pembayaran Piutang Karyawan"
'            mVarTipe = "BU"
'            DGPurchase.Columns(3).Caption = "Bayar Piutang Karyawan"
'End Select
'Label3 = DGPurchase.Columns(3).Caption
'If mEdit = True Then
'   MyDDE.GetFieldByName("TransID") = MyData.PrepareIndex(tmbVoucher, 5, mVarTipe, TglIndex)
   'MyDDE.GetFieldByName("TransID") = NoVoucher(0)
'End If
'End Sub
'==================================================================================

Private Sub cmdLink_Click(Index As Integer)
OpenPartner Index
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim ISta As Byte
Select Case ColIndex
       Case 4:
            If DGPurchase.Columns(4) = "" Then DGPurchase.Columns(4) = 0
            If DGPurchase.Columns(3).Value = "0" Then DGPurchase.Columns(3).Value = CekBayar(NoVoucher(2), RcDetail.Fields("TransID"))
            If CDbl(DGPurchase.Columns(4)) > CDbl(DGPurchase.Columns(3)) Then
               MessageBox "Nilai Terlalu Besar Dengan Hutang Yang Ada.", "Peringatan", msgOkOnly
               DGPurchase.Columns(4) = 0
'               DGPurchase.Columns(3).Value = CekBayar(NoVoucher(2), RcDetail.Fields("TransID"))
'               DGPurchase.Columns(3).Value = DGPurchase.Columns(3) - DGPurchase.Columns(4)
            Else
                DGPurchase.Columns(3).Value = CekBayar(NoVoucher(2), RcDetail.Fields("TransID"))
                DGPurchase.Columns(3).Value = DGPurchase.Columns(3) - DGPurchase.Columns(4)
                If DGPurchase.Columns(3).Value = 0 Then
                   RcDetail.Fields("Status") = True
                   ISta = 1
                Else
                   RcDetail.Fields("Status") = False
                   ISta = 0
                End If
            End If
            '
            HitungTotal True
End Select
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.MarqueeStyle = dbgHighlightRowRaiseCell
   Exit Sub
End If
Select Case DGPurchase.col
       Case 0, 1, 2, 3:
            DGPurchase.MarqueeStyle = dbgHighlightCell
            DGPurchase.AllowUpdate = False
       Case 4:
            
            DGPurchase.MarqueeStyle = dbgFloatingEditor
            DGPurchase.AllowUpdate = True
End Select
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Select Case UCase(TagForm)
       Case "MASTER KAS":
             MyDDE.GetFieldByName("Kode Kas") = mCall.GetFieldByName(0)
             lblAlamatBank = mCall.GetFieldByName(1)   '& vbcrlf & = mCall.GetFieldByName(0)
             lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
       Case "MASTER PARTNER":
             lblAlamat = mCall.GetFieldByName(1) & vbCrLf & mCall.GetFieldByName(2) & vbCrLf & mCall.GetFieldByName(3) ' & mCall.GetFieldByName(4)
             NoVoucher(2) = mCall.GetFieldByName(0)
             MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
             OpenDetail IIf(Not IsNull(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), "xxxx") '
             'TotalTrans IIf(Not IsNull(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), "xxxx") '
             HitungTotal
             mVarTempBayar = CCur(LblAmount)
      Case "MASTER KARYAWAN":
             lblABayar = 0
             lblAlamat = mCall.GetFieldByName(1) '& mCall.GetFieldByName(2) & mCall.GetFieldByName(3) & mCall.GetFieldByName(4)
             NoVoucher(2) = mCall.GetFieldByName(0)
             MyDDE.GetFieldByName("EmpID") = mCall.GetFieldByName(0)
             OpenDetail IIf(Not IsNull(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), "xxxx") '
             'TotalTrans IIf(Not IsNull(mCall.GetFieldByName(0)), mCall.GetFieldByName(0), "xxxx") '
             HitungTotal
             mVarTempBayar = CCur(LblAmount)
             lblABayar = 0
End Select
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell
Set mCall = New frmCaller
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT GLAccount.NoAccount As [Kode Kas], GLAccount.AccountName as [Nama Kas] FROM         GLAccount LEFT OUTER JOIN                      [Tabel Pembantu] ON GLAccount.NoAccount = [Tabel Pembantu].NoAccount WHERE     (GLAccount.Type = N'Kas' OR                      GLAccount.Type = N'setara kas' OR                      GLAccount.Type = N'BANK') AND (GLAccount.[Group] = N'Detail List Account') GROUP BY GLAccount.NoAccount, GLAccount.AccountName", CNN, lckLockReadOnly
       Case 1:
            '===============================Edit =============================
            'If cboVoucher = "Pelunasan Hutang" Then
            '   RcPartner.DBOpen "SELECT     [Voucher Batch].PartnerID AS [Kode Supplier], PartnerDB.CompanyName AS Nama, PartnerDB.Address AS Alamat, PartnerDB.City AS Kota,                        PartnerDB.Phone AS Telepon FROM         [Voucher Batch] INNER JOIN                       PartnerDB ON [Voucher Batch].PartnerID = PartnerDB.PartnerID GROUP BY [Voucher Batch].PartnerID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, PartnerDB.Phone, [Voucher Batch].TypeTrans HAVING      ([Voucher Batch].TypeTrans = N'AP') OR                       ([Voucher Batch].TypeTrans = N'RB') ORDER BY [Voucher Batch].PartnerID, PartnerDB.CompanyName", CNN, lckLockReadOnly
            'ElseIf cboVoucher = "Pelunasan Piutang" Then
               RcPartner.DBOpen "SELECT     [Voucher Batch].PartnerID AS [Kode Customer], PartnerDB.CompanyName AS Nama, PartnerDB.Address AS Alamat, PartnerDB.City AS Kota,                       PartnerDB.Phone AS Telepon FROM         [Voucher Batch] INNER JOIN                       PartnerDB ON [Voucher Batch].PartnerID = PartnerDB.PartnerID GROUP BY [Voucher Batch].PartnerID, PartnerDB.CompanyName, PartnerDB.Address, PartnerDB.City, PartnerDB.Phone, [Voucher Batch].TypeTrans HAVING      ([Voucher Batch].TypeTrans = N'AR') or ([Voucher Batch].TypeTrans = N'BR') ORDER BY [Voucher Batch].PartnerID", CNN, lckLockReadOnly
            'End If
            '================================================================
End Select
If RcPartner.Recordcount <> 0 Then
    If Index = 0 Then
       mCall.FromTagActive = "Master Kas"
       mCall.txtCari = NoVoucher(1)
    Else
       '========================================= Edit =======================
       'If cboVoucher = "Pelunasan Hutang" Then
       '   mCall.FromTagActive = "Master Partner"
       'ElseIf cboVoucher = "Pelunasan Piutang" Then
          mCall.FromTagActive = "Master Partner"
       'End If
       '======================================================================
       OpenDetail NoVoucher(2)
       mCall.txtCari = NoVoucher(2)
    End If
    Set mCall.FormData = RcPartner.DBRecordset
    mCall.LookUp Me
    'cboVoucher.Enabled = False
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
End If

Exit Sub
Hell:
    'MsgBox Err.Description
    Err.Clear
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If RcDetail.Recordcount <> 0 Then
                  If NoVoucher(1) <> "" Then
                    If CekGrid = True Then
                       Select Case Left(MyDDE.GetFieldByName("TransID"), 2)
                              Case "BP", "PP":
                                   'HitungTotal True
                                   If CCur(lblABayar) > CDbl(lblTotalKas) Then
                                      MessageBox "Total Saldo Kas Belum Mencukupi Untuk Melakukan Transaksi Pembayaran", "Transaksi Pembayaran", msgOkOnly
                                      MyDDE.IsChildMemberReady = False
                                      Exit Sub
                                   End If
                       End Select
                       MyDDE.IsChildMemberReady = True
                       MyDDE.GetFieldByName("DateTrans") = dtDate(0).Value
                       
                    Else
                       MessageBox "Belum Ada Nilai Untuk Pembayaran/Penerimaan.", "Peringatan", msgOkOnly
                       MyDDE.IsChildMemberReady = False
                    End If
                  Else
                     MessageBox "Data bank atau kas belum dipilih.", "Peringatan", msgOkOnly
                     MyDDE.IsChildMemberReady = False
                  End If
               Else
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
NoVoucher(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            MEdit = True
            DGPurchase.Columns(3).Visible = True
       Case tmbAddNew:
            MEdit = True
            dtDate(0).Value = CDate(Format(Date, "dd/mm/yyyy"))
            mVarTipe = "BR"
            mVarDetail = "AR"
            MyDDE.GetFieldByName("DateTrans") = dtDate(0).Value
            '================ Edit =========================================
            'Select Case cboVoucher
                   'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses", "Piutang Karyawan", "Pembayaran Piutang Karyawan": MyDDE.GetFieldByName("TransID") = MyData.PrepareIndex(tmbVoucher, 5, mVarTipe, TglIndex)
                   'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses", "Piutang Karyawan", "Pembayaran Piutang Karyawan":
                   'MyDDE.GetFieldByName("TransID") = MyData.PrepareIndex(tmbVoucher, 5, mVarTipe, TglIndex)
                    MyDDE.GetFieldByName("TransID") = IDGen.GetID("VS")
                   'Case "Piutang Karyawan": Mydde.GetFieldByName("TransID") = MyData.PrepareIndex(tmbTransaksiPiutangKaryawan, 5, mVarTipe, TglIndex)
            'End Select
            '===============================================================
            MyDDE.GetFieldByName("RefNotes") = "-"
            DGPurchase.Columns(3).Visible = True
            DGPurchase.Columns(4).Visible = True
            GridLayout
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               If PrepareByrVoucher = True Then
                  SimpanVoucher
                  RcDetail.Requery
                  MEdit = False
                  OpendDetailVoucher NoVoucher(0)
                  HitungTotal True
                  'HitungTotal T
                  MyDDE.RefreshDatabase
                  
               End If
            End If
       Case tmbCancel:
            MEdit = False
       Case tmbPrint:
            'Select Case Left(NoVoucher(0), 2)
            '       Case "BP", "PP", "BE":
            '           CallRPTReport "Bank Payment.rpt", "Select * from [Bank Payment] where TransID=N'" & NoVoucher(0) & "'"
            '       Case "BR"
                       CallRPTReport "Bank Receipt.rpt", "Select * from [Bank Receipt] where TransID=N'" & NoVoucher(0) & "'"
            '       Case "BK": CallRPTReport "BKK PKaryawan.rpt", "select * from [Bkk pkaryawan] where [no Piutang] =N'" & NoVoucher(0) & "'"
            'End Select
End Select
cmdLink(0).Enabled = MEdit
cmdLink(1).Enabled = MEdit
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpendDetailVoucher IIf(Not IsNull(MyDDE.GetFieldByName("TransID")), MyDDE.GetFieldByName("TransID"), "xxxx")

With MyDDE
     If MyDDE.ActiveRecordset.Recordcount <> 0 Then
        'lblAlamat = .GetFieldByName("Nama") & .GetFieldByName("alamat") & .GetFieldByName("Telepon") & .GetFieldByName("Mobile") & .GetFieldByName("Faximile")
        lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
     Else
        lblAlamat = ""
        lblTotalKas = 0
     End If
End With
'=======================================================Edit====================
'Select Case Left(MyDDE.GetFieldByName("TransID"), 2)
'       Case "BP":
'            cboVoucher.ListIndex = 0
'            NoVoucher(2) = IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "xxxx")
'       Case "BR":
            'cboVoucher.ListIndex = 1
            
NoVoucher(2) = IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "xxxx")

'       Case "PP":
'            cboVoucher.ListIndex = 2
'            NoVoucher(2) = IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "xxxx")
'       Case "BE":
'            cboVoucher.ListIndex = 3
'            NoVoucher(2) = IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "xxxx")
'       Case "BK":
'            cboVoucher.ListIndex = 2
'            NoVoucher(2) = IIf(Not IsNull(MyDDE.GetFieldByName("empID")), MyDDE.GetFieldByName("empID"), "xxxx")
'       Case "BU":
'            cboVoucher.ListIndex = 3
'            NoVoucher(2) = IIf(Not IsNull(MyDDE.GetFieldByName("empID")), MyDDE.GetFieldByName("empID"), "xxxx")
            
'End Select
'=========================================================================

OpenPartnerVoucher IIf(Not IsNull(MyDDE.GetFieldByName("PartnerID")), MyDDE.GetFieldByName("PartnerID"), "xxxx"), IIf(Not IsNull(MyDDE.GetFieldByName("empID")), MyDDE.GetFieldByName("empID"), "xxxx")
End Sub

Private Function TglIndex() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String
TglIndex = mVarTipe & "-" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
End Function

Private Sub OpenDetail(ByVal ParameterString As String)
Dim rs As New Recordset
CloseDB RcDetail
Set RcDetail = New Recordset
RcDetail.CursorLocation = adUseClient
If ParameterString = "" Then ParameterString = "xxxxxxxx"
With RcDetail
'=============================================== Edit =======================
     'Select Case cboVoucher
            'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses": .Open " SELECT TransID, DateTrans, RefNotes, Debet, Credit,Referense,Status FROM [Voucher Batch] WHERE (TypeTrans = N'" & mVarDetail & "') AND (PartnerID = N'" & ParameterString & "') ORDER BY DateTrans, TransID", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
            'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses":
            .Open " SELECT TransID, DateTrans, RefNotes, Debet, Credit,Referense,Status FROM [Voucher Batch] WHERE (TypeTrans = N'" & mVarDetail & "') AND (PartnerID = N'" & ParameterString & "') ORDER BY DateTrans, TransID", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
     '       Case "Piutang Karyawan": .Open " SELECT     [No Piutang] AS TransID, DateTrans, EmpID AS RefNotes, Jumlah AS Debet, Angsuran AS Credit, EmpID AS Referense, Status FROM         [BKK Karyawan] WHERE     (EmpID = N'" & ParameterString & "') and Typetrans =N'PPK'  ORDER BY [No Piutang]", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
     '       Case "Pembayaran Piutang Karyawan": .Open " SELECT     [No Piutang] AS TransID, DateTrans, EmpID AS RefNotes, Jumlah AS Debet, Angsuran AS Credit, EmpID AS Referense, Status FROM         [BKK Karyawan] WHERE     (EmpID = N'" & ParameterString & "') and Typetrans =N'BPK'  ORDER BY [No Piutang]", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
     'End Select
     Set DGPurchase.DataSource = RcDetail
' ==========================================================================
End With
End Sub


Private Sub SimpanVoucher()
Dim rctl As New Recordset
Dim Avdata As Variant
Dim Md As Currency
Dim Mc As Currency
Dim mykey As String
Dim I As Long
rctl.CursorLocation = adUseClient
'====================== Edit
'Select Case cboVoucher
       'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses": rctl.Open " SELECT Debet, Credit,TRANSID,PurchaseID FROM [Voucher Batch] where      (PartnerID = N'" & NoVoucher(2) & "')  ORDER BY DateTrans, TransID", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
       'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses":
       rctl.Open " SELECT Debet, Credit,TRANSID,PurchaseID FROM [Voucher Batch] where      (PartnerID = N'" & NoVoucher(2) & "')  ORDER BY DateTrans, TransID", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
'       Case "Piutang Karyawan", "Pembayaran Piutang Karyawan": rctl.Open " SELECT   Jumlah AS Debet, Angsuran AS Credit,  [No Piutang] AS TransID,  EmpID AS PurchaseID  FROM         [BKK Karyawan] WHERE (EmpID = N'" & NoVoucher(2) & "') and (TypeTrans = N'PPK') ORDER BY DateTrans,[No Piutang]", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
'End Select

Md = 0
Mc = 0
With rctl
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            'SendDataToServer ("INSERT INTO [Voucher Ref] (TransID, DNID) VALUES     (N'" & NoVoucher(0) & "', N'" & Avdata(2, i) & "')")
            
            Md = Md + Avdata(0, I)
            Mc = Mc + Avdata(1, I)
        Next I
     End If
End With
'CloseDB rctl

MyData.PreparePayVoucher NoVoucher(0), "-", mVarTempBayar, Mc, NoVoucher(2), NoVoucher(0)
'Select Case cboVoucher
'       Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Piutang Karyawan", "Expenses": MyData.PreparePayVoucher NoVoucher(0), "-", mVarTempBayar, Mc, NoVoucher(2), NoVoucher(0)
'       'Case "Piutang Karyawan": SendDataToServer (" INSERT INTO [BKK Karyawan]" & _
'                                                  " ([No Piutang], Kode Kas,DateTrans, EmpID, Jumlah, Angsuran, TypeTrans)" & _
'                                                  " VALUES (N'" & NoVoucher(0) & "',N'" & NoVoucher(1) & "', CONVERT(DATETIME, '" & Format(dtDate(0), "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', " & mVarTempBayar & ", " & Mc & ", N'PK')")
'End Select

'Select Case Left(MyDDE.GetFieldByName("TransID"), 2)
'       Case "BP", "PP", "PK": SendDataToServer ("UPDATE    [Temp Bank] SET             Amount =Amount - " & Mc & " WHERE     (Kode Kas = N'" & NoVoucher(1) & "')")
'       Case "BR", "BE", "BK", "BU": SendDataToServer ("UPDATE    [Temp Bank] SET             Amount =Amount + " & Mc & " WHERE     (Kode Kas = N'" & NoVoucher(1) & "')")
'End Select

lblTotalKas = FormatNumber(TotalKas(IIf(Not IsNull(MyDDE.GetFieldByName("Kode Kas")), MyDDE.GetFieldByName("Kode Kas"), "XXXXX")), 0)
Set rctl = RcDetail.Clone(adLockReadOnly)
'Select Case cboVoucher
'       Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses":
            With rctl
                 If .Recordcount <> 0 Then
                    Avdata = .Getrows(.Recordcount, adBookmarkFirst)
                    For I = 0 To UBound(Avdata, 2)
                        If Avdata(4, I) <> 0 Then
                         '  Debug.Print Avdata(0, i)
                'old      'SendDataToServer ("UPDATE TransData SET PurchaseID = N'" & PORef(Avdata(0, I)) & "' WHERE     (TransID = N'" & NoVoucher(0) & "')")
                           SendDataToServer ("UPDATE TransData SET PurchaseID = N'" & Avdata(0, I) & "' WHERE TransID = N'" & NoVoucher(0) & "'")
                           'Set rsTrans = CNN.Execute("UPDATE TransData SET PurchaseID = N'" & Avdata(0, i) & "' WHERE TransID = N'" & NoVoucher(0) & "'")
                           SendDataToServer ("INSERT INTO [Voucher Ref] (TransID, DNID) VALUES     (N'" & NoVoucher(0) & "', N'" & Avdata(0, I) & "')")
                           ClosedTransaksi Avdata(0, I)
                        End If
                    Next I
                 End If
            End With
Set Avdata = Nothing
CloseDB rctl
      'Case "Piutang Karyawan": SendDataToServer (" INSERT INTO [BKK Karyawan]" & _
                                                  " ([No Piutang], DateTrans, EmpID, Jumlah, Angsuran, TypeTrans)" & _
                                                  " VALUES     (N'" & NoVoucher(0) & "', CONVERT(DATETIME, '" & Format(dtDate(0), "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', " & mVarTempBayar & ", " & Mc & ", N'PK')")
'End Select
mykey = IdxAuto

'================================================Edit=====================================
'Select Case cboVoucher
       'Case "Pelunasan Hutang":
       '      If SendDataToServer(" INSERT INTO [Table Journal]" & _
                                 " (JournalID, TransID,  NoAccount, PartnerID, Currency, DateTrans,  Periode, TypeTrans,Nourut)" & _
                                 " VALUES     (N'" & mykey & "', N'" & NoVoucher(0) & "',  N'" & NoVoucher(1) & "',N'" & NoVoucher(2) & "',  N'IDR', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BKK','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "')") = True Then
                                 
       '         SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
                                  " VALUES   (N'" & mykey & "', N'" & CariTypeAccount(28) & "', N'" & NoVoucher(2) & "', " & CCur(lblABayar) & ", 0)")
                                  
       '         SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
                                  " VALUES   (N'" & mykey & "', N'" & NoVoucher(1) & "', N'" & NoVoucher(1) & "', 0, " & CCur(lblABayar) & ")")
                                  
            ' End If
  '     Case "Pelunasan Piutang":
             If SendDataToServer(" INSERT INTO [Table Journal]" & _
                                 " (JournalID, TransID,  NoAccount, PartnerID, Currency, DateTrans,  Periode, TypeTrans,Nourut)" & _
                                 " VALUES     (N'" & mykey & "', N'" & NoVoucher(0) & "',  N'" & NoVoucher(1) & "',N'" & NoVoucher(2) & "',  N'IDR', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BKM','" & MyData.PrepareIndex(tmbTransaksiNOJOURNAL, 13, Format(Year(dDateBegin), "yyyy"), "JR" & Format(Year(dDateBegin), "yyyy") & "-") & "')") = True Then
                                 
                SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
                                  " VALUES   (N'" & mykey & "', N'" & CariTypeAccount(39) & "', N'" & NoVoucher(2) & "', 0, " & CCur(lblABayar) & ")")
                                  
                SendDataToServer (" INSERT INTO [Detail Journal]" & _
                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
                                  " VALUES   (N'" & mykey & "', N'" & NoVoucher(1) & "', N'" & NoVoucher(1) & "', " & CCur(lblABayar) & ", 0)")
                                  
             End If
       
'       Case "Piutang Karyawan":
''             If SendDataToServer(" INSERT INTO [Table Journal]" & _
''                                 " (JournalID, TransID,  Kode Kas, PartnerID, Currency, DateTrans,  Periode, TypeTrans)" & _
''                                 " VALUES     (N'" & Mykey & "', N'" & NoVoucher(0) & "',  N'" & NoVoucher(1) & "',N'" & NoVoucher(2) & "',  N'IDR', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BKK')") = True Then
''
''                SendDataToServer (" INSERT INTO [Detail Journal]" & _
''                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
''                                  " VALUES   (N'" & Mykey & "', N'" & CariAkun("Piutang Karyawan", "Kode Karyawan") & "', N'" & NoVoucher(2) & "', " & CCur(lblABayar) & ", 0)")
''
''                SendDataToServer (" INSERT INTO [Detail Journal]" & _
''                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
''                                  " VALUES   (N'" & Mykey & "', N'" & CariAkun("Kas Keluar", "Kode Kas") & "', N'" & NoVoucher(1) & "', 0, " & CCur(lblABayar) & ")")
''
''             End If
'
'       Case "Pembayaran Piutang Karyawan":
'             If SendDataToServer(" INSERT INTO [Table Journal]" & _
'                                 " (JournalID, TransID,  Kode Kas, PartnerID, Currency, DateTrans,  Periode, TypeTrans)" & _
'                                 " VALUES     (N'" & Mykey & "', N'" & NoVoucher(0) & "',  N'" & NoVoucher(1) & "',N'" & NoVoucher(2) & "',  N'IDR', CONVERT(DATETIME, '" & Format(dtDate(0).Value, "dd/mm/yy") & "', 3), " & mVarPeriode & ", N'BKM')") = True Then
'
'                SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
'                                  " VALUES   (N'" & Mykey & "', N'" & CariAkun("Piutang Karyawan", "Kode Karyawan") & "', N'" & NoVoucher(2) & "', " & CCur(lblABayar) & ", 0)")
'
'                SendDataToServer (" INSERT INTO [Detail Journal]" & _
'                                  " (JournalID, NoAccount, [Doc Reff], Debet, Credit) " & _
'                                  " VALUES   (N'" & Mykey & "', N'" & CariAkun("Kas Masuk", "Kode Kas") & "', N'" & NoVoucher(1) & "', 0, " & CCur(lblABayar) & ")")
'
'             End If
'End Select
'=========================================================================================
End Sub

Private Sub HitungTotal(Optional ByVal Tipical As Boolean)
Dim RcTotal As New Recordset
Dim Avdata As Variant
Dim mTmp As Currency
Dim mTotal As Currency
Dim I As Long
RcTotal.CursorLocation = adUseClient
Set RcTotal = RcDetail.Clone(adLockReadOnly)
mTotal = 0
mTmp = 0
mDebet = 0
mCredit = 0
With RcTotal
     If Tipical = False Then
        If .Recordcount <> 0 Then
           Avdata = .Getrows(.Recordcount, adBookmarkFirst)
           For I = 0 To UBound(Avdata, 2)
               If CBool(Avdata(6, I)) = True Then
                  mDebet = mDebet + Avdata(3, I)
                  mCredit = mCredit + Avdata(4, I)
               End If
               
               mTotal = mTotal + ((Avdata(3, I) + Avdata(4, I)) - Avdata(4, I))
               
           Next I
        Else
           mTotal = 0
        End If
        LblAmount = FormatNumber(Abs(mTotal), 0)
     Else
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
           ' If CBool(AvData(6, I)) = True Then
               mDebet = mDebet + Avdata(3, I)
               mCredit = mCredit + Avdata(4, I)
          ' End If
            Select Case Left(Avdata(0, I), 2)
                   Case "BP": mTotal = mCredit 'mTotal + (AvData(3, I)) '- AvData(4, I))
                   Case "BR": mTotal = mDebet 'mTotal + (AvData(3, I)) '- AvData(4, I))
            End Select
            
        Next I
        LblAmount = FormatNumber(mDebet, 0)
        lblABayar = FormatNumber(mCredit, 0)
     Else
        mTotal = 0
     End If
     End If
End With


Set Avdata = Nothing
CloseDB RcTotal
End Sub

Private Sub PrepareQuery()
On Error Resume Next
With MyDDE
'======================================== Edit ======================================
'Select Case Left(MyDDE.GetFieldByName("TransID"), 2)
      ' Case "BP", "BR":
            .PrepareAppend = " INSERT INTO TransData (TransID, DateTrans, DateIssued, PartnerId, TypeTrans, BankID, RefNotes, Status)" & _
                             " VALUES     (N'" & NoVoucher(0) & "', CONVERT(DATETIME, '" & Format(dtDate(0), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', N'" & mVarTipe & "', N'" & NoVoucher(1) & "', N'" & ValidString(txtNotes) & "', 1)"
       'Case "BK", "BU":
          '  .PrepareAppend = " INSERT INTO TransData (TransID, DateTrans, DateIssued, empid, TypeTrans, BankID, RefNotes, Status)" & _
                             " VALUES     (N'" & NoVoucher(0) & "', CONVERT(DATETIME, '" & Format(dtDate(0), "dd/mm/yy") & "', 3), CONVERT(DATETIME, '" & Format(dDateBegin, "dd/mm/yy") & "', 3), N'" & NoVoucher(2) & "', N'" & mVarTipe & "', N'" & NoVoucher(1) & "', N'" & ValidString(txtNotes) & "', 1)"
       
'End Select
'====================================================================================
            .PrepareDelete = " DELETE FROM TransData WHERE (TransID = N'" & NoVoucher(0) & "')"
End With
Err.Clear
End Sub

Private Sub OpendDetailVoucher(ByVal ParamVoucher As String)
On Error Resume Next
Dim mVarTmp As Variant
CloseDB RcDetail
Set RcDetail = New Recordset
RcDetail.CursorLocation = adUseClient
RcDetail.Open "SELECT  TransData.TransID, TransData.DateTrans, TransData.RefNotes, [Detail TransData].Debet, [Detail TransData].Credit FROM [Detail TransData] INNER JOIN TransData ON [Detail TransData].TransID = TransData.TransID WHERE     (TransData.TransID = N'" & ParamVoucher & "')", CNN, adOpenStatic, adLockBatchOptimistic, adCmdText
Set DGPurchase.DataSource = RcDetail
Select Case Left(ParamVoucher, 2)
'       Case "BP", "PP", "BE":
'            DGPurchase.Columns(3).Visible = False
'            DGPurchase.Columns(4).Visible = True
'            HitungTotal True
'            LblAmount = FormatNumber((LblAmount - lblABayar), 0)
       Case "BR":
            DGPurchase.Columns(3).Visible = True
            DGPurchase.Columns(4).Visible = False
            DGPurchase.Columns(2).width = 4000
            HitungTotal True
            mVarTmp = FormatNumber((lblABayar - LblAmount), 0)
            lblABayar = LblAmount
            LblAmount = mVarTmp
'       Case "BK":
'            DGPurchase.Columns(3).Visible = False
'            DGPurchase.Columns(4).Visible = True
'            HitungTotal True
'            mVarTmp = FormatNumber((lblABayar - LblAmount), 0)
'            lblABayar = LblAmount
'            LblAmount = mVarTmp
'       Case "BU":
'            DGPurchase.Columns(3).Visible = False
'            DGPurchase.Columns(4).Visible = True
'            HitungTotal True
'            mVarTmp = FormatNumber((lblABayar - LblAmount), 0)
'            lblABayar = mVarTmp
'            LblAmount = LblAmount
            'LblAmount = FormatNumber((LblAmount - lblABayar), 0)
End Select


Err.Clear
End Sub

Private Function CekBayar(ByVal PartnerId As String, ByVal NoBukti As String) As Variant
Dim RcByr As New DBQuick
' ============================================== Edit ===================================
'Select Case cboVoucher
       'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses": RcByr.DBOpen " SELECT     Debet FROM         [Voucher Batch] WHERE     (TransID = N'" & NoBukti & "') AND (PartnerID = N'" & PartnerId & "')", CNN, lckLockReadOnly
       'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses":
       RcByr.DBOpen " SELECT     Debet FROM         [Voucher Batch] WHERE     (TransID = N'" & NoBukti & "') AND (PartnerID = N'" & PartnerId & "')", CNN, lckLockReadOnly
       'Case "Piutang Karyawan", "Pembayaran Piutang Karyawan": RcByr.DBOpen " SELECT    Jumlah AS Debet FROM         [BKK Karyawan] WHERE   [No Piutang]=N'" & NoBukti & "' and   (EmpID = N'" & PartnerId & "') ORDER BY [No Piutang]", CNN, lckLockReadOnly
'End Select
'========================================================================================
With RcByr.DBRecordset
     If .Recordcount <> 0 Then
        CekBayar = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekBayar = 0
     End If
End With
RcByr.CloseDB
End Function

Private Function PrepareByrVoucher() As Boolean
Dim RcByr As New DBQuick
Dim Avdata As Variant
Dim iByr As Integer
Dim Mi As Byte

Set RcByr.DBRecordset = RcDetail.Clone(adLockReadOnly)
With RcByr.DBRecordset
     If .Recordcount <> 0 Then
        Avdata = .Getrows(.Recordcount, adBookmarkFirst)
        For iByr = 0 To UBound(Avdata, 2)
            If Avdata(6, iByr) = False Then
               Mi = 0
            Else
               Mi = 1
            End If
            '===================================== Edit ==========================
            'Select Case cboVoucher
            '       Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses": PrepareByrVoucher = SendDataToServer(" UPDATE [Voucher Batch] SET  status = " & Mi & ",Debet = " & CCur(Avdata(3, iByr)) & ", Credit = " & CCur(Avdata(4, iByr)) & " WHERE (TransID = N'" & Avdata(0, iByr) & "') and (PartnerID=N'" & NoVoucher(2) & "')")
            'Case "Pelunasan Hutang", "Pelunasan Piutang", "Petty Cash", "Expenses":
            PrepareByrVoucher = SendDataToServer(" UPDATE [Voucher Batch] SET  status = " & Mi & ",Debet = " & CCur(Avdata(3, iByr)) & ", Credit = " & CCur(Avdata(4, iByr)) & " WHERE (TransID = N'" & Avdata(0, iByr) & "') and (PartnerID=N'" & NoVoucher(2) & "')")
            '       Case "Piutang Karyawan", "Pembayaran Piutang Karyawan": PrepareByrVoucher = SendDataToServer(" UPDATE [BKK KARYAWAN] SET  status = " & Mi & ",Jumlah = " & CCur(Avdata(3, iByr)) & ", Angsuran = " & CCur(Avdata(4, iByr)) & " WHERE ([No Piutang] = N'" & Avdata(0, iByr) & "') and (EmpID=N'" & NoVoucher(2) & "')")
            'End Select
            '=====================================================================
        Next iByr
     End If
End With
RcByr.CloseDB
End Function

Private Function CekGrid() As Boolean
Dim RcGrd As New DBQuick
Set RcGrd.DBRecordset = RcDetail.Clone(adLockReadOnly)
RcGrd.DBRecordset.Filter = "Credit <> 0"
With RcGrd.DBRecordset
     If .Recordcount <> 0 Then
        CekGrid = True
     Else
        CekGrid = False
     End If
End With
RcGrd.CloseDB
End Function

Private Sub ClosedTransaksi(ByVal NoRNBukti As String)
Dim TmpPo As String
If CBool(RcDetail.Fields("Status")) = True Then
   TmpPo = CekBuktiRN(NoRNBukti)
   SendDataToServer ("UPDATE TransData SET  StatusInvoice =1 WHERE     (TransID = N'" & NoRNBukti & "') AND (PurchaseID = N'" & TmpPo & "')")
   SendDataToServer ("UPDATE    [PO Order] SET              StatusSJ =1 WHERE     (PurchaseID = N'" & TmpPo & "')")
End If
End Sub

Private Function CekBuktiRN(ByVal NoRNBukti As String) As String
Dim RcBkt As New DBQuick
RcBkt.DBOpen "SELECT     PurchaseID FROM   TransData WHERE     (TransID = N'" & NoRNBukti & "') GROUP BY PurchaseID", CNN, lckLockReadOnly
With RcBkt.DBRecordset
     If .Recordcount <> 0 Then
        CekBuktiRN = IIf(Not IsNull(.Fields(0)), .Fields(0), "XXXXXXX")
     Else
        CekBuktiRN = "xxxxxxxx"
     End If
End With
RcBkt.CloseDB
End Function

Private Function PORef(ByVal NoRefData As String) As String
Dim Mrc As New DBQuick
Mrc.DBOpen "SELECT     PurchaseID FROM         TransData  WHERE     (TransID = N'" & NoRefData & "')", CNN, lckLockReadOnly
PORef = "XXXXX"
If Mrc.Recordcount <> 0 Then
   PORef = Mrc.Fields(0)
End If
Mrc.CloseDB
End Function

Private Sub OpenPartnerVoucher(ByVal NoPartner As String, ByVal NoEmpId As String)
Dim RcPrt As New DBQuick
'Select Case Left(MyDDE.GetFieldByName("Transid"), 2)
'       Case "BP", "BR":
             RcPrt.DBOpen "SELECT     PartnerID, CompanyName, Address, City, Phone FROM         PartnerDB WHERE     (PartnerID = N'" & NoPartner & "')", CNN, lckLockReadOnly
'       Case "BK", "BU":
'            RcPrt.DBOpen "SELECT     EmpID, FullName FROM         Employees WHERE     (EmpID = N'" & NoEmpId & "')", CNN, lckLockReadOnly
'End Select

NoVoucher(2) = ""
lblAlamat = ""
With RcPrt
     If .Recordcount <> 0 Then
        'Select Case Left(MyDDE.GetFieldByName("Transid"), 2)
        '       Case "BP", "BR":
                    NoVoucher(2) = IIf(Not IsNull(.DBRecordset.Fields(0)), .DBRecordset.Fields(0), "")
                    lblAlamat = IIf(Not IsNull(.DBRecordset.Fields(1)), .DBRecordset.Fields(1), "") & vbCrLf & _
                                IIf(Not IsNull(.DBRecordset.Fields(2)), .DBRecordset.Fields(2), "") & vbCrLf & _
                                IIf(Not IsNull(.DBRecordset.Fields(3)), .DBRecordset.Fields(3), "")
        '      Case "BK", "BU":
        '           NoVoucher(2) = IIf(Not IsNull(.DBRecordset.Fields(0)), .DBRecordset.Fields(0), "")
        '           lblAlamat = IIf(Not IsNull(.DBRecordset.Fields(1)), .DBRecordset.Fields(1), "")
        'End Select
     End If
End With
RcPrt.CloseDB
End Sub

Private Function IdxAuto() As String
'================================================================== Edit =================
'Select Case cboVoucher
       'Case "Pelunasan Hutang", "Piutang Karyawan": IdxAuto = MyData.PrepareIndex(tmbVoucher, 5, mVarTipe, TglIndexJournal)
       'Case "Pelunasan Hutang", "Piutang Karyawan":
       IdxAuto = MyData.PrepareIndex(tmbVoucher, 5, mVarTipe, TglIndexJournal)
       'Case "Pelunasan Piutang", "Pembayaran Piutang Karyawan": IdxAuto = MyData.PrepareIndex(tmbVoucher, 5, mVarTipe, TglIndexJournal)
'End Select
'=========================================================================================
End Function

Private Function TglIndexJournal() As String
Dim TglHari As String
Dim TglBulan As String
Dim TglTahun As String

'Select Case cboVoucher
       'Case "Pelunasan Hutang", "Piutang Karyawan": TglIndexJournal = "BP/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
       'Case "Pelunasan Hutang", "Piutang Karyawan":
        TglIndexJournal = "BP/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
       'Case "Pelunasan Piutang", "Pembayaran Piutang Karyawan": TglIndexJournal = "BR/" & Format(Day(dDateBegin), "0#") & Format(Month(dDateBegin), "0#") & Right(Format(Year(dDateBegin), "0#"), 2) & "-"
'End Select
End Function

Private Function CariTypeAccount(ByVal Params As Long) As String
Dim RcAkum As New DBQuick
RcAkum.DBOpen "SELECT     GLAccount.NoAccount, AccType.ID, GLAccount.AccountName FROM         AccType INNER JOIN                       GLAccount ON AccType.Tipe = GLAccount.Type WHERE     (GLAccount.[Group] = N'Detail List Account') AND (AccType.ID = " & Params & ")", CNN, lckLockReadOnly
With RcAkum.DBRecordset
     If .Recordcount <> 0 Then
        CariTypeAccount = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
     End If
End With
End Function
Private Sub GridLayout()
With DGPurchase
    .Columns(0).width = 2000    'TRANSID
    .Columns(1).width = 1600    'DateTrans
    .Columns(2).width = 1500    'RefNotes
    .Columns(3).width = 2500    'Debet
    .Columns(4).width = 2500    'Credit
    .Columns(0).Caption = "Nomor Bukti"
    .Columns(1).Caption = "Tanggal Bukti"
    .Columns(2).Caption = "Keterangan"
    .Columns(3).Caption = "Total Piutang"
    .Columns(1).NumberFormat = ShortDateFormGaris
    .HoldFields
End With
End Sub
