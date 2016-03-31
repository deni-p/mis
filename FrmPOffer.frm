VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPurchaseOffer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Offer"
   ClientHeight    =   5880
   ClientLeft      =   1635
   ClientTop       =   1920
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "FrmPOffer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Tag             =   "Purchase Order"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   11460
      TabIndex        =   6
      Top             =   0
      Width           =   11460
      Begin VB.TextBox txtBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   2
         Left            =   9090
         MaxLength       =   15
         TabIndex        =   20
         Top             =   4935
         Width           =   2235
      End
      Begin VB.TextBox lblBank 
         Appearance      =   0  'Flat
         DataField       =   "CompanyName"
         DataSource      =   "MyDDE"
         Height          =   330
         Index           =   0
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "SPPH"
         Top             =   850
         Width           =   2925
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4515
         MaskColor       =   &H000000C0&
         Picture         =   "FrmPOffer.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "SPPH"
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "SPPHID"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   0
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "SPPH"
         Top             =   150
         Width           =   2235
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "SPHID"
         DataSource      =   "MyDDE"
         Height          =   315
         Index           =   1
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "SPPH"
         Top             =   1200
         Width           =   2235
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   2865
         Left            =   105
         TabIndex        =   8
         Top             =   2040
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   5054
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "NoItem"
            Caption         =   "No ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ItemName"
            Caption         =   "Nama Barang"
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
            DataField       =   "uom"
            Caption         =   "Unit"
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
            DataField       =   "QTY_spph"
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
            DataField       =   "Price"
            Caption         =   "Harga"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fTotal"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Discount"
            Caption         =   "Diskon"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "GrandTotal"
            Caption         =   "Grand Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "refNote"
            Caption         =   "Keterangan"
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
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DateTrans"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1590
         TabIndex        =   2
         Tag             =   "SPPH"
         Top             =   480
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
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
         Format          =   72417283
         CurrentDate     =   38272
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "DateSPH"
         DataSource      =   "MyDDE"
         Height          =   330
         Left            =   1590
         TabIndex        =   5
         Tag             =   "SPPH"
         Top             =   1545
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
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
         Format          =   72417283
         CurrentDate     =   38272
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   7620
         X2              =   9120
         Y1              =   5235
         Y2              =   5235
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   7620
         TabIndex        =   21
         Top             =   4995
         Width           =   1050
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   120
         X2              =   1620
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   120
         X2              =   1620
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   1620
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   555
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   885
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No SPPH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblSupp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         DataField       =   "Address"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   7320
         TabIndex        =   16
         Tag             =   "spph"
         Top             =   480
         Width           =   3570
      End
      Begin VB.Label lblSupp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         DataField       =   "city"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   7320
         TabIndex        =   15
         Tag             =   "spph"
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label lblSupp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         DataField       =   "PostalCode"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   7320
         TabIndex        =   14
         Tag             =   "spph"
         Top             =   960
         Width           =   3750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   6480
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   6480
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblSupp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         DataField       =   "CompanyName"
         DataSource      =   "MyDDE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   7320
         TabIndex        =   11
         Tag             =   "spph"
         Top             =   240
         Width           =   3750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. SPH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1260
         Width           =   675
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   1620
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal SPH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1635
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   120
         X2              =   1620
         Y1              =   1860
         Y2              =   1860
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5310
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   1005
      BindFormTAG     =   "SPPH"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPurchaseOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMytr                                           As New DBQuick
Private RcUang                                            As New DBQuick
Private RcDetail                                          As New DBQuick
Attribute RcDetail.VB_VarHelpID = -1
Private RcPartner                                         As New DBQuick
Private WithEvents mCall                                  As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData                                            As New clsTransaksi
Private MEdit, mEditPO, mFirstCaller, mVarDetailPOClose   As Boolean
Private mAccount                                          As String
Dim SQLInit As String
Private IsSPH As Boolean
Private IsSPPH As Boolean
Dim IDGen As New IDGenerator
Private xParams As String
Private isHistoryMode As Boolean


Public Property Let IDParams(vData As String)
   isHistoryMode = True
   xParams = vData
End Property


Public Property Let OperationMode(vData As String)
   If vData = "SPPH" Then IsSPPH = True Else IsSPPH = False
   IsSPH = Not IsSPPH
End Property


Private Sub CboBayar_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub CboUang_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub cmdLink_Click(Index As Integer)
  If MyDDE.ChildRecordset.Recordcount > 0 Then
   MessageBox "Item Barang Sudah Diisi Berdasarkan Supplier Lain. " & Chr(13) & "Hapus Semua Item untuk mengganti Supplier", "Peringatan", msgOkOnly, msgCrtical
  Else
   OpenPartner Index
  End If

End Sub

Private Sub cmdLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then OpenPartner Index
End Sub

Private Sub DGPurchase_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
   If MEdit = True Then
      Select Case ColIndex
         Case 3, 4, 5, 6, 7:
            DGPurchase.AllowUpdate = True
      End Select
   End If
End Sub

Private Sub DGPurchase_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
If IsSPPH Then
   Me.Caption = "Surat Permintaan Penawaran Harga"
   MyDDE.SetPermissions = aksess.MayDo("Permintaan Penawaran Harga")
End If

If IsSPH Then
   Me.Caption = "Surat Penawaran Harga"
   MyDDE.SetPermissions = aksess.MayDo("Penawaran Harga")
End If

If IsSPH Then
   MyDDE.SetPermissions = UserAddnewDenied
Else
   MyDDE.SetPermissions = UserOk
End If

HiasFormManTell Picture2, Me
''HiasForm Picture1, Me
Set mCall = New frmCaller
DTPicker1.Value = dDateBegin
LoadData
End Sub

Private Sub LoadData()
   With MyDDE
       .EditModeReplace = False
       Set .BindForm = Me
       Set .ActiveConnection = CNN
       If isHistoryMode Then
         .PrepareQuery = " SELECT * FROM QueryPurchaseOffer where SPPHID='" & xParams & "'"
       Else
         .PrepareQuery = " SELECT * FROM QueryPurchaseOffer where Status = 0"
       End If
   End With
   
   lbl(0).Visible = IsSPH
   lbl(1).Visible = IsSPH
   Line1(1).Visible = IsSPH
   Line1(2).Visible = IsSPH
   txtBox(1).Visible = IsSPH
   DTPicker2.Visible = IsSPH
   MyDDE.SetReadOnlyMode = isHistoryMode
   
End Sub

 

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
RcUang.CloseDB
clsMytr.CloseDB
Set mCall = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next

Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmPurchasing = Nothing
xParams = ""
isHistoryMode = False
End Sub

Private Sub mCall_BeforeUnload()
On Error Resume Next
Select Case mCall.FromTagActive
       Case "Inventory List":
            If FindOwnRecordset(MyDDE.ChildRecordset, "NoItem = '" & MyDDE.ChildRecordset.Fields("NoItem") & "'") = True Then
               MessageBox "Record -> " & MyDDE.ChildRecordset.Fields("Noitem") & " Sudah Ada....! Silahkan Diulangi", "Peringatan", msgOkOnly, msgCrtical
               MyDDE.ChildRecordset.CancelBatch adAffectCurrent
               If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
            Else
               If Not IsNull(MyDDE.ChildRecordset.Fields(0)) = True Then
                  If MyDDE.ChildRecordset.Fields(0) = "" Then
                     MyDDE.ChildRecordset.CancelBatch adAffectCurrent
                     If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
                  End If
               End If
            End If
            mFirstCaller = False
            If DGPurchase.Enabled = True Then
               DGPurchase.AllowUpdate = True
               DGPurchase.col = 3
               DGPurchase.SetFocus
            End If
       Case "Supplier List":
            txtBox(1).SetFocus
End Select
End Sub

Private Sub mCall_CallLinkForm()
If mCall.FromTagActive <> "MASTER BARANG" Then
   frmMasterSup.SetFocus
   frmMasterSup.ZOrder (0)
Else
   FrmItemData.SetFocus
   FrmItemData.ZOrder (0)
End If
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If pRecordset.Recordcount <> 0 Then
Select Case TagForm:
      Case "Supplier List":
          MyDDE.GetFieldByName("PartnerID") = mCall.GetFieldByName(0)
          MyDDE.GetFieldByName("CompanyName") = mCall.GetFieldByName(1)
          MyDDE.GetFieldByName("Address") = mCall.GetFieldByName("Alamat")
          MyDDE.GetFieldByName("City") = mCall.GetFieldByName("Kota")
          MyDDE.GetFieldByName("PostalCode") = mCall.GetFieldByName("Kode pos")
      Case "Inventory List":
         If mCall.GetFieldByName("NoItem") = "" Then
            MyDDE.ChildRecordset.Delete
         Else
            MyDDE.ChildRecordset.Fields("NoItem").Value = mCall.GetFieldByName("NoItem")
            MyDDE.ChildRecordset.Fields("ItemName").Value = mCall.GetFieldByName("ItemName")
            MyDDE.ChildRecordset.Fields("UOM").Value = mCall.GetFieldByName("UOM")
            MyDDE.ChildRecordset.Fields("QTY_SPPH").Value = 1
            MyDDE.ChildRecordset.Fields("Price").Value = 0
            MyDDE.ChildRecordset.Fields("Discount").Value = 0
            MyDDE.ChildRecordset.Fields("RefNote").Value = ""
            MyDDE.ChildRecordset.Fields("fTotal").Value = 0
            MyDDE.ChildRecordset.Fields("GrandTotal").Value = 0
         End If
End Select
End If
End Sub

Private Sub DGPurchase_AfterColEdit(ByVal ColIndex As Integer)
Dim I As Integer
Dim mStok As Long
Dim lTotal, lGrandTotal As Double
Dim mTmp As Variant
Select Case ColIndex
       Case 3, 4, 5, 6:
          If DGPurchase.Columns(ColIndex) = "" Or IsNull(DGPurchase.Columns(ColIndex)) Then DGPurchase.Columns(ColIndex).Value = 0
          If MyDDE.ChildRecordset.Fields("Qty_spph").Value = 0 Or MyDDE.ChildRecordset.Fields("Price").Value = 0 Then
            DGPurchase.Columns(5).Text = "0"
          Else
            lTotal = Val(MyDDE.ChildRecordset.Fields("Qty_spph").Value) * Val(MyDDE.ChildRecordset.Fields("Price").Value)
            MyDDE.ChildRecordset.Fields("fTotal").Value = lTotal
            MyDDE.ChildRecordset.Fields("GrandTotal") = Val(MyDDE.ChildRecordset.Fields("fTotal").Value) - (Val(MyDDE.ChildRecordset.Fields("fTotal").Value) * Val(MyDDE.ChildRecordset.Fields("discount").Value) / 100)
          End If
       Case 7:
          If DGPurchase.Columns(ColIndex) = "" Or IsNull(DGPurchase.Columns(ColIndex)) Then DGPurchase.Columns(ColIndex).Value = " "
End Select
'HitungTotal
End Sub

Private Sub DGPurchase_KeyDown(KeyCode As Integer, Shift As Integer)
If MEdit = False Then Exit Sub
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If MEdit = False Then
   DGPurchase.AllowUpdate = False
   DGPurchase.MarqueeStyle = dbgFloatingEditor
   Exit Sub
End If
With DGPurchase
     Select Case .col
            Case 0, 1, 2:
                .AllowUpdate = False
            Case Else:
                .AllowUpdate = True
     End Select
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_OnReverseAction()
   LoadData
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
      Case tmbAddNew:
         If IsSPH Then
            MyDDE.CancelTrans = True
            MessageBox "Transaksi Tidak Bisa Dilakukan pada Form Ini.Silahkan Entry pada form SPPH", "Peringatan", msgOkOnly, msgCrtical
         End If
      Case tmbCancel:
         
      Case tmbDelete:
            If MyDDE.ActiveRecordset.Recordcount <> 0 Then
               MyDDE.CancelTrans = Not IsHeaderOk
               If MyDDE.CancelTrans = True Then MessageBox "Transaksi Tidak Bisa Dihapus.Karena SPPH sudah dikirim", "Peringatan", msgOkOnly, msgCrtical
            End If
      Case tmbDetail:
         If IsSPH Then
            MyDDE.CancelTrans = IsSPH
            MessageBox "Menambah Item Data Hanya Bisa dilakukan pada Form SPPH", "Informasi", msgOkOnly, msgInfo
         End If

      Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               If MyDDE.ChildRecordset.Recordcount <> 0 Then
                  If IsUncompleteGridEntry Then
                     MessageBox "Data Yang dimasukkan belum lengkap !", "Informasi", msgOkOnly, msgInfo
                     MyDDE.CancelTrans = True
                  Else
                     MyDDE.IsChildMemberReady = True
                     MyDDE.GetFieldByName("DatePurchase") = DTPicker1.Value
                     PrepareQuery
                  End If
               Else
                  MessageBox "Silahkan Masukkan Item Barang dulu...", "Peringatan", msgOkOnly, msgInfo
                  MyDDE.IsChildMemberReady = False
                  MyDDE.CancelTrans = True
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
            'cmdLink(0).Enabled = False
End Select
End Sub

Private Function IsUncompleteGridEntry()
   Dim isLos As Boolean
   isLos = False
   MyDDE.ChildRecordset.MoveFirst
   Do While Not MyDDE.ChildRecordset.EOF
      If IsNull(DGPurchase.Columns(3).Value) Or DGPurchase.Columns(3).Value = 0 Then
         isLos = True
         Exit Do
      End If
      If IsSPH Then
         If IsNull(DGPurchase.Columns(4).Value) Or DGPurchase.Columns(4).Value = 0 Then
            isLos = True
            Exit Do
         End If
      End If
      MyDDE.ChildRecordset.MoveNext
   Loop
   IsUncompleteGridEntry = isLos
End Function

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'On Error Resume Next
txtBox(0).Enabled = False
'lblBank(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
            DTPicker1.Enabled = Not IsSPH
            cmdLink(0).Enabled = Not IsSPH
            Call DGPurchase_RowColChange(DGPurchase.row, DGPurchase.col)
            MEdit = True
            DTPicker2.Value = Now
            MyDDE.GetFieldByName("DateSPH") = Now
            'MyDDE.GetFieldByName("SPHID") = " "
            
       Case tmbAddNew:
            DTPicker1.Value = CDate(Format(Date, "dd/mm/yyyy"))
            MyDDE.GetFieldByName("DateTrans") = DTPicker1.Value
            MyDDE.GetFieldByName("SPPHID") = IDGen.GetID("OF")   'MyData.PrepareIndex(tmbTransaksiPO, 5, "1", TglIndex)
            MyDDE.GetFieldByName("Result") = True
            MyDDE.GetFieldByName("Status") = False
            MyDDE.GetFieldByName("DateSPH") = Now
            MyDDE.GetFieldByName("SPHID") = " "
            
            DTPicker1.SetFocus
            MEdit = True
            cmdLink(0).Enabled = True
            'MyDDE.ChildRecordset.AddNew
            
       Case tmbSave:
            If MyDDE.IsChildMemberReady = True Then
               SimpanDetail
               MEdit = False
               mEditPO = False
               'MyData.EditHeaderRN txtBox(0), mVarLoginActive, CboUang.BoundText, MyDDE.GetFieldByName("PartnerID"), txtBox(1), CDbl(txtBox(2)), txtBox(4), False, MyDDE.ChildRecordset
               OpenDetail txtBox(0)
               mVarDetailPOClose = False
               cmdLink(0).Enabled = False
            Else
               MessageBox "Detail transaksi Purchase belum ada datanya.", "Peringatan", msgOkOnly, msgCrtical
            End If

       Case tmbCancel:
            MyDDE.CancelTrans = True
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               MEdit = False
               mVarDetailPOClose = False
             Else
               'DGPurchase.Columns(6).Visible = False
               'DGPurchase.Columns(7).Visible = True
             End If
             cmdLink(0).Enabled = False
       Case tmbDetail:
               OpenPartner 3
               
       Case tmbPrint:
            If Not IsNull(MyDDE.GetFieldByName("Approved_by")) Then
               Dim ReportView As New utility
               ReportView.CallReportView "Select * From Spph_report where SpphID ='" & txtBox(0) & "'", "Spph.rpt", ReportPath, Caption
               'CallRPTReport "Spph.rpt", "Select * From Spph_report where SpphID ='" & txtBox(0) & "'"
               Set ReportView = Nothing
            Else
               MessageBox "Dokumen ini belum di Approve", "Informasi", msgOkOnly, msgInfo
            End If
       Case tmbQuit:
            Unload Me
            Set MyDDE.BindForm = Nothing
End Select

Err.Clear
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If IsNull(MyDDE.GetFieldByName("SPPHID")) Then
   OpenDetail ""
Else
   OpenDetail MyDDE.GetFieldByName("SPPHID")
End If

txtBox(2) = IIf(IsNull(MyDDE.GetFieldByName("Approved_by")), "", MyDDE.GetFieldByName("Approved_by"))

MEdit = False
End Sub

Private Sub OpenPartner(ByVal Index As Integer)
Dim strSQL As String
On Error GoTo Hell:
Select Case Index
       Case 0:
            RcPartner.DBOpen MyData.UploadQuery("Supplier"), CNN, lckLockReadOnly
       Case 1:
            RcPartner.DBOpen MyData.UploadQuery("BANK", MyDDE.GetFieldByName("PartnerID")), CNN, lckLockReadOnly
       Case 2:
            RcPartner.DBOpen "SELECT NoItem AS [No Barang], ItemName AS [Nama Barang], UOM, PPn,PriceIn AS Harga FROM Inventory WHERE     (Manufacture = 0) ORDER BY NoItem", CNN, lckLockReadOnly
       Case 3:
            strSQL = "SELECT inventory.NoItem, Inventory.ItemName, Inventory.UOM FROM inventory where PartnerID='" & MyDDE.GetFieldByName("PartnerID") & "' and inventory.Manufacture = 0"
            Debug.Print strSQL
            RcPartner.DBOpen strSQL, CNN, lckLockReadOnly
            mFirstCaller = True
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "Supplier List"
            mCall.txtCari = lblBank(0)
            mCall.CaptionLink = "Supplier"
          Case 1:
            mCall.FromTagActive = "Bank List"
            mCall.txtCari = lblBank(1)
          Case 2:
            mCall.FromTagActive = "Remindier"
            mCall.txtCari = lblBank(1)
          Case 3:
            mCall.FromTagActive = "Inventory List"
            mCall.CaptionLink = "Barang"
            'If MyDDE.ChildRecordset.Recordcount <> 0 Then mCall.txtCari = MyDDE.ChildRecordset.Fields("Noitem")
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
   
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly, msgCrtical
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If
End If
'
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub OpenDetail(ByVal ParameterString As String)
Set RcDetail = New DBQuick
If ParameterString = "" Then ParameterString = "xxxxxxxx"
RcDetail.DBOpen "SELECT  SPPHID, NoItem, itemName, UOM, Qty_SPPH, Price, ftotal, Discount,ftotal as grandTotal, RefNote from QueryDetailPurchaseOffer WHERE  (SPPHID = N'" & ParameterString & "') ", CNN, lckLockBatch
'MessageBox Rcdetail.DBRecordset.Source
Set MyDDE.ChildRecordset = RcDetail.DBRecordset.Clone(adLockBatchOptimistic)
RcDetail.CloseDB
Set DGPurchase.DataSource = MyDDE.ChildRecordset
   
   With MyDDE.ChildRecordset
      If .Recordcount > 0 Then
         While Not .EOF
            .Fields("fTotal") = Val(.Fields("price")) * Val(.Fields("Qty_SPPH"))
            .Fields("GrandTotal") = Val(.Fields("ftotal")) - (Val(.Fields("fTotal")) * Val(.Fields("discount")) / 100)
            .MoveNext
         Wend
      End If
   End With

End Sub

Private Sub SimpanDetail()
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           .MoveFirst
           If SendDataToServer("DELETE FROM [SPPH_line] WHERE  (SPPHID = '" & MyDDE.GetFieldByName("SPPHID") & "')") = True Then
           Do
              If .EOF = True Then Exit Do
              SendDataToServer " INSERT INTO [SPPH_line] (SPPHID, noItem, QTY_SPPH, Price, Discount, refnote) " & _
                               " VALUES (N'" & MyDDE.GetFieldByName("SPPHID") & _
                                     "', N'" & .Fields("NoItem") & _
                                     "', " & FQty(.Fields("QTY_SPPH")) & _
                                     ", " & FQty(.Fields("Price")) & _
                                     ", " & FQty(.Fields("Discount")) & _
                                     ",'" & .Fields("RefNote") & "')"
              .MoveNext
           Loop
           End If
           .MoveLast
           DGPurchase.Refresh
     End If
End With
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub


Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then KeyEnter KeyCode
End Sub

Private Function TglIndex() As String
Dim TglHari, TglBulan, TglTahun As String
TglIndex = "PO/" & Format(Day(Date), "0#") & Format(Month(Date), "0#") & Right(Format(Year(Date), "0#"), 2) & "-"
End Function

Private Sub HitungTotal()
On Error Resume Next
Dim RcTotal As New DBQuick
Dim Avdata As Variant
Dim mDisc, mPPn, mTotal, mStDisc As Variant
Dim mTmpDisc As Byte
Dim I As Long
Set RcTotal.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
mTotal = 0
mDisc = 0
mPPn = 0
mStDisc = 0
mTmpDisc = IIf(Not IsNull(MyDDE.GetFieldByName("Discount")), MyDDE.GetFieldByName("Discount"), 0)
With RcTotal
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        ' 3 = QTY  4 = Harga 5 = Vat
        For I = 0 To UBound(Avdata, 2)
            If mTmpDisc > 0 Then
               mDisc = mDisc + (Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100)
               mStDisc = mStDisc + ((Avdata(3, I) * Avdata(4, I)) - ((Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100)))
            Else
               mStDisc = mStDisc + (Avdata(3, I) * Avdata(4, I))
               mDisc = mDisc + 0
            End If
            If Avdata(5, I) > 0 Then
               mPPn = mPPn + ((((Avdata(3, I) * Avdata(4, I)) - ((Avdata(3, I) * Avdata(4, I)) * (mTmpDisc / 100))) * (Avdata(5, I) / 100)))
            Else
               mPPn = mPPn + 0
            End If
            mTotal = mTotal + Avdata(3, I) * Avdata(4, I)
        Next I
     Else
        mTotal = 0
     End If
End With
Set Avdata = Nothing
Set mTotal = Nothing
Set mDisc = Nothing
Set mPPn = Nothing
Set mStDisc = Nothing
Err.Clear
End Sub

Private Sub PrepareQuery()
On Error Resume Next
Dim mPoSc As String
Dim strSQL As String
With MyDDE
    .PrepareAppend = " INSERT INTO  SPPH_Header(SPPHID,DateTrans,PartnerID,UserReqst,Result,status) " & _
                     " VALUES ('" & .GetFieldByName("SPPHID") & _
                            "','" & Format(.GetFieldByName("DateTrans"), "yyyy-MM-dd") & _
                            "','" & .GetFieldByName("PartnerID") & _
                            "','" & .GetFieldByName("UserReqst") & _
                            "',1,0)"
                     
    strSQL = " UPDATE SPPH_Header Set DateTrans ='" & Format(.GetFieldByName("DateTrans"), "yyyy-mm-dd") & _
                                           "',PartnerID ='" & .GetFieldByName("PartnerID") & _
                                           "',UserReqst ='" & .GetFieldByName("UserReqst") & _
                                           "',Result =" & SQLBoolean(.GetFieldByName("result")) & _
                                           ",status =" & "0" & _
                                           ",SPHID =" & IIf(IsSPH, "'" & .GetFieldByName("SPHID") & "'", "Null") & _
                                           ",DateSPH =" & IIf(IsSPH, "'" & Format(.GetFieldByName("DateSPH"), "yyyy-MM-dd") & "'", "Null") & _
                     " where SPPHID='" & .GetFieldByName("SPPHID") & "'"
   Debug.Print strSQL
   .PrepareUpdate = strSQL
                     
   If IsSPPH Then
      .PrepareDelete = " DELETE FROM  [SPPH_Header] WHERE (SPPHID = '" & .GetFieldByName("SPPHID") & "')"
   Else
      .PrepareDelete = " DELETE FROM  [SPPH_Header] WHERE (SPPHID = 'xxxxxxx')"
   End If
End With
Err.Clear
End Sub

Private Function SQLBoolean(Value As Boolean) As String
   If Value Then
      SQLBoolean = "1"
   Else
      SQLBoolean = "0"
   End If
End Function

Private Function IsHeaderOk() As Boolean
If IsSPPH Then
   IsHeaderOk = True
Else
   IsHeaderOk = False
End If
'Dim RcIs As New DBQuick
'RcIs.DBOpen "SELECT  StatusSJ FROM [PO Order] WHERE     (PurchaseID = N'" & NoPo & "')", CNN, lckLockReadOnly
'IsHeaderOk = False
'With RcIs
'     If .Recordcount <> 0 Then IsHeaderOk = CBool(.Fields(0))
'End With
'RcIs.CloseDB
End Function

Private Function IsStatusPO(Optional ByVal NoItem As String) As Boolean
Dim RcIs As New DBQuick
If NoItem = "" Then
   RcIs.DBOpen "SELECT SUM(QTY_Receive) AS QTY FROM [Detail TransData] WHERE     (DNID = N'" & txtBox(0) & "')", CNN, lckLockReadOnly
Else
   RcIs.DBOpen "SELECT     QTY_Receive AS QTY FROM         [Detail TransData] WHERE     (DNID = N'" & txtBox(0) & "') AND (NoItem = N'" & NoItem & "')", CNN, lckLockReadOnly
End If
With RcIs
     If .Recordcount <> 0 Then If .Fields(0) <> 0 Then IsStatusPO = True
End With
RcIs.CloseDB
End Function

'Private Function IsDetailOK(ByVal Noitem As String) As Boolean
'Dim RcIs As New DBQuick
'RcIs.DBOpen "SELECT     [Detail PO].StatusTrans FROM         [Detail PO] INNER JOIN                       [PO Order] ON [Detail PO].PurchaseID = [PO Order].PurchaseID WHERE     ([PO Order].PurchaseID = N'" & txtBox(0) & "') AND ([Detail PO].NoItem = N'" & Noitem & "') GROUP BY [Detail PO].StatusTrans HAVING      ([Detail PO].StatusTrans = 1)", Cnn, lckLockReadOnly
'With RcIs
'     If .Recordcount <> 0 Then IsDetailOK = CBool(.Fields(0))
'End With
'RcIs.CloseDB
'Set RcIs = Nothing
'End Function

Private Sub OpenTypeBayarPO()
clsMytr.DBOpen MyData.UploadQuery("franco beli"), CNN, lckLockReadOnly
'Set CboBayar.RowSource = clsMytr.DBRecordset
End Sub

Private Sub MataUang()
RcUang.DBOpen MyData.UploadQuery("mata uang"), CNN, lckLockReadOnly
'Set CboUang.RowSource = RcUang.DBRecordset
End Sub

Private Sub UpdateTotal()
Dim rcUpdate As New DBQuick
Dim iLast, mRow As Integer
Dim Avdata As Variant
Set rcUpdate.DBRecordset = MyDDE.ChildRecordset.Clone(adLockBatchOptimistic)
With rcUpdate
     If .Recordcount <> 0 Then
        mRow = MyDDE.ChildRecordset.AbsolutePosition
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        For iLast = 0 To UBound(Avdata, 2)
            .AbsolutePosition = iLast + 1
            .Fields("Tmp") = Avdata(7, iLast)
        Next iLast
     End If
End With
Set MyDDE.ChildRecordset = rcUpdate.DBRecordset.Clone(adLockBatchOptimistic)
If MyDDE.ChildRecordset.Recordcount <> 0 Then
   MyDDE.ChildRecordset.AbsolutePosition = mRow
End If
rcUpdate.CloseDB
End Sub

Private Function CekDetailItem(ByVal PoNumber As String, ByVal NoItemData As String) As Boolean
Dim RcCek As New DBQuick
RcCek.DBOpen "SELECT NoItem, PurchaseID FROM [Detail PO] WHERE     (NoItem = N'" & NoItemData & "') AND (PurchaseID = N'" & PoNumber & "')", CNN, lckLockReadOnly
If RcCek.Recordcount <> 0 Then CekDetailItem = True
RcCek.CloseDB
End Function

Private Sub ListTotalDeliver(ByVal ParamString As String)
Dim RcDN As New DBQuick
If ParamString = "" Then ParamString = "XXXXX"
RcDN.DBOpen "SELECT DateTrans FROM TransData GROUP BY DateTrans, PurchaseID HAVING      (PurchaseID = N'" & ParamString & "')", CNN, lckLockReadOnly
With RcDN
     If .Recordcount <> 0 Then
        'LblDeliVer = Abs(CDate(Format(MyDDE.GetFieldByName("DatePurchase"), "dd/mm/yyyy")) - CDate(Format(.Fields(0), "dd/mm/yyyy")))
     Else
        'LblDeliVer = 0
     End If
End With
End Sub

Private Function CekGridKosong() As Boolean
Dim RcKsg As New DBQuick
Dim Avdata As Variant
Dim I As Integer
Dim Temp As String
Set RcKsg.DBRecordset = MyDDE.ChildRecordset.Clone(adLockReadOnly)
With RcKsg
     If .Recordcount <> 0 Then
        Avdata = .DBRecordset.Getrows(.Recordcount, adBookmarkFirst)
        For I = 0 To UBound(Avdata, 2)
            Temp = IIf(Not IsNull(Avdata(0, I)), Avdata(0, I), "")
            If Temp <> "" Then
                If Val(IIf(Not IsNull(Avdata(3, I)), Avdata(3, I), 0)) = 0 Or Val(IIf(Not IsNull(Avdata(4, I)), Avdata(4, I), 0)) = 0 Then
                   MessageBox "Quantity Atau Harga harus diisi.", "Peringatan", msgOkOnly, msgCrtical
                   CekGridKosong = True
                   Exit For
                End If
            Else
               MessageBox "Data Item Tidak Lengkap.Harap Dicek Dulu", "Peringatan", msgOkOnly, msgCrtical
               CekGridKosong = True
               Exit For
            End If
        Next I
     Else
        CekGridKosong = True
     End If
End With
RcKsg.CloseDB
End Function

Private Function CekStock(ByVal NoItem As String) As Long
Dim RcCek As New Recordset
RcCek.CursorLocation = adUseClient
RcCek.Open "SELECT  SUM([Inventory Tabel].StockTmp)  AS QTY FROM [Inventory Tabel] INNER JOIN  Inventory ON [Inventory Tabel].NoItem = Inventory.NoItem GROUP BY [Inventory Tabel].NoItem, LEFT([Inventory Tabel].RefTrans, 2), Inventory.MinStock HAVING      (LEFT([Inventory Tabel].RefTrans, 2) = N'RN') AND ([Inventory Tabel].NoItem = N'" & NoItem & "')", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
With RcCek
     If .Recordcount <> 0 Then
        CekStock = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        CekStock = 0
     End If
     .Close
End With
Set RcCek = Nothing
End Function

Private Sub CekBankName(ByVal PartnerId As String, ByVal NoRekening As String)
Dim RcBnk As New DBQuick
RcBnk.DBOpen "SELECT     Account, [Bank Name] FROM         [Bank Partner] WHERE     (PartnerID = N'" & PartnerId & "') AND (Account = N'" & NoRekening & "')", CNN, lckLockReadOnly
With RcBnk
     If .Recordcount <> 0 Then
         lblBank(1) = .Fields(1)
     Else
         lblBank(1) = ""
     End If
End With
RcBnk.CloseDB
End Sub

Private Sub GridLayout()
DGPurchase.Columns(0).width = 1814.74
DGPurchase.Columns(1).width = 2324.977
DGPurchase.Columns(2).width = 764.7874
DGPurchase.Columns(3).width = 764.7874
DGPurchase.Columns(4).width = 1335.118
DGPurchase.Columns(5).width = 764.7874
DGPurchase.Columns(6).width = 1440
DGPurchase.Columns(7).width = 1440
DGPurchase.Columns(8).width = 1440
End Sub
