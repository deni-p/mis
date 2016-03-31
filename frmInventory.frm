VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInventory 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stok Barang"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInventory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1560
      Left            =   6060
      ScaleHeight     =   1500
      ScaleWidth      =   2565
      TabIndex        =   9
      Top             =   1620
      Visible         =   0   'False
      Width           =   2625
      Begin VB.CommandButton Command3 
         Caption         =   "&Tutup"
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   360
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   585
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   1305
         TabIndex        =   11
         Text            =   "0"
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sisa hari"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   12
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Konsumsi / hari"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   11115
      TabIndex        =   8
      Top             =   300
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Filter"
      Height          =   330
      Left            =   10425
      TabIndex        =   7
      Top             =   300
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   300
      Width           =   6435
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2805
      Top             =   1725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":6852
            Key             =   "ParentPeriode"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":D0B4
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":13916
            Key             =   "Periode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":1A178
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":209DA
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":2723C
            Key             =   "Gudang"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":2DA9E
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":34300
            Key             =   "FIFO"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":3AB62
            Key             =   "REFRESH"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":413C4
            Key             =   "CHILD"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":47C26
            Key             =   "CLOSING"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":4E488
            Key             =   "CALENDAR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":54CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":5A4DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":60D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":675A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":685F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":6EE54
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventory.frx":756B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgTraiHeader 
      Height          =   2820
      Left            =   3465
      TabIndex        =   0
      Top             =   660
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   4974
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "noItem"
         Caption         =   "Kode Barang"
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
         DataField       =   "internalName"
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
         Caption         =   "Satuan"
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
         DataField       =   "saldo"
         Caption         =   "Saldo"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Button          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   3420
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6033
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid dgTrailDetail 
      Height          =   4305
      Left            =   45
      TabIndex        =   4
      Top             =   3780
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   7594
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "sl_no"
         Caption         =   "ID"
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
         DataField       =   "REFERENCE"
         Caption         =   "Keterangan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "dateTrans"
         Caption         =   "TANGGAL MASUK"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "qty_in"
         Caption         =   "MASUK"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "qty_out"
         Caption         =   "KELUAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "StockTmp"
         Caption         =   "AKHIR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
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
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.Label Filter 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter"
      Height          =   255
      Left            =   3525
      TabIndex        =   5
      Top             =   345
      Width           =   705
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "  Data Barang Perlokasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   3465
      TabIndex        =   3
      Top             =   45
      Width           =   8385
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "  Detail Transaksi Barang By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   45
      TabIndex        =   2
      Top             =   3540
      Width           =   11805
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsInv As New DBQuick
Private rsInvDetail As New DBQuick
Private lOrder As String

Private Sub Command1_Click()
On Error GoTo xErr
Dim x As Integer
   If Text1.Text <> "" Then
      rsInv.DBOpen "select * from view_stok_header where  (noItem like '%" & Text1 & "%' or internalName like '%" & Text1 & "%' or uom like '%" & Text1 & "%') order by " & lOrder, CNN, lckLockBatch
      If rsInv.DBRecordset.Recordcount > 0 Then
         For x = 1 To TreeView2.Nodes.Count
            If TreeView2.Nodes(x).Key = rsInv.DBRecordset.Fields("warehouse") Then
               TreeView2.Nodes(x).Selected = True
            End If
         Next
      End If
   Else
      rsInv.DBOpen "Select * from view_stok_header where warehouse ='" & TreeView2.SelectedItem.Key & "' order by " & lOrder, CNN, lckLockBatch
   End If
   
   Set DgTraiHeader.DataSource = rsInv.DBRecordset
Exit Sub
xErr:
   Err.Clear
End Sub

Private Sub Command2_Click()
   Text1.Text = ""
   Command1_Click
End Sub

Private Sub Command3_Click()
   Picture1.Visible = False
End Sub

Private Sub DgTraiHeader_ButtonClick(ByVal ColIndex As Integer)
   Picture1.Visible = True
   Text2.Text = 0
   Text3.Text = 0
End Sub

Private Sub DgTraiHeader_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
      Case 0: lOrder = "noItem"
              DgTraiHeader.Columns(0).Caption = "KODE BARANG"
              DgTraiHeader.Columns(1).Caption = "Nama Barang"
              DgTraiHeader.Columns(2).Caption = "Satuan"
              DgTraiHeader.Columns(3).Caption = "Saldo"
      Case 1: lOrder = "ItemName"
              DgTraiHeader.Columns(0).Caption = "Kode Barang"
              DgTraiHeader.Columns(1).Caption = "NAMA BARANG"
              DgTraiHeader.Columns(2).Caption = "Satuan"
              DgTraiHeader.Columns(3).Caption = "Saldo"
      Case 2: lOrder = "uom"
              DgTraiHeader.Columns(0).Caption = "Kode Barang"
              DgTraiHeader.Columns(1).Caption = "Nama Barang"
              DgTraiHeader.Columns(2).Caption = "SATUAN"
              DgTraiHeader.Columns(3).Caption = "Saldo"
      Case 3: lOrder = "saldo"
              DgTraiHeader.Columns(0).Caption = "Kode Barang"
              DgTraiHeader.Columns(1).Caption = "Nama Barang"
              DgTraiHeader.Columns(2).Caption = "Satuan"
              DgTraiHeader.Columns(3).Caption = "SALDO"
   End Select
   LoadData
End Sub

Private Sub DgTraiHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo xErr
   Dim x As Integer
   
   For x = 1 To TreeView2.Nodes.Count
      If TreeView2.Nodes(x).Key = rsInv.DBRecordset.Fields("warehouse") Then
         TreeView2.Nodes(x).Selected = True
      End If
   Next
   rsInvDetail.DBOpen "select sl_no,refTrans,dateTrans,qty_in,qty_out,stocktmp from [inventory tabel] where noItem='" & rsInv.DBRecordset.Fields("noItem") & "' and lockFIFO=0", CNN, lckLockBatch
   Set dgTrailDetail.DataSource = rsInvDetail.DBRecordset
Exit Sub
xErr:
      Err.Clear
End Sub

Private Sub Form_Load()
   lOrder = "noItem"
   LoadGudang
End Sub

Private Sub LoadGudang()
   Dim rsGudang As New DBQuick
   TreeView2.ImageList = MainMenu.ImageList1
   TreeView2.Nodes.Clear
   rsGudang.DBOpen "select warehouse,[warehouse name] from warehouse", CNN, lckLockBatch
   With rsGudang.DBRecordset
      If .Recordcount > 0 Then
         While Not .EOF
            TreeView2.Nodes.Add , , .Fields("warehouse"), .Fields("warehouse name"), "biru"
            .MoveNext
         Wend
      End If
   End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text2_LostFocus()
   On Error Resume Next
   Text3.Text = Val(rsInv.DBRecordset.Fields("saldo")) / Val(Text2.Text)
End Sub

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
   Text1.Text = ""
   LoadData
End Sub

Private Sub LoadData()
On Error GoTo xErr
   If Text1.Text <> "" Then
      rsInv.DBOpen "select * from view_stok_header where warehouse ='" & TreeView2.SelectedItem.Key & "' and (noItem like '%" & Text1 & "%' or itemName like '%" & Text1 & "%' or uom like '%" & Text1 & "%') order by " & lOrder, CNN, lckLockBatch
      
   Else
      rsInv.DBOpen "Select * from view_stok_header where warehouse ='" & TreeView2.SelectedItem.Key & "' order by " & lOrder, CNN, lckLockBatch
   End If
   
   Set DgTraiHeader.DataSource = rsInv.DBRecordset
Exit Sub
xErr:
   Err.Clear
End Sub

