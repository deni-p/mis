VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmSalesForecast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forcast Penjualan"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "FrmSalesForecast.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Tag             =   "Sales Forecasting"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   11430
      TabIndex        =   11
      Top             =   0
      Width           =   11430
      Begin VB.TextBox txtperiode 
         Appearance      =   0  'Flat
         DataField       =   "periode_no"
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
         Height          =   330
         Index           =   0
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "PO"
         Top             =   870
         Width           =   1530
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "sales_pic"
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
         Index           =   2
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "PO"
         Top             =   150
         Width           =   3465
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "fcast_id"
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
         Index           =   0
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "PO"
         Top             =   150
         Width           =   3450
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "description"
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
         Index           =   1
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "PO"
         Top             =   510
         Width           =   3450
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "issue_to"
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
         Height          =   315
         ItemData        =   "FrmSalesForecast.frx":6852
         Left            =   1485
         List            =   "FrmSalesForecast.frx":6859
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "PO"
         Top             =   1950
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker FCastDate 
         DataField       =   "date_stop"
         Height          =   330
         Index           =   1
         Left            =   1485
         TabIndex        =   6
         Tag             =   "PO"
         Top             =   1584
         Width           =   1890
         _ExtentX        =   3334
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
         Format          =   139329539
         CurrentDate     =   38649
      End
      Begin MSComCtl2.DTPicker FCastDate 
         DataField       =   "date_start"
         Height          =   330
         Index           =   0
         Left            =   1485
         TabIndex        =   5
         Tag             =   "PO"
         Top             =   1230
         Width           =   1890
         _ExtentX        =   3334
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
         Format          =   139329539
         CurrentDate     =   38649
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "FrmSalesForecast.frx":6867
         Left            =   9720
         List            =   "FrmSalesForecast.frx":6869
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdTransfer 
         Caption         =   "&Transfer to MPS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8595
         TabIndex        =   10
         Top             =   975
         Width           =   1920
      End
      Begin VB.CommandButton cmdlinkSales 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10185
         Picture         =   "FrmSalesForecast.frx":686B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   158
         Width           =   330
      End
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   10185
         Picture         =   "FrmSalesForecast.frx":6BF5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   518
         Width           =   330
      End
      Begin VB.ComboBox CmbPeriod 
         Appearance      =   0  'Flat
         DataField       =   "periode_type"
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
         Height          =   315
         ItemData        =   "FrmSalesForecast.frx":6F7F
         Left            =   3030
         List            =   "FrmSalesForecast.frx":6F8C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "PO"
         Top             =   870
         Width           =   1905
      End
      Begin MSDataGridLib.DataGrid DGPurchase 
         Height          =   4335
         Left            =   120
         TabIndex        =   24
         Tag             =   "Partner"
         Top             =   2400
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "customer"
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
         Left            =   6720
         TabIndex        =   12
         Tag             =   "PO"
         Top             =   510
         Width           =   3465
      End
      Begin VB.Label lblcustomerID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "customer"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9120
         TabIndex        =   23
         Tag             =   "PO"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblsalesID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "periode_no"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9120
         TabIndex        =   22
         Tag             =   "PO"
         Top             =   840
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   5370
         X2              =   6870
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Customer"
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
         Index           =   14
         Left            =   5385
         TabIndex        =   20
         Top             =   585
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   135
         X2              =   1560
         Y1              =   1899
         Y2              =   1899
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5370
         X2              =   6795
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   135
         X2              =   1560
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   120
         X2              =   1545
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   135
         X2              =   1560
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   135
         X2              =   1560
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Mulai"
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
         Left            =   150
         TabIndex        =   19
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Sales"
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
         Left            =   5385
         TabIndex        =   18
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forecast Name"
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
         Left            =   150
         TabIndex        =   17
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Akhir"
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
         Left            =   150
         TabIndex        =   16
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
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
         Index           =   4
         Left            =   150
         TabIndex        =   15
         Top             =   945
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Index           =   9
         Left            =   150
         TabIndex        =   14
         Top             =   585
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue To"
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
         Left            =   150
         TabIndex        =   13
         Top             =   2010
         Width           =   615
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   135
         X2              =   1635
         Y1              =   2250
         Y2              =   2250
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmSalesForecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcJadwal As New Recordset
Private RcBacaMPS As New Recordset
Private RcBacaFC As New Recordset
Private clsMytr As New DBQuick
Private RcUang As New DBQuick
Private RcDetail As New DBQuick
Private RcDetailTest As New Recordset
Private RcPartner As New DBQuick
Private mAwal As Integer
Private mAkhir As Integer
Private mCount As Integer
Private mList As String
Private mRowLast As Long
Private Rctest As New Recordset
Private RcTestIsi As New Recordset
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Private MyData As New clsTransaksi
Dim strSQL As String
Dim RsDetail As DBQuick
Dim ChekBaris, ChekBarisDel As Integer
Private MEdit, mEditPO, mFirstCaller As Boolean

Private Sub CmbPeriod_Click()
'If CmbPeriod = "MONTHLY" Then
'    txtperiode(0).Text = "12"
'ElseIf CmbPeriod = "WEEK" Then
'    txtperiode(0).Text = "48"
'Else
'   txtperiode(0).Text = "365"
'End If
' GenerateJadwal 0, txtperiode(0)
End Sub

Private Sub cmdLink_Click(Index As Integer)
OpenPartner 0
End Sub

Private Sub cmdlinkSales_Click()
TSalesforcast = True
frmCallerBaru.Show 1
End Sub

Private Sub Command1_Click()
Dim Rctest As New Recordset
Dim RcTestIsi As New Recordset
Dim I, j, h As Integer

Set Rctest = New Recordset
GenerateJadwal 0, 12

Rctest.Open "SELECT  dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.fcast_item " & _
            "FROM  dbo.sales_fcast INNER JOIN " & _
            "dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id INNER JOIN " & _
            "dbo.sales_fcast_item ON dbo.sales_fcast_line.fcast_item = dbo.sales_fcast_item.fcast_item " & _
            "GROUP BY dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.fcast_item", CNN, adLockReadOnly


RcTestIsi.Open "SELECT  dbo.sales_fcast_line.fcast_item, dbo.sales_fcast_line.list_value1, dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.no_urut " & _
               "FROM    dbo.sales_fcast INNER JOIN " & _
               "dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id INNER JOIN " & _
               "dbo.sales_fcast_item ON dbo.sales_fcast_line.fcast_item = dbo.sales_fcast_item.fcast_item " & _
               "GROUP BY dbo.sales_fcast_line.fcast_item, dbo.sales_fcast_line.list_value1, dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.no_urut", CNN, adLockReadOnly
h = 0
Do While h <= Rctest.Recordcount - 1   'baris
    MyDDE.ChildRecordset.AddNew
h = h + 1
Loop

With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           I = 1 ' colom
           Do While I <= FQty(txtperiode(0).Text) + 1 'kolom
              DGPurchase.col = I
              j = 0
              Rctest.MoveFirst
              Do While j <= Rctest.Recordcount - 1 'DGPurchase.Row '2 baris
                 DGPurchase.col = I
                 DGPurchase.row = j '
                 MyDDE.ChildRecordset("item no") = Rctest.Fields("item_no")
                 MyDDE.ChildRecordset("fcast item") = Rctest.Fields("fcast_item")
                 j = j + 1
                 Rctest.MoveNext
              Loop
            I = I + 1
          Loop
    End If
End With


'digunakan untuk nampilkan data di ambil dari tabel

'With MyDDE.ChildRecordset
'     If RcTestIsi.Recordcount <> 0 Then
'        i = 1
'        Do While i <= 4 'FQty(txtperiode(0).text) + 1
'           DGPurchase.Col = i
'           j = 0
'           RcTestIsi.MoveFirst
'           Do While j <= 5 And RcTestIsi.EOF = False
'              DGPurchase.Row = j '
'              DGPurchase.Col = i
'              If DGPurchase.Columns(0).value = RcTestIsi.Fields("item_no") And Trim(DGPurchase.Columns(1).value) = Trim(RcTestIsi.Fields("Fcast_item")) And DGPurchase.Columns(i).text = RcTestIsi.Fields("no_urut") Then
'                 MyDDE.ChildRecordset(i) = IIf(IsNull(RcTestIsi.Fields("list_value1") = True), 0, RcTestIsi.Fields("list_value1"))
'                 j = j + 1
'                 List2.AddItem j & "     " & i
'              End If
'              RcTestIsi.MoveNext
'            Loop
'            i = i + 1
'            MyDDE.ChildRecordset.MoveNext
'        Loop
'    End If
'End With


'digunakan isi data di grid
With MyDDE.ChildRecordset
     If RcTestIsi.Recordcount <> 0 Then
        I = 0
        Do While I <= Rctest.Recordcount - 1 'FQty(txtperiode(0).text) + 1
           DGPurchase.col = I
           j = 1
           Do While j <= FQty(txtperiode(0).Text) + 1
              RcTestIsi.MoveFirst
              Do While RcTestIsi.EOF = False
                 DGPurchase.row = I '
                 DGPurchase.col = j
                 If DGPurchase.Columns(0).Value = RcTestIsi.Fields("item_no") And Trim(DGPurchase.Columns(1).Value) = Trim(RcTestIsi.Fields("Fcast_item")) And DGPurchase.Columns(j).Caption = RcTestIsi.Fields("no_urut") Then
                    MyDDE.ChildRecordset(j) = IIf(IsNull(RcTestIsi.Fields("list_value1") = True), 0, RcTestIsi.Fields("list_value1"))
                 End If
                 RcTestIsi.MoveNext
              Loop
              j = j + 1
            Loop
            I = I + 1
            MyDDE.ChildRecordset.MoveNext
        Loop
    End If
End With
End Sub

Private Sub CmdTransfer_Click()
Dim I, j As Integer
I = MessageBox("Anda Yakin Transfer Data", "Peringatan", msgYesNo, msgQuestion)
If I = 1 Then
    j = 0
    Do While j <= ChekBaris 'DGPurchase.Row '2
        DGPurchase.row = j
        Insert_Trasfer_MPS DGPurchase.Columns(0), CmbPeriod.Text, FCastDate(0).Value, FCastDate(1).Value
        j = j + 1
    Loop
    MessageBox "Data Transfer Success", "Pesan", msgOkOnly, msgInfo
End If
End Sub

Private Sub DGPurchase_Click()
txtIsi.Visible = False
End Sub

Private Sub DGPurchase_DblClick()
'On Error Resume Next
'If DGPurchase.Col >= 2 Then
'    txtIsi.Move DGPurchase.Columns(DGPurchase.Col).Left + 120, DGPurchase.RowTop(DGPurchase.Row) + 2400, DGPurchase.Columns(DGPurchase.Col).Width, DGPurchase.RowHeight
'    txtIsi.Visible = True
'    txtIsi.Text = MyDDE.ChildRecordset(DGPurchase.Col)
'    txtIsi.SetFocus
'End If
End Sub

Private Sub DGPurchase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGPurchase.col >= 2 Then
   DGPurchase.AllowUpdate = True
Else
   DGPurchase.AllowUpdate = False
End If
End Sub

Private Sub FCastDate_Change(Index As Integer)
If Index = 1 Then
   mList = CmbPeriod.Text
   txtperiode(0).Text = HitungHari
   GenerateJadwal 0, txtperiode(0)
End If
End Sub

Private Sub FCastDate_Click(Index As Integer)
If Index = 1 Then
   mList = CmbPeriod.Text
   txtperiode(0).Text = HitungHari
   GenerateJadwal 0, txtperiode(0)
End If

End Sub

Private Sub Form_Activate()
'digunakan membaca awal banyaknya baris di grid
If MyDDE.ChildRecordset.Recordcount > 0 Then
    MyDDE.ChildRecordset.MoveLast
    ChekBaris = DGPurchase.row
End If
End Sub

Private Sub Form_Load()
'GridLayout
HiasFormManTell Picture2, Me
Set mCall = New frmCaller
FCastDate(0).Value = dDateBegin
FCastDate(0).Enabled = False
FCastDate(1).Value = dDateBegin
FCastDate(1).Enabled = False
With MyDDE
     .EditModeReplace = False
      Set .BindForm = FrmSalesForecast
      Set .ActiveConnection = CNN
     .BindFormTAG = "PO"
     .PrepareQuery = "SELECT * from sales_fcast"
End With
GenerateJadwal 0, 0
Load_FCast_Item
list_dataGrid MyDDE.GetFieldByName("fcast_id") 'digunakan untuk tampilkan data di grid

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set MyData = Nothing
MyDDE.ClearRecordset
'RcUang.CloseDB
'clsMytr.CloseDB
'RcPartner.CloseDB
Set mCall = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSalesForecast = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Dim I As Integer
Select Case TagForm:
       Case "MASTER CUSTOMER":
            MyDDE.GetFieldByName("customer") = mCall.GetFieldByName(0)
            lblBank(1) = mCall.GetFieldByName(1)
       Case "MASTER BARANG":
             For I = 0 To 2   'digunakan input list 3
               ' If I = 0 Then
                   MyDDE.ChildRecordset("Item No") = mCall.GetFieldByName(0)
                   MyDDE.ChildRecordset.AddNew
               ' Else
                 '   MyDDE.ChildRecordset.AddNew
               ' End If
             Next I
             MyDDE.ChildRecordset.Delete
             insert_Item_Fcast 'digunakan insert Sales_Fcast_Item yang di ambil dari tabel
             MyDDE.ChildRecordset.MoveLast
             ChekBaris = DGPurchase.row  'digunakan chek baris terakhir
End Select
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
txtperiode(0).Enabled = False
Select Case AdReasonActiveDb
       Case tmbEdit:
             MEdit = True
             MyDDE.ChildRecordset.MoveLast
             FCastDate(0).Enabled = True
             FCastDate(1).Enabled = True
       Case tmbAddNew:
             MEdit = False
             cmdlinkSales.Enabled = True
             cmdLink(0).Enabled = True
             clear_grid  'di gunakan clear grid untuk tambah data
             FCastDate(0).Enabled = True
             FCastDate(1).Enabled = True
       Case tmbSave:
          If MyDDE.IsChildMemberReady = True Then
               SimpanDetail True
               MEdit = False
               mEditPO = False
            End If
            FCastDate(0).Enabled = False
            FCastDate(1).Enabled = False
       Case tmbCancel:
            If MyDDE.ChildRecordset.Recordcount = 0 Then
               list_dataGrid MyDDE.GetFieldByName("Fcast_id")
               MEdit = False
            Else
               MEdit = True
            End If
            FCastDate(0).Enabled = False
            FCastDate(1).Enabled = False
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               OpenPartner 1
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbNextRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbPreviousRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbBottomRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbTopRecord:
            list_dataGrid txtBox(0)
            Form_Activate
       Case tmbPrint:
          '  CallRPTReport "Sales Contract.rpt", "Select * From [Sales Contract] where PurchaseID ='" & txtBox(0) & "'"
       Case tmbQuit:
'            Unload Me

End Select
GridLayout
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO  sales_fcast ( fcast_id, description, periode_no, Periode_type, date_start, date_stop, issue_to, sales_pic, Customer) " & _
                     " VALUES (N'" & txtBox(0) & "', N'" & txtBox(1) & "'," & txtperiode(0).Text & ",'" & CmbPeriod.Text & "',convert(Datetime, '" & Format(FCastDate(0).Value, "dd/mm/yy") & "',3),convert(Datetime,'" & Format(FCastDate(1).Value, "dd/mm/yy") & "',3),'" & Combo1 & "','" & txtBox(2) & "','" & lblcustomerID.Caption & "')"
    
    .PrepareUpdate = "Update sales_fcast" & _
                     " set description='" & txtBox(1) & "', periode_no=" & txtperiode(0).Text & ",periode_type='" & CmbPeriod.Text & "' , date_start=convert(Datetime, '" & Format(FCastDate(0).Value, "dd/mm/yy") & "',3),date_stop=convert(Datetime,'" & Format(FCastDate(1).Value, "dd/mm/yy") & "',3),issue_to='" & Combo1.Text & "',sales_pic='" & txtBox(2) & "'," & _
                     " customer='" & lblcustomerID.Caption & "' where fcast_id='" & MyDDE.GetFieldByName("fcast_id") & "'"
    .PrepareDelete = "delete from sales_fcast where fcast_id = '" & MyDDE.GetFieldByName("fcast_id") & "'"
End With
End Sub


Private Sub OpenPartner(ByVal Index As Integer)
On Error GoTo Hell:
Set RcPartner = New DBQuick
Select Case Index
       Case 0:
            RcPartner.DBOpen " SELECT PartnerID AS [Partner ID],CompanyName as Perusahaan, Address AS Alamat, City AS Kota, PostalCode AS [Kode Pos], Country AS Negara, Phone AS Telp FROM PartnerDB WHERE (PartnerType = N'CUSTOMER') ORDER BY PartnerID", CNN, lckLockReadOnly
       Case 1:
             RcPartner.DBOpen "SELECT NoItem AS [BOM Id], ItemName AS Keterangan, UOM AS UOM FROM Inventory WHERE     (Manufacture = 1) ORDER BY NoItem", CNN, lckLockReadOnly
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case Index
          Case 0:
            mCall.FromTagActive = "MASTER CUSTOMER"
            mCall.CaptionLink = "Customer"
          Case 1:
            mCall.FromTagActive = "MASTER BARANG"
            mCall.CaptionLink = "Barang"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
End If
Exit Sub
Hell:
    Err.Clear
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
PrepareQuery
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Set rsDetail = New DBQuick
''rsDetail.DBOpen "select   timestamp, fcast_id, item_no as [No Item], fcast_item as [FCast Item], time_days, list_value1, no_urut from sales_fcast_line where fcast_id = '" & MyDDE.GetFieldByName("fcast_id") & "'", CNN
'rsDetail.DBOpen "select * from sales_fcast_line where fcast_id = '" & MyDDE.GetFieldByName("fcast_id") & "'", CNN
'Set MyDDE.ChildRecordset = rsDetail.DBRecordset
'Set DataGrid1.DataSource = MyDDE.ChildRecordset

'OpenCucakRowo IIf(Not IsNull(MyDDE.GetFieldByName("NoItem")), MyDDE.GetFieldByName("NoItem"), "xxx"), True
'GenerateJadwal 0, txtperiode(0)

FCastDate(0) = IIf(Not IsNull(MyDDE.GetFieldByName("date_start")), MyDDE.GetFieldByName("date_start"), Date)
FCastDate(1) = IIf(Not IsNull(MyDDE.GetFieldByName("date_stop")), MyDDE.GetFieldByName("date_stop"), Date)

End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
     Case tmbEdit:
     Case tmbDelete:
          MyDDE.ChildRecordset.CancelBatch
     Case tmbDetail: '
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
    Case tmbCancel:
        MyDDE.ChildRecordset.CancelBatch
End Select
End Sub

Private Sub SimpanDetail(ByVal Tipical As Boolean)
On Error Resume Next
Dim I, j, ntime As Integer


If MEdit = True Then
With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           I = 1 ' colom
           Do While I <= FQty(txtperiode(0).Text) + 1
              DGPurchase.col = I
              j = 0
              Do While j <= ChekBaris 'DGPurchase.Row '2
                DGPurchase.col = I
                DGPurchase.row = j
                SendDataToServer "DELETE FROM sales_fcast_line WHERE     (fcast_id = N'" & txtBox(0) & "')"

                j = j + 1
              Loop
            I = I + 1
          Loop
        '  End If
    End If
End With
End If


With MyDDE.ChildRecordset
     If .Recordcount <> 0 Then
           I = 1 ' colom
           Do While I <= FQty(txtperiode(0).Text) + 1
              DGPurchase.col = I
              j = 0
              Do While j <= ChekBaris 'DGPurchase.Row '2
                DGPurchase.col = I
                DGPurchase.row = j
                ntime = I
                SendDataToServer " INSERT INTO sales_fcast_line ( fcast_id, item_no, fcast_item,no_urut,list_value1,time_days)" & _
                                 " VALUES (N'" & txtBox(0) & "','" & DGPurchase.Columns(0).Value & "' ,'" & DGPurchase.Columns(1).Value & "'," & j + 1 & " ," & FQty(DGPurchase.Columns(1 + I).Value) & "," & ntime & "  )"

                j = j + 1
              Loop
            I = I + 1
          Loop
    End If
End With
End Sub

Private Sub GridLayout()
Dim I As Integer
If DGPurchase.Splits.Count < 2 Then DGPurchase.Splits.Add 1
DGPurchase.Splits(0).Columns(0).Visible = True
DGPurchase.Splits(0).Columns(1).Visible = True
DGPurchase.Splits(0).Columns(1).width = 1750
DGPurchase.Splits(1).Columns(0).Visible = False
DGPurchase.Splits(1).Columns(1).Visible = False
DGPurchase.Splits(0).ScrollBars = dbgHorizontal
DGPurchase.Splits(1).ScrollBars = dbgBoth
'DGPurchase.Splits(1).RecordSelectors = False
DGPurchase.Splits(0).AllowSizing = False
DGPurchase.Splits(1).Size = 2
DGPurchase.Splits(1).SizeMode = dbgScalable
'DGPurchase.Splits(1).LeftCol = 0
'DGPurchase.Splits(0).Columns(0).Caption = "No Item"
For I = 0 To RcJadwal.Fields.Count - 1
    If I <= 1 Then
       DGPurchase.Columns(I).width = 1700
       DGPurchase.Columns(I).Alignment = dbgLeft
    Else
       DGPurchase.Columns(I).width = 600
       DGPurchase.Columns(I).Alignment = dbgRight
       DGPurchase.Columns(I).NumberFormat = "#,##0;(#,##0)"
       DGPurchase.Splits(0).Columns(I).Visible = False
       DGPurchase.Splits(1).Columns(I).AllowSizing = True
    End If
    DGPurchase.Columns(I).DividerStyle = dbgRaised
Next
DGPurchase.TabAcrossSplits = True
DGPurchase.Refresh

'untuk split ke 1
'DGPurchase.Splits(0).Columns(0).Visible = False
'DGPurchase.Splits(0).Columns(1).Visible = False
'DGPurchase.Splits(0).Columns(4).Visible = False
'DGPurchase.Splits(0).Columns(5).Visible = False
'DGPurchase.Splits(0).Columns(6).Visible = False
'
''Untuk split ke 2
'DGPurchase.Splits(1).Columns(0).Visible = False
'DGPurchase.Splits(1).Columns(1).Visible = False
'DGPurchase.Splits(1).Columns(2).Visible = False
'DGPurchase.Splits(1).Columns(3).Visible = False
'DGPurchase.Splits(1).Columns(4).Visible = False
'DGPurchase.Splits(1).Columns(5).Visible = False
'DGPurchase.Splits(1).Columns(6).Visible = False
End Sub


Private Sub GenerateJadwal(ByVal vAwal As Integer, ByVal vAkhir As Integer)
On Error Resume Next
Dim I As Integer
Set RcJadwal = Nothing
Set RcJadwal = New Recordset
With RcJadwal
     .Fields.Append "Item No", adBSTR
     .Fields.Append "FCast Item", adBSTR
     '.Fields.Append "Stock", adInteger
     For I = 1 To vAkhir
        .Fields.Append vAwal + I, adInteger
     Next
End With
RcJadwal.Open
Set DGPurchase.DataSource = RcJadwal
mAwal = vAwal
mAkhir = RcJadwal.Fields.Count
mCount = vAkhir
GridLayout
Set MyDDE.ChildRecordset = RcJadwal
Err.Clear
End Sub

Private Sub OpenCucakRowo(ByVal vKode As String, Optional ByVal Tipical As Boolean = False)
On Error Resume Next
Dim Rc As New DBQuick
Dim RcDetail As New Recordset
Dim vWeek As Integer
Dim I As Integer
Dim iJ As Integer
Dim mLast As Integer
Dim Avdata As Variant
'Dim mStart As Boolean
mRowLast = 0
'mTotal = CDbl(txtMrp(6))
vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
mList = IIf(Not IsNull(MyDDE.GetFieldByName("Plan Horizon")), MyDDE.GetFieldByName("Plan Horizon"), "Week")
Select Case MyDDE.GetFieldByName("Plan Horizon")
       Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), HitungHari
       Case "Week": GenerateJadwal vWeek, HitungHari
'       Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
       Case Else: GenerateJadwal vWeek, HitungHari
End Select
'Tolong Jarno ben sak garis wae, debug-e dhek sql cek enak....Please.
If Tipical = False Then
  ' Rc.DBOpen " SELECT [BOM Component Detail].Component, Inventory.ItemName, Inventory.LeadTimeDays AS LeadTimeDays,  ISNULL([Inventory Tabel].QTY_IN - [Inventory Tabel].QTY_OUT, 0) AS QTY, [BOM Component Detail].QTYUsage FROM  [BOM Component Detail] INNER JOIN  Inventory ON [BOM Component Detail].Component = Inventory.NoItem LEFT OUTER JOIN                       [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem WHERE  ([BOM Component Detail].NoItem = N'" & vKode & "') ORDER BY [BOM Component Detail].Component", CNN, lckLockReadOnly
   With Rc.DBRecordset
        If .Recordcount <> 0 Then
        'MsgBox .Source
           'LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), MyDDE.GetFieldByName("QTY") - MyDDE.GetFieldByName("Unit Aloc"), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
           Do
              If .EOF = True Then Exit Do
                ' LoadJadwal .Fields(0), .Fields(2), .Fields(3), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True '.Fields(3)
                 mRowLast = mRowLast + 6
                .MoveNext
           Loop
        End If
   End With
Else
 ' Rc.DBOpen "Shape{SELECT [MRP Detail].Component, [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date] FROM [MRP INVENTORY] INNER JOIN [MRP Detail] ON [MRP INVENTORY].NoItem = [MRP Detail].NoItem WHERE ([MRP INVENTORY].NoItem = N'" & vKode & "') GROUP BY [MRP INVENTORY].[Require Date], [MRP INVENTORY].[End Date], [MRP Detail].Component ORDER BY [MRP Detail].Component} Append({SELECT [MRP Detail].Component AS Component, [MRP Detail].[Time Days] AS [Plan Horizon], [MRP Detail].[List Value1] AS Amount, [MRP Detail].[No Urut] FROM [MRP Detail] INNER JOIN Inventory ON [MRP Detail].NoItem = Inventory.NoItem WHERE     (Inventory.Manufacture = 1) AND ([MRP Detail].NoItem = N'" & vKode & "') GROUP BY [MRP Detail].[Time Days], [MRP Detail].[List Value1], [MRP Detail].[No Urut], [MRP Detail].Component ORDER BY [MRP Detail].Component, [MRP Detail].[No Urut]} As ChildMD Relate Component To Component)", CNN, lckLockBatch
  With Rc.DBRecordset
       If .Recordcount <> 0 Then
'          vWeek = Format(CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)), "ww")
'          mList = MyDDE.GetFieldByName("Plan Horizon")
'          Select Case MyDDE.GetFieldByName("Plan Horizon")
'                 Case "Day": GenerateJadwal (Day(MyDDE.GetFieldByName("Require Date"))), HitungHari
'                 Case "Week": GenerateJadwal vWeek, HitungHari
''                 Case "Monthly": GenerateJadwal MonthOfYear(DTPicker1(0)), HitungHari
'                 Case Else: GenerateJadwal vWeek, HitungHari
'          End Select
          Set RcDetail = Rc.DBRecordset("ChildMD").UnderlyingValue
'            RcDetail.MoveFirst
'            MsgBox RcDetail.GetString(adClipString)
'             LoadJadwal MyDDE.GetFieldByName("NoItem"), MyDDE.GetFieldByName("Lead Time"), MyDDE.GetFieldByName("QTY") - MyDDE.GetFieldByName("Unit Alloc"), IIf(Not IsNull(MyDDE.GetFieldByName("Safety Stock")), MyDDE.GetFieldByName("Safety Stock"), 0), True
             Do
               If .EOF Then Exit Do
                  mLast = 1
                  iJ = 0
                  If RcDetail.Recordcount <> 0 Then
                        Avdata = RcDetail.Getrows(RcDetail.Recordcount, adBookmarkFirst)
                        For I = 0 To UBound(Avdata, 2)
                            iJ = iJ + 1
                            If mLast <> Avdata(3, I) Then iJ = 1
                            Select Case Avdata(3, I)
                                   Case 1:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(0) = .Fields(0): RcJadwal.Fields(1) = "Gross Requirement"
                                   Case 2:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Schedule Receipt"
                                   Case 3:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "On Hand"
                                   Case 4:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Net Requirement"
                                   Case 5:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Plan Order Receipt"
                                   Case 6:
                                        If iJ = 1 Then RcJadwal.AddNew: RcJadwal.Fields(1) = "Plan Order Release"
                            End Select
'                            MsgBox RcJadwal.Fields(10).Name
                            RcJadwal.Fields((2 + iJ) - 1) = Avdata(2, I)
                            'RcJadwal.Fields(iJ - 1) = Avdata(2, I)
                            mLast = Avdata(3, I)
                        Next I
                  End If
               .MoveNext
             Loop
             .MoveFirst
       Else
           
       End If
  End With
End If
End Sub


Private Function HitungHari() As Long
Dim mTotal As Long
'mTotal = CDate(IIf(Not IsNull(MyDDE.GetFieldByName("Require Date")), MyDDE.GetFieldByName("Require Date"), Date)) - CDate(IIf(Not IsNull(MyDDE.GetFieldByName("End Date")), MyDDE.GetFieldByName("End Date"), Date))
'mTotal = CDate(IIf(Not IsNull(FCastDate(0).value), FCastDate(0).value, Date)) - CDate(IIf(Not IsNull(FCastDate(1).value), FCastDate(1).value, Date))
mTotal = CDate(IIf(Not IsNull(FCastDate(1).Value), FCastDate(1).Value, Date)) - CDate(IIf(Not IsNull(FCastDate(0).Value), FCastDate(0).Value, Date))
Select Case mList
       Case "DAY": HitungHari = mTotal
       Case "WEEK": HitungHari = Round(mTotal / 7)
       Case "MONTHLY": HitungHari = Round(mTotal / 30)
End Select
End Function

Private Sub insert_FCast_Item()
Dim I As Integer
Set RcDetailTest = New ADODB.Recordset
RcDetailTest.Open "select * from sales_fcast_item", CNN
RcDetailTest.MoveFirst
If RcDetailTest.Recordcount > 0 Then
    I = 0
    Do While RcDetailTest.EOF = False
        DGPurchase.row = I
        MyDDE.ChildRecordset(1) = RcDetailTest.Fields("fcast_item")
        RcDetailTest.MoveNext
        I = I + 1
    Loop
End If
End Sub

Private Sub Load_FCast_Item()
Dim I As Integer
Set RcDetailTest = New ADODB.Recordset
RcDetailTest.Open "select * from sales_fcast_item", CNN
RcDetailTest.MoveFirst
If RcDetailTest.Recordcount > 0 Then
    Do While RcDetailTest.EOF = False
        List1.AddItem RcDetailTest.Fields("fcast_item")
        RcDetailTest.MoveNext
    Loop
End If
End Sub

Private Sub txtIsi_Change()
'MyDDE.ChildRecordset(DGPurchase.Col) = txtIsi.Text
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   MyDDE.ChildRecordset(DGPurchase.col) = txtIsi.Text
   txtIsi.Visible = False
   txtIsi.Text = ""
End If
End Sub

Private Sub insert_Item_Fcast()
Dim I, j As Integer
 j = 0
 For I = 0 To DGPurchase.row
    DGPurchase.row = I
    MyDDE.ChildRecordset(1) = List1.List(j)
    j = j + 1
    If j = 3 Then
       j = 0
    End If
 Next I
End Sub

Private Sub list_dataGrid(idfcast As String)
Dim I, j, h As Integer

Set Rctest = New Recordset
Set RcTestIsi = New Recordset

If txtperiode(0).Text <> "" Then
    GenerateJadwal 0, FQty(txtperiode(0).Text)
    
    Rctest.Open "SELECT  dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.fcast_item " & _
                "FROM  dbo.sales_fcast INNER JOIN " & _
                "dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id INNER JOIN " & _
                "dbo.sales_fcast_item ON dbo.sales_fcast_line.fcast_item = dbo.sales_fcast_item.fcast_item " & _
                " where dbo.sales_fcast_line.fcast_id='" & idfcast & "'" & _
                "GROUP BY dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.fcast_item", CNN, adOpenKeyset, adLockReadOnly
                
    
    
    RcTestIsi.Open "SELECT  dbo.sales_fcast_line.fcast_item, dbo.sales_fcast_line.list_value1, dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.no_urut,dbo.sales_fcast_line.time_days " & _
                   "FROM    dbo.sales_fcast INNER JOIN " & _
                   "dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id INNER JOIN " & _
                   "dbo.sales_fcast_item ON dbo.sales_fcast_line.fcast_item = dbo.sales_fcast_item.fcast_item where dbo.sales_fcast_line.fcast_id = '" & idfcast & "'" & _
                   "GROUP BY dbo.sales_fcast_line.fcast_item, dbo.sales_fcast_line.list_value1, dbo.sales_fcast_line.item_no, dbo.sales_fcast_line.no_urut,dbo.sales_fcast_line.time_days", CNN, adOpenKeyset, adLockReadOnly
    h = 0
    Do While h <= Rctest.Recordcount - 1   'baris
        MyDDE.ChildRecordset.AddNew
    h = h + 1
    Loop
    
    With MyDDE.ChildRecordset
         If .Recordcount <> 0 Then
               I = 1 ' colom
               Do While I <= FQty(txtperiode(0).Text) + 1 'kolom
                  DGPurchase.col = I
                  j = 0
                  Rctest.MoveFirst
                  Do While j <= Rctest.Recordcount - 1 'DGPurchase.Row '2 baris
                     DGPurchase.col = I
                     DGPurchase.row = j '
                     MyDDE.ChildRecordset("item no") = Rctest.Fields("item_no")
                     MyDDE.ChildRecordset("fcast item") = Rctest.Fields("fcast_item")
                     j = j + 1
                     Rctest.MoveNext
                  Loop
                I = I + 1
              Loop
        End If
    End With
    
    
    'digunakan isi atau nampilkan data di grid
    With MyDDE.ChildRecordset
         If RcTestIsi.Recordcount <> 0 Then
            I = 0
            Do While I <= Rctest.Recordcount - 1 'FQty(txtperiode(0).text) + 1
              ' DGPurchase.Col = I
               DGPurchase.row = I
               j = 1
               Do While j <= FQty(txtperiode(0).Text) + 1
                  RcTestIsi.MoveFirst
                  Do While RcTestIsi.EOF = False
                     DGPurchase.row = I '
                     DGPurchase.col = j
                     'If DGPurchase.Columns(0).value = RcTestIsi.Fields("item_no") And Trim(DGPurchase.Columns(1).value) = Trim(RcTestIsi.Fields("Fcast_item")) And DGPurchase.Columns(j).Caption = RcTestIsi.Fields("no_urut") Then
                     If DGPurchase.Columns(0).Value = RcTestIsi.Fields("item_no") And Trim(DGPurchase.Columns(1).Value) = Trim(RcTestIsi.Fields("Fcast_item")) And DGPurchase.Columns(j).Caption = RcTestIsi.Fields("time_days") Then
                        MyDDE.ChildRecordset(j) = IIf(IsNull(RcTestIsi.Fields("list_value1") = True), 0, RcTestIsi.Fields("list_value1"))
                     End If
                     RcTestIsi.MoveNext
                  Loop
                  j = j + 1
                Loop
                I = I + 1
                MyDDE.ChildRecordset.MoveNext
            Loop
        End If
    End With
End If

Set Rctest = Nothing
Set RcTestIsi = Nothing
End Sub

Private Sub clear_grid()
Dim I As Integer
MyDDE.ChildRecordset.MoveLast
I = DGPurchase.row
Do While I >= 0 'baris
    DGPurchase.row = I
    MyDDE.ChildRecordset.Delete
    I = I - 1
Loop
End Sub

Private Sub clear_grid_cancel()
Dim I, h As Integer
h = 1
MyDDE.ChildRecordset.MoveLast
I = DGPurchase.row
Do While I >= 0 'baris
    If h <= 3 Then
       DGPurchase.row = I
       MyDDE.ChildRecordset.Delete
    End If
    h = h + 1
    I = I - 1
Loop
End Sub

Private Sub Insert_Trasfer_MPS(ByVal NoItem As String, ByVal Periode As String, ByVal TglStart As String, ByVal tglakhir As String)
On Error Resume Next
'Dim I, j, Ntime As Integer
'SendDataToServer " INSERT INTO  [MPS Header] ( No_MPS, description, periode_no, Periode_type, [Require Date], [end date], issue_to) " & _
'                     " VALUES (N'" & txtBox(0) & "', N'" & txtBox(1) & "'," & txtperiode(0).Text & ",'" & CmbPeriod.Text & "',convert(Datetime, '" & Format(FCastDate(0).Value, "dd/mm/yy") & "',3),convert(Datetime,'" & Format(FCastDate(1).Value, "dd/mm/yy") & "',3),'" & Combo1 & "')"
'With MyDDE.ChildRecordset
'     If .Recordcount <> 0 Then
'           I = 1 ' colom
'           Do While I <= FQty(txtperiode(0).Text) + 1
'              DGPurchase.col = I
'              j = 0
'              Do While j <= ChekBaris 'DGPurchase.Row '2
'                DGPurchase.col = I
'                DGPurchase.row = j
'                Ntime = I
'                SendDataToServer " INSERT INTO [mps detail] (no_mps, noitem, fcast_item,no_urut,list_value1,time_days)" & _
'                                 " VALUES (N'" & txtBox(0) & "','" & DGPurchase.Columns(0).Value & "' ,'" & DGPurchase.Columns(1).Value & "'," & j + 1 & " ," & FQty(DGPurchase.Columns(1 + I).Value) & "," & Ntime & ")"
'
'                j = j + 1
'              Loop
'            I = I + 1
'          Loop
'    End If
'End With


'digunakan membaca MPS
Set RcBacaMPS = New Recordset
strSQL = "SELECT dbo.[MPS Header].Periode_no, dbo.[MPS Header].Periode_type, dbo.[MPS Header].[Require Date], dbo.[MPS Header].[End Date]," & _
               "dbo.[MPS Detail].NoItem , dbo.[MPS Detail].fcast_item, dbo.[MPS Detail].Time_Days, dbo.[MPS Detail].list_value1 " & _
               "FROM dbo.[MPS Detail] INNER JOIN " & _
               "dbo.[MPS Header] ON dbo.[MPS Detail].No_MPS = dbo.[MPS Header].No_MPS " & _
               "WHERE (dbo.[MPS Header].Periode_type = '" & Periode & "') AND (dbo.[MPS Header].[Require Date] >= '" & Format(TglStart, "yyyy-mm-dd") & "') AND (dbo.[MPS Header].[End Date] <= '" & Format(tglakhir, "yyyy-mm-dd") & "')" & _
               " AND (dbo.[MPS Detail].fcast_item = 'FCQTY')" ' AND (dbo.[MPS Detail].NoItem = '" & Noitem & "')"

RcBacaMPS.Open strSQL, CNN, adOpenKeyset, adLockReadOnly



'Di gunakan membaca Sales Forcases
Set RcBacaFC = New Recordset
strSQL = "SELECT dbo.sales_fcast.periode_no, dbo.sales_fcast.periode_type, dbo.sales_fcast.date_start, dbo.sales_fcast.date_stop, dbo.sales_fcast_line.item_no," & _
                " dbo.sales_fcast_line.fcast_item , dbo.sales_fcast_line.time_days, dbo.sales_fcast_line.list_value1 " & _
                " FROM dbo.sales_fcast INNER JOIN " & _
                " dbo.sales_fcast_line ON dbo.sales_fcast.fcast_id = dbo.sales_fcast_line.fcast_id " & _
                " WHERE (dbo.sales_fcast.periode_type = '" & Periode & "') AND (dbo.sales_fcast.date_start >= '" & Format(TglStart, "yyyy-mm-dd") & "') AND (dbo.sales_fcast.date_stop <= '" & Format(tglakhir, "yyyy-mm-dd") & "') AND " & _
                " (dbo.sales_fcast_line.fcast_item = 'FCQTY')" ' AND (dbo.sales_fcast_line.item_no = '" & Noitem & "')"


RcBacaFC.Open strSQL, CNN, adOpenKeyset, adLockReadOnly

'digunakan untuk update di MPS
If RcBacaMPS.Recordcount > 0 Then
   RcBacaMPS.MoveFirst
   Do While RcBacaMPS.EOF <> True
        RcBacaFC.MoveFirst
        Do While RcBacaFC.EOF <> True
            SendDataToServer "Update [MPS Detail] set list_value1 =" & RcBacaFC("list_value1") & "" & _
                             "where noitem='" & NoItem & "' and fcast_item = 'FCQTY' and  time_days =" & RcBacaFC("time_days") & ""
                             
            RcBacaFC.MoveNext
        Loop
    RcBacaMPS.MoveNext
  Loop
Else
    MessageBox "Data Di MPS Tdk Ada", "Pesan", msgOkOnly, msgInfo
End If
End Sub
