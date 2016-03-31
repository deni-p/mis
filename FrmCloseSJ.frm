VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F2DD8007-5788-48C8-839C-E57EEDFCBFC6}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCloseSJ 
   Caption         =   "Tutup Surat Jalan"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCloseSJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9945
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
      ForeColor       =   &H80000010&
      Height          =   6195
      Left            =   105
      ScaleHeight     =   6135
      ScaleWidth      =   9600
      TabIndex        =   3
      Top             =   135
      Width           =   9660
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5715
         Left            =   105
         ScaleHeight     =   5715
         ScaleWidth      =   9435
         TabIndex        =   4
         Top             =   285
         Width           =   9435
         Begin MSDataGridLib.DataGrid DgHeader 
            Height          =   5190
            Left            =   105
            TabIndex        =   0
            Tag             =   "SJ"
            Top             =   405
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   9155
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "Tgl DN"
               Caption         =   "Tgl DN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MMM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "DN"
               Caption         =   "DN"
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
               DataField       =   "No Ref"
               Caption         =   "No Ref"
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
               DataField       =   "Expedisi"
               Caption         =   "Expedisi"
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
               DataField       =   "No Pol"
               Caption         =   "No Pol"
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
               DataField       =   "Truck"
               Caption         =   "Truck"
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
            BeginProperty Column06 
               DataField       =   "Status"
               Caption         =   "Status"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Tutup SJ"
                  FalseValue      =   "Open SJ"
                  NullValue       =   "Open SJ"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1590.236
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1590.236
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1590.236
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   5190
            Left            =   4125
            ScaleHeight     =   5160
            ScaleWidth      =   5145
            TabIndex        =   7
            Top             =   390
            Width           =   5175
            Begin MSDataGridLib.DataGrid DgItem 
               Height          =   1350
               Left            =   105
               TabIndex        =   1
               Top             =   1860
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   2381
               _Version        =   393216
               AllowUpdate     =   0   'False
               Appearance      =   0
               HeadLines       =   1
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
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "NoItem"
                  Caption         =   "No Item"
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
                  DataField       =   "ItemName"
                  Caption         =   "Nama Item"
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
                  DataField       =   "QTY MASUK"
                  Caption         =   "QTY. Kirim"
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
                  DataField       =   "ActQty"
                  Caption         =   "Aktual QTY"
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
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   2819.906
                  EndProperty
                  BeginProperty Column02 
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
               EndProperty
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   105
               X2              =   1560
               Y1              =   3915
               Y2              =   3915
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   105
               X2              =   1560
               Y1              =   3540
               Y2              =   3540
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   105
               X2              =   1560
               Y1              =   1410
               Y2              =   1410
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   120
               X2              =   1575
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   105
               X2              =   1560
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   90
               X2              =   1545
               Y1              =   375
               Y2              =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status SJ"
               Height          =   210
               Index           =   6
               Left            =   105
               TabIndex        =   20
               Top             =   135
               Width           =   765
            End
            Begin VB.Label lblItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "STATUS"
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Closed"
                  FalseValue      =   "Open"
                  NullValue       =   "Open"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
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
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   5
               Left            =   1365
               TabIndex        =   19
               Tag             =   "SJ"
               Top             =   105
               Width           =   3675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. SJ"
               Height          =   210
               Index           =   5
               Left            =   105
               TabIndex        =   18
               Top             =   480
               Width           =   525
            End
            Begin VB.Label lblItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "DN"
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
               Height          =   285
               Index           =   4
               Left            =   1365
               TabIndex        =   17
               Top             =   450
               Width           =   3675
            End
            Begin VB.Label lblItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "Type Trans"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   3
               Left            =   1365
               TabIndex        =   16
               Top             =   3645
               Width           =   3675
            End
            Begin VB.Label lblItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "EXPEDISI"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   1365
               TabIndex        =   15
               Top             =   3270
               Width           =   3675
            End
            Begin VB.Label lblItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "Perusahaan"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   1365
               TabIndex        =   14
               Top             =   1140
               Width           =   3675
            End
            Begin VB.Label lblItem 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "Gudang"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   1365
               TabIndex        =   13
               Top             =   810
               Width           =   3675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type Trans"
               Height          =   210
               Index           =   4
               Left            =   105
               TabIndex        =   12
               Top             =   3645
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Shipment"
               Height          =   210
               Index           =   3
               Left            =   105
               TabIndex        =   11
               Top             =   3285
               Width           =   780
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   "Daftar Item"
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
               Height          =   300
               Index           =   2
               Left            =   105
               TabIndex        =   10
               Top             =   1530
               Width           =   4935
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer"
               Height          =   210
               Index           =   1
               Left            =   105
               TabIndex        =   9
               Top             =   1200
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gudang"
               Height          =   210
               Index           =   0
               Left            =   105
               TabIndex        =   8
               Top             =   825
               Width           =   630
            End
         End
         Begin VB.OptionButton OptDN 
            BackColor       =   &H80000010&
            Caption         =   "Outgoing"
            Enabled         =   0   'False
            ForeColor       =   &H80000005&
            Height          =   255
            Index           =   1
            Left            =   1785
            TabIndex        =   6
            Top             =   75
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.OptionButton OptDN 
            BackColor       =   &H80000010&
            Caption         =   "Incoming"
            Enabled         =   0   'False
            ForeColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   5
            Top             =   75
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   975
            Top             =   4395
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Bilinus Man;Data Source=BULIRCOMP"
            OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Bilinus Man;Data Source=BULIRCOMP"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"FrmCloseSJ.frx":08CA
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   1217
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmCloseSJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mClose As Boolean
Private RcDetail As Recordset

Private Sub DgHeader_ButtonClick(ByVal ColIndex As Integer)
If DGHeader.Columns(ColIndex).Value = True Then
   DGHeader.Columns(ColIndex).Value = False
Else
   DGHeader.Columns(ColIndex).Value = True
End If
End Sub

Private Sub DgHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGHeader.Col = 6 Then
   DGHeader.MarqueeStyle = dbgFloatingEditor
Else
   DGHeader.MarqueeStyle = dbgHighlightRow
End If
End Sub

Private Sub DgItem_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mClose = True Then
   If DgItem.Col = 3 Then
      DgItem.AllowUpdate = True
   Else
      DgItem.AllowUpdate = False
   End If
Else
   DgItem.AllowUpdate = False
End If
End Sub

Private Sub Form_Activate()
If Me.WindowState = 0 Then Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmCloseSJ
    .BindFormTAG = "SJ"
    Set .ActiveConnection = Cnn
    .PrepareQuery = "SELECT TransData.DateTrans AS [Tgl DN], TransData.TransID AS DN, Transport.Expedisi, TransData.[No Pol], TransData.TypeTruck AS Truck, TransData.StatusInvoice AS STATUS, [PO Order].PurchaseID AS [NO REF] FROM TransData INNER JOIN [PO Order] ON TransData.PurchaseID = [PO Order].PurchaseID INNER JOIN Transport ON TransData.ID = Transport.ID WHERE     (TransData.TypeTrans = N'dn') AND (TransData.Status = 1) AND (TransData.StatusInvoice = 0)"
    .SetPermissions = UserAddnewDeleteDenied
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

CloseDB RcDetail
End Sub

Private Sub Form_Resize()


HiasForm Picture1, Me
CenterForm Picture2, Me
Picture3.BackColor = Picture2.BackColor
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmCloseSJ = Nothing
End Sub


Private Sub Opendetail(ByVal Parameterkode As String)
CloseDB RcDetail
Set RcDetail = New Recordset
RcDetail.CursorLocation = adUseClient
RcDetail.Open " SELECT TransData.TransID AS DN, [Detail TransData].NoItem, WareHouse.[WareHouse Name] AS GUDANG, Inventory.ItemName, PartnerDB.CompanyName AS PERUSAHAAN, PartnerDB.Address AS ALAMAT, PartnerDB.City AS KOTA, TransData.TypeTrans AS [TYPE TRANS],  [Detail TransData].QTY_Receive AS [QTY MASUK], [Detail TransData].QTY_OUT AS [QTY KELUAR], Transport.Expedisi,  [Detail TransData].ACTQTY" & _
              " FROM Inventory INNER JOIN WareHouse ON Inventory.WareHouse = WareHouse.WareHouse INNER JOIN TransData INNER JOIN" & _
              " [Detail TransData] ON TransData.TransID = [Detail TransData].TransID ON Inventory.NoItem = [Detail TransData].NoItem INNER JOIN PartnerDB ON TransData.PartnerId = PartnerDB.PartnerID INNER JOIN Transport ON TransData.ID = Transport.ID WHERE     (TransData.TransID = N'" & Parameterkode & "') ORDER BY [Detail TransData].NoItem", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
Set RcDetail.ActiveConnection = Nothing
With RcDetail
     Set DgItem.DataSource = RcDetail
     Set lblItem(0).DataSource = RcDetail
     Set lblItem(1).DataSource = RcDetail
     Set lblItem(2).DataSource = RcDetail
     Set lblItem(3).DataSource = RcDetail
     Set lblItem(4).DataSource = RcDetail
     'Set lblItem(5).DataSource = RcDetail
End With
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit:
       Case tmbSave:
            'If Check1.Value = 1 Then
               SimpanDetail lblItem(4)
               'MyDDE.RefreshDatabase
            'End If
       Case tmbPrint:
            CallRPTReport "Close SJ.rpt", "Select * from [Close SJ] where DN =N'" & lblItem(4) & "'"
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Opendetail IIf(Not IsNull(MyDDE.GetFieldByName("DN")), MyDDE.GetFieldByName("DN"), "xxx")
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareUpdate = " UPDATE [TRANSDATA]" & _
                     " SET StatusInvoice = StatusInvoice WHERE     (TRANSID = N'" & lblItem(4) & "')"
End With
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbEdit: mClose = True
       Case tmbCancel: mClose = False
       Case tmbSave:
            'If Check1.Value = 1 Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
               mClose = False
'            Else
'               MessageBox ("Kotak Cek Belum Diaktifkan."), "Peringatan", msgOkOnly
'               MyDDE.IsChildMemberReady = False
'            End If
End Select
End Sub

Private Sub SimpanDetail(ByVal ParamString As String)
With RcDetail
     .MoveFirst
     Do
        If .EOF = True Then Exit Do
        SendDataToServer " UPDATE [Detail TRANSDATA] Set ActQTY = " & .Fields("ActQty") & _
                         " WHERE NoItem = N'" & .Fields("Noitem") & "' AND TRANSID = N'" & ParamString & "'"
        .MoveNext
     Loop
     .MoveLast
End With
End Sub
