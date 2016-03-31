VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   AutoRedraw      =   -1  'True
   Caption         =   "Import Dari MS Excel"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   9285
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   6060
      Left            =   0
      ScaleHeight     =   6060
      ScaleWidth      =   9285
      TabIndex        =   5
      Top             =   0
      Width           =   9285
      Begin TabDlg.SSTab SSTab1 
         Height          =   5805
         Left            =   60
         TabIndex        =   6
         Top             =   105
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   10239
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   15380335
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Source Template"
         TabPicture(0)   =   "frmImport.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TreeView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Import Excel"
         TabPicture(1)   =   "frmImport.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DataGrid1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "LstViewData"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5280
            Left            =   -74910
            TabIndex        =   7
            Top             =   390
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   9313
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
                  LCID            =   1033
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
                  LCID            =   1033
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
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5325
            Left            =   75
            TabIndex        =   1
            Top             =   390
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   9393
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin MSComctlLib.ListView LstViewData 
            Height          =   5355
            Left            =   -74940
            TabIndex        =   8
            Top             =   360
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   9446
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6930
      Top             =   6420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   6060
      Width           =   9285
      Begin VB.CommandButton CmdOk 
         Caption         =   "Template"
         Height          =   420
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   1845
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Import Dari Folder"
         Height          =   420
         Index           =   1
         Left            =   2385
         TabIndex        =   3
         Top             =   120
         Width           =   1845
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Transfer To Tabel"
         Height          =   420
         Index           =   2
         Left            =   7230
         TabIndex        =   4
         Top             =   120
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ExcellRc As New Recordset
Private myNode As MSComctlLib.Node
Private mVarLoadFilename As String
Private mVarLokFile As String

Private Sub cmdOk_Click(Index As Integer)
On Error GoTo Hell
Select Case Index
       Case 0: ExportToExcel myNode.Key, myNode.Text
       Case 1:
            With CommonDialog1
               .CancelError = False
               .InitDir = App.Path
               .Filter = "Excell Worksheet (*.xls)|*.xls"
               .ShowOpen
               mVarLoadFilename = .FileTitle
               mVarLokFile = Replace(.Filename, mVarLoadFilename, "")
            End With
            OpenExcel mVarLoadFilename
       Case 2: SaveTable
End Select
Exit Sub
Hell:
    MessageBox Err.Description
    Err.Clear
End Sub

Private Sub Form_Load()
cmdOk(0).Left = 50
cmdOk(0).Top = Me.Height - (cmdOk(0).Height * 2)
cmdOk(1).Top = cmdOk(0).Top
cmdOk(1).Left = cmdOk(0).Left + cmdOk(0).width + 50
cmdOk(2).Top = cmdOk(1).Top
cmdOk(2).Left = cmdOk(1).Left + cmdOk(1).width + 50
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmImport = Nothing
End Sub

Private Sub OpenExcel(ByVal ExcelFileName As String)
If Not ExcellRc Is Nothing Then
   If ExcellRc.State = 1 Then ExcellRc.Close
End If
Set ExcellRc = Nothing
Set ExcellRc = New Recordset
ExcellRc.CursorLocation = adUseClient
ExcellRc.Open " Select * from [" & Replace(UCase(ExcelFileName), ".XLS", "") & "$]", "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DBQ=" & mVarLokFile & mVarLoadFilename & ";DefaultDir=" & mVarLokFile & ";Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=" & App.Path & "Import Tes.dsn;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;ReadOnly=1;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;;Initial Catalog=" & mVarLokFile & Replace(UCase(mVarLoadFilename), ".XLS", "") & ", adOpenForwardOnly, adLockReadOnly, adCmdText"
Set DataGrid1.DataSource = ExcellRc
End Sub

Private Sub SaveTable()
Dim I As Integer
Dim j As Integer
Dim Avadata As Variant
Dim Fld As Field
Dim mVarListAddquery As String
Dim mVarTableName As String
Dim mvarSendDataToServer As String
Dim mVarListValue As String
On Error GoTo xErr
If Not ExcellRc Is Nothing Then
   If ExcellRc.State = 1 Then
      With ExcellRc
           Select Case Replace(UCase(mVarLoadFilename), ".XLS", "")
                  Case "MASTER CURRENCY": mVarTableName = "CURRENCY TABLE"
           End Select
           If .Recordcount <> 0 Then
              Avadata = .Getrows(.Recordcount, adBookmarkFirst)
              mVarListAddquery = ""
              j = 0
              For Each Fld In .Fields
                  If mVarListAddquery = "" Then
                     mVarListAddquery = Trim("[" & Fld.Name & "]")
                  Else
                     mVarListAddquery = Trim(mVarListAddquery & ",[" & Fld.Name & "]")
                  End If
                  j = j + 1
              Next
              For I = 0 To UBound(Avadata, 2)
                  mVarListValue = ""
                  j = 0
                  For Each Fld In .Fields
                      If mVarListValue = "" Then
                         mVarListValue = "'" & Avadata(j, I) & "'"
                      Else
                         mVarListValue = mVarListValue & ",'" & Avadata(j, I) & "'"
                      End If
                      j = j + 1
                  Next
                  mvarSendDataToServer = " INSERT INTO [" & mVarTableName & "] " & _
                                         " (" & mVarListAddquery & ")" & _
                                         " Values (" & mVarListValue & ")"
                  If SendDataToServer(mvarSendDataToServer) = False Then MessageBox "Data Sudah Ada.", "Peringatan"
              Next I
           End If
      End With
   End If
End If
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear

End Sub

Private Sub OpenTable()
With TreeView1.Nodes.Add(, "Master Data", "Master Data", "Master Data")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Master Data", tvwChild, "Currency Table", "Master Currency"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Gudang", "Master Gudang"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Kelompok", "Master Kelompok"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Item Barang", "Master Item Barang"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Regional", "Master Regional"
TreeView1.Nodes.Add "Master Data", tvwChild, "Tipe Pengiriman", "Tipe Pengiriman"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Transporter", "Master Transporter"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Karyawan", "Master Karyawan"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Customer", "Master Customer"
TreeView1.Nodes.Add "Master Data", tvwChild, "Master Supplier", "Master Supplier"
TreeView1.Nodes.Add "Master Data", tvwChild, "Bank Partner", "Bank Partner"

With TreeView1.Nodes.Add(, "Distribusi", "Distribusi", "Distribusi")
     .Bold = True
     .Expanded = True
End With
TreeView1.Nodes.Add "Distribusi", tvwChild, "Order Pembelian", "Order Pembelian"
TreeView1.Nodes.Add "Distribusi", tvwChild, "Penerimaan Barang", "Penerimaan Barang"
TreeView1.Nodes.Add "Distribusi", tvwChild, "Order Penjualan", "Order Penjualan"
TreeView1.Nodes.Add "Distribusi", tvwChild, "Surat Jalan", "Surat Jalan"
TreeView1.Nodes.Add "Distribusi", tvwChild, "Penagihan / Invoice", "Penagihan / Invoice"
TreeView1.Nodes.Add "Distribusi", tvwChild, "Retur Pembelian", "Retur Pembelian"
TreeView1.Nodes.Add "Distribusi", tvwChild, "Retur Penjualan", "Retur Penjualan"
End Sub

Private Function ExportToExcel(ByVal DBTableName As String, ByVal AliasTableName As String) As Recordset
Dim RcEx As New DBQuick
Dim StrQry As String
Select Case DBTableName
       Case "Currency Setup":
            StrQry = "select * from [" & DBTableName & "]"
End Select
If StrQry <> "" Then
   RcEx.DBOpen StrQry, CNN, lckLockReadOnly
   Opendata RcEx.DBRecordset, AliasTableName
   'SaveRecInText RcEx.DBRecordset, App.Path & "\Export Template\" & AliasTableName & ".xls"
End If
End Function

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Set myNode = Node
End Sub

Private Sub Opendata(ByVal Rc As Recordset, ByVal SheetName As String)
Dim ApExcel As Excel.Application
Dim ApSheets As Excel.Workbook
Dim MyCol As String
Dim Response As Integer
Dim MyIndex As Integer
Dim InitRow As Integer
Set ApExcel = CreateObject("Excel.application")
ApExcel.Visible = True
Set ApSheets = ApExcel.Workbooks.Add
'ApSheets.Worksheets(1) = SheetName
For MyIndex = 0 To Rc.Fields.Count - 1
    ApExcel.Cells(1, (MyIndex + 1)).Formula = Rc.Fields(MyIndex).Name
    ApExcel.Cells(1, (MyIndex + 1)).Font.Bold = True
    ApExcel.Cells(1, (MyIndex + 1)).Interior.ColorIndex = 37
    ApExcel.Cells(1, (MyIndex + 1)).WrapText = True
Next
Set ApExcel = Nothing
End Sub
