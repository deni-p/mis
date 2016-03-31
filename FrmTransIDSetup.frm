VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmTransIDSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Nomor Transaksi"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTransIDSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10575
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5310
      Top             =   2805
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":209DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":2723C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransIDSetup.frx":2DA9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10515
      TabIndex        =   2
      Top             =   5820
      Width           =   10575
      Begin VB.CommandButton Cmd 
         Caption         =   "&Batal"
         Height          =   555
         Index           =   3
         Left            =   780
         Picture         =   "FrmTransIDSetup.frx":34300
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   720
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Edit"
         Height          =   555
         Index           =   2
         Left            =   60
         Picture         =   "FrmTransIDSetup.frx":3AB52
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   720
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Keluar"
         Height          =   555
         Index           =   0
         Left            =   2220
         Picture         =   "FrmTransIDSetup.frx":413A4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   720
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Simpan"
         Height          =   555
         Index           =   1
         Left            =   1500
         Picture         =   "FrmTransIDSetup.frx":42E9E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   0
      ScaleHeight     =   5850
      ScaleWidth      =   10575
      TabIndex        =   3
      Top             =   0
      Width           =   10575
      Begin VB.ListBox ListType 
         Appearance      =   0  'Flat
         Height          =   1395
         ItemData        =   "FrmTransIDSetup.frx":496F0
         Left            =   6480
         List            =   "FrmTransIDSetup.frx":49703
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.TreeView tvConfig 
         Height          =   5475
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   9657
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grid 
         Height          =   5130
         Index           =   0
         Left            =   3330
         TabIndex        =   1
         Top             =   480
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   9049
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "ID"
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
            DataField       =   "IDTrans"
            Caption         =   "IDTrans"
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
            DataField       =   "No Index"
            Caption         =   "No "
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Group Name"
            Caption         =   "Level"
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
            DataField       =   "Length per Account"
            Caption         =   "Length Level"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Prefix"
            Caption         =   "Prefix"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "BitMe"
            Caption         =   "BitMe"
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
         BeginProperty Column07 
            DataField       =   "type"
            Caption         =   "Type"
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
         BeginProperty Column08 
            DataField       =   "Note"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Purchase Order"
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
         Height          =   300
         Left            =   3330
         TabIndex        =   5
         Top             =   120
         Width           =   7125
      End
   End
End
Attribute VB_Name = "FrmTransIDSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LocalError As Boolean
Dim curPos As Integer  'penanda aktif posisi item treeview saat ini
Dim strSQL As String
Dim IDGen As New IDGenerator
Dim rsSet() As DBQuick
Dim SelectNode As Node
Dim myToys As New utility

Private Sub CmdKeluar_Click()
   Unload Me
End Sub


Private Sub cmd_Click(Index As Integer)
Dim x As Integer

Select Case Index
    Case 0  'EXIT
        Unload Me
    Case 1  'SAVE
        If Not SelectNode Is Nothing Then
'            Debug.Print IDGen.TransactionID(SelectNode.Index - 1)
            SaveSetting rsSet(SelectNode.Index - 2), IDGen.TransactionID(SelectNode.Index - 1)
            'clear deleted record
            If SendDataToServer("delete from [Trans Setup] where bitMe=1") Then
                MessageBox "Setup Nomor Transaksi telah tersimpan...", "Setup Transaksi", msgOkOnly, msgInfo
            Else
                MessageBox "Setup Nomor Transaksi gagal disimpan...", "Setup Transaksi", msgOkOnly, msgExclamation
            End If
            
            LockGrid True
            tvConfig.Enabled = True
            tvConfig.SetFocus
            tvConfig.Nodes(curPos).Selected = True
            TVConfig_NodeClick tvConfig.SelectedItem
        Else
            LockGrid True
            tvConfig.Enabled = True
        End If
    Case 2  'EDIT
'        Debug.Print IDGen.TransactionID(SelectNode.Index - 1)
        If Not SelectNode Is Nothing Then
            If myToys.IsFounded("SELECT * From [PO Order] WHERE (TypeTrans = N'" & IDGen.TransactionID(SelectNode.Index - 1) & "')") Then
                MessageBox "Nomor transaksi sedang digunakan...", "Setup Transaksi", msgOkOnly, msgExclamation
            Else
                LockGrid False
                tvConfig.Enabled = False
            End If
        End If
    Case 3  'BATAL
        LockGrid True
        tvConfig.Enabled = True
End Select
End Sub

Private Sub LockGrid(bStatus As Boolean)

grid(0).AllowUpdate = Not bStatus
grid(0).AllowAddNew = Not bStatus
grid(0).AllowDelete = Not bStatus
cmd(2).Enabled = bStatus
cmd(1).Enabled = Not bStatus
cmd(3).Enabled = Not bStatus
End Sub
Private Sub GridLayout()
      grid(0).Columns(0).Visible = False   'ID
      
      grid(0).Columns(1).Visible = False   'TRans ID
      
      grid(0).Columns(2).Caption = "No"    'no Index
      grid(0).Columns(2).width = 500 'no Index
      
      grid(0).Columns(3).Caption = "Level"  ' group Name
      grid(0).Columns(3).width = 1500 'no Index
      
      grid(0).Columns(4).Caption = "Length" ' length per account
      grid(0).Columns(4).width = 700  ' length per account
      
      grid(0).Columns(5).Caption = "Prefix"       ' prefix
      grid(0).Columns(5).width = 500
      
      grid(0).Columns(6).Visible = False           ' bitme
      
      grid(0).Columns(8).width = 2000
End Sub

Private Function CheckNull(ColumnNumber As Integer) As Boolean
CheckNull = True
On Error GoTo xErr
'    Debug.Print grid(0).Columns(ColumnNumber).Value
   If grid(0).Columns(ColumnNumber).Value <> "" Then
      CheckNull = False
   End If
'   CheckNull = False
Exit Function
xErr:
   Err.Clear
End Function


Private Sub SaveSetting(Dataset As DBQuick, Flag As String)
On Error GoTo xErr
   Dim lPrefix As String
   Dim lNote As String
   Dim lType As String
   Set grid(0).DataSource = Dataset.DBRecordset
   Dataset.DBRecordset.MoveFirst
    While Not Dataset.DBRecordset.EOF
'        Debug.Print Dataset.DBRecordset.Fields(5).Name & "-" & Dataset.DBRecordset.Fields(5).Value
        lPrefix = CheckPrefix(Dataset.DBRecordset.Fields("prefix").Value)
'        If CheckNull(5) Then
'            Dataset.DBRecordset.Fields("prefix").Value = " "
'            lPrefix = "Null"
'      Else
'         lPrefix = "'" & Dataset.DBRecordset.Fields("prefix").Value & "'"
'      End If
      
      If CheckNull(7) Then
         MessageBox "Kolom Type tidak boleh kosong, secara Default Akan menjadi Fix Character"
         lType = "Fix Character"
      Else
         lType = grid(0).Columns(7).Value
      End If
      lNote = CheckPrefix(Dataset.DBRecordset.Fields(8).Value)
'      If CheckNull(8) Then
'         lNote = " "
'      Else
'         lNote = grid(0).Columns(8).Value
'      End If
      
      If grid(0).Columns(0).Value = "" Then
         strSQL = "insert into [Trans Setup] (IDTrans,[No Index],[Group Name],[Length per Account],prefix,type,note) values " & _
                                                   "('" & Flag & _
                                                  "'," & grid(0).Columns(2).Value & _
                                                  ",'" & grid(0).Columns(3).Value & _
                                                  "'," & grid(0).Columns(4).Value & _
                                                  "," & lPrefix & _
                                                  ",'" & lType & _
                                                  "'," & lNote & ")"
         
      Else
         strSQL = "update [Trans Setup] set IDTrans='" & Flag & _
                                             "',[no Index]=" & grid(0).Columns(2).Value & _
                                             ", [group Name]='" & grid(0).Columns(3).Value & _
                                             "',[length per Account]=" & grid(0).Columns(4).Value & _
                                             ", prefix =" & lPrefix & _
                                             ", type='" & lType & _
                                             "', note=" & lNote & _
                          " where ID = " & grid(0).Columns(0).Value
      
      End If
'      Debug.Print strSQL
      SendDataToServer strSQL
      Dataset.DBRecordset.MoveNext
   Wend
Exit Sub
xErr:
   If Err.Number = 3021 Then
      Err.Clear
   ElseIf Err.Number = 13 Then
      MessageBox "Semua Data Harus Diisi, tekan spasi untuk mengosongkan data"
      LocalError = True
      Err.Clear
   Else
      MessageBox Err.Number & " : " & Err.Description
   End If
End Sub
Public Function CheckPrefix(vData)
If Len(Trim(vData)) = 0 Then
    CheckPrefix = "NULL"
Else
    If IsNull(vData) Then
        CheckPrefix = "NULL"
    Else
        CheckPrefix = "'" & Replace(vData, "'", "''") & "'"
    End If
End If
End Function
Private Sub LoadTree()
Dim vNode As Node
Dim x As Integer

tvConfig.Nodes.Clear
Set vNode = tvConfig.Nodes.Add(, , "A", "Setup", 4)
vNode.Expanded = True
vNode.Bold = True
For x = 1 To IDGen.TransactionCount
   Set vNode = tvConfig.Nodes.Add("A", tvwChild, IDGen.TransactionID(x), IDGen.TransactionName(x), 6, 5)
Next
tvConfig.Nodes.Item(2).Selected = True
curPos = 2
End Sub

Private Sub LoadList()
   Dim x As Integer
   ListType.Clear
   For x = 1 To IDGen.ItemTypeCount
      ListType.AddItem IDGen.GetItemType(x)
   Next
End Sub

Private Sub Form_Load()
   ReDim rsSet(IDGen.TransactionCount)
   Dim x As Integer
   
   For x = 1 To IDGen.TransactionCount
      Set rsSet(x - 1) = New DBQuick
   Next
   
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   LoadTree
   LoadList
   
   LocalError = False
   strSQL = "select * from [Trans Setup] where IDTrans="
   
   For x = 1 To IDGen.TransactionCount
      rsSet(x - 1).DBOpen strSQL & "'" & IDGen.TransactionID(x) & "'", CNN
   Next
   
   GridLayout
   'load Datagrid
   Set grid(0).DataSource = rsSet(curPos - 2).DBRecordset
   LockGrid True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SendDataToServer "update [trans Setup] set bitMe=0"
End Sub

Private Function GetRS() As ADODB.Recordset
   If tvConfig.SelectedItem.Key = "A" Then
      tvConfig.Nodes.Item(curPos).Selected = True
   End If
   Set GetRS = rsSet(curPos - 2).DBRecordset
End Function

Private Sub grid_AfterDelete(Index As Integer)
   Dim rsMove As ADODB.Recordset
   Dim x As Integer
   Set rsMove = GetRS
   rsMove.MoveFirst
   x = 1
   While Not rsMove.EOF
      grid(0).Columns(2).Value = x
      x = x + 1
      rsMove.MoveNext
   Wend
   rsMove.MoveFirst
   x = 1
   While x < Index
      rsMove.MoveNext
      x = x + 1
   Wend
End Sub

Private Sub grid_AfterInsert(Index As Integer)
   Dim rsMove As ADODB.Recordset
   Dim x As Integer
   Set rsMove = GetRS
   rsMove.MoveFirst
   x = 1
   While Not rsMove.EOF
      grid(0).Columns(2).Value = x
      x = x + 1
      rsMove.MoveNext
   Wend
End Sub

Private Sub grid_BeforeDelete(Index As Integer, Cancel As Integer)
   If grid(Index).Columns(0).Value = "" Then
   Else
      SendDataToServer "update [Trans Setup] set BitMe = 1 where ID=" & grid(Index).Columns(0).Value
   End If
End Sub

Private Sub grid_ButtonClick(Index As Integer, ByVal ColIndex As Integer)
If grid(0).AllowUpdate = True Then
    If ColIndex = -1 Then Exit Sub
    
    If ListType.Visible = True Then
        ListType.Visible = False
    Else
        ListType.Visible = True
        ListType.Move grid(0).Columns(7).Left + 3350, (grid(0).RowTop(grid(0).row) + grid(0).RowHeight + 500), grid(0).Columns(ColIndex).width
        ListType.SetFocus
    End If
Else
    MessageBox "Tekan tombol EDIT untuk melakukan pengeditan", "Setup Transaksi", msgOkOnly, msgExclamation
    '   listx.Clear
    '   Set rsx = IDGen.GetDataLookup
    '   Loop
    '      lisx.AddItem (rsx.Field(0))
    '   Next
End If
End Sub

Private Sub grid_Click(Index As Integer)
   ListType.Visible = False
End Sub

Private Sub grid_GotFocus(Index As Integer)
   ListType.Visible = False
End Sub

Private Sub ListType_DblClick()
   ListType.Visible = False
   grid(0).Columns(7).Value = ListType.List(ListType.ListIndex)
End Sub

Private Sub ListType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then ListType.Visible = False
End Sub

Private Sub TVConfig_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node.Key <> "A" Then
      lbl.Caption = UCase(Node.Text)
      lbl.FontBold = True
      lbl.BackColor = &H0&
      lbl.ForeColor = &HFFFFFF
   End If
   curPos = Node.Index
   Set SelectNode = Node
   If curPos > 1 Then Set grid(0).DataSource = rsSet(curPos - 2).DBRecordset
End Sub
