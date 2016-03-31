VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConfig 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Laporan"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9810
   Begin MSComctlLib.ImageList imgAccount 
      Left            =   4650
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfig.frx":6852
            Key             =   "SEGITIGA"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfig.frx":D0B4
            Key             =   "ABANG"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfig.frx":13916
            Key             =   "BIRU"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfig.frx":1A178
            Key             =   "IJO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   9234
      _Version        =   393217
      Indentation     =   617
      Style           =   7
      ImageList       =   "imgAccount"
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
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9810
      TabIndex        =   8
      Top             =   5805
      Width           =   9810
      Begin VB.CommandButton CmdTombol 
         Caption         =   "New Form"
         Height          =   555
         Index           =   1
         Left            =   150
         Picture         =   "FrmConfig.frx":209DA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   100
         Width           =   945
      End
      Begin VB.Frame FrTombol 
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
         Height          =   30
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   9945
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Edit Filter"
         Height          =   555
         Index           =   2
         Left            =   1095
         Picture         =   "FrmConfig.frx":2722C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   100
         Width           =   945
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Cancel"
         Height          =   555
         Index           =   4
         Left            =   2760
         Picture         =   "FrmConfig.frx":2DA7E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Delete"
         Height          =   555
         Index           =   5
         Left            =   3480
         Picture         =   "FrmConfig.frx":342D0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Save"
         Height          =   555
         Index           =   3
         Left            =   2040
         Picture         =   "FrmConfig.frx":3AB22
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   100
         Width           =   720
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "E&xit"
         Height          =   555
         Index           =   0
         Left            =   4200
         Picture         =   "FrmConfig.frx":41374
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   100
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid GridConfig 
      Bindings        =   "FrmConfig.frx":42E6E
      Height          =   5505
      Left            =   5565
      TabIndex        =   1
      Tag             =   "KP"
      Top             =   135
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9710
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "FORM_NAME"
         Caption         =   "Form Name"
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
         DataField       =   "KODE_FORM"
         Caption         =   "Kode Form"
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
         DataField       =   "FIELD_NAME"
         Caption         =   "FILTER"
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
         DataField       =   "FIELD_TYPE"
         Caption         =   "TIPE"
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
         DataField       =   "OBJECT_TYPE"
         Caption         =   "Object Type"
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
         DataField       =   "idx"
         Caption         =   "idx"
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
            DividerStyle    =   6
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   6
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DAFTAR LAPORAN"
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
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5325
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsConfig As ADODB.Recordset
Dim myPart As New utility
Dim vNode As Node
Dim strSQL As String
Dim SelectNode As Node

Private Sub ConvertFieldType()
  '11  Bit
  '135 DateTime
  '202 nvarchar
  '203 ntext
  '3   int
  '5   Float
  '6   Money

End Sub

Private Sub GridLayout()

  With GridConfig
    .Columns(0).Visible = False
    .Columns(1).Visible = False
    .Columns(5).Visible = False
    .Columns(2).Locked = True
    .Columns(3).Locked = True
    .Columns(4).Locked = True
    .Columns(5).Locked = True
    .Columns(4).Button = True
    .Columns(5).Button = True
    .Columns(3).Alignment = dbgCenter
  End With

End Sub

Private Sub LoadTree()
On Error GoTo xErr


Dim rsForms As DBQuick
Dim rsChild As New Recordset
Dim No  As Integer

TreeView1.Nodes.Clear
Set rsForms = New DBQuick

strSQL = "Shape{SELECT GroupID , GroupName FROM [report group]} as ParentNode append " & _
  " ({SELECT * FROM [Report Modules] ORDER BY ReportGroup, [Alias Report] } as ChildNode relate GroupID to ReportGroup)"
  
rsForms.DBOpen strSQL, CNN, lckLockReadOnly
No = 1

If rsForms.Recordcount > 0 Then
'    rsForms.MoveFirst
    FirstNode = Trim(rsForms.DBRecordset.Fields(0).Value)
    Set rsChild = rsForms.DBRecordset("ChildNode").Value
    With rsForms.DBRecordset
        Do While Not .EOF
            With TreeView1.Nodes.Add(, , .Fields(0).Value, .Fields(1).Value, "BIRU")
                .Bold = True
'                .Expanded = True
            End With
            If rsChild.Recordcount <> 0 Then
                Do While Not rsChild.EOF
                    Set vNode = TreeView1.Nodes.Add(.Fields(0).Value, tvwChild, CStr(rsChild.Fields("NoIdx").Value) & "A", rsChild.Fields("Alias Report").Value, "ABANG")
                    vNode.Tag = rsChild.Fields("ViewObject").Value
                    rsChild.MoveNext
                Loop
            End If
            .MoveNext
        Loop
    End With
    
'    While Not rsForms.EOF
'        Set vNode = TreeView1.Nodes.Add(, , CStr(rsForms.Fields("NoIdx").Value) & "A", Trim(IIf(IsNull(rsForms.Fields( _
'                "Alias Report").Value), " ", rsForms.Fields("Alias Report").Value)))
'        'Set vNode = TVConfig.Nodes.Add("A", tvwChild, "A" & Trim(rsForms.Fields(0).Value), Trim(rsForms.Fields(0).Value))
'        vNode.Tag = No
'        vNode.Expanded = True
'        No = No + 1
'        rsForms.MoveNext
'    Wend
End If

  rsForms.CloseDB
  Set rsForms = Nothing
  
'  Set rsConfig = myPart.OpenDB("SELECT [REPORT_ID], [FIELD_NAME], [FIELD_TYPE] From Report_Filter WHERE ([REPORT_ID]= N'" & _
'          UCase(FirstNode) & "')")
'  Set GridConfig.DataSource = rsConfig
Exit Sub
xErr:
   MessageBox Err.Description, "Error : " & Err.Number, msgOkOnly, msgExclamation
   Err.Clear
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  Set SelectNode = Node
  If SelectNode.Root.Text = "" Then Exit Sub
  SetDataConfig
End Sub
Private Sub SetDataConfig()
Dim idxToStr As String
  
idxToStr = TreeView1.SelectedItem.Key
If Not SelectNode.Parent Is Nothing Then
    Set rsConfig = myPart.OpenDB("SELECT * From Report_Filter WHERE (REPORT_ID = N'" & Replace(idxToStr, "A", "", 1, Len( _
            TreeView1.SelectedItem.Key)) & "')")
    Set GridConfig.DataSource = rsConfig
    FirstNode = Replace(idxToStr, "A", "", 1, Len(TreeView1.SelectedItem.Key))
End If
End Sub
Private Sub CmdTombol_Click(Index As Integer)
  On Error GoTo Hell
Dim NodeIdx As Long

If SelectNode Is Nothing Then
    If Index = 0 Then
        Unload Me
    Else
        Exit Sub
    End If
Else
    NodeIdx = SelectNode.Index
End If
Select Case Index
    Case 0
        Unload Me

    Case 1
        frmEntryForm.OperationMode = "Insert"
        frmEntryForm.FormName = Trim(TreeView1.SelectedItem.Text)
        frmEntryForm.Show vbModal
        LoadTree
        TreeView1.SetFocus
        TreeView1.Nodes(NodeIdx).Selected = True
        TreeView1_NodeClick TreeView1.SelectedItem
        
    Case 2
        If Not SelectNode.Parent Is Nothing Then
            frmEntryForm.OperationMode = "Edit"
            frmEntryForm.FormName = Trim(TreeView1.SelectedItem.Text)
'            frmEntryForm.ViewObject = FirstNode
            frmEntryForm.Show vbModal
            LoadTree
            TreeView1.SetFocus
            TreeView1.Nodes(NodeIdx).Selected = True
            TreeView1_NodeClick TreeView1.SelectedItem
        End If
    
    Case 3
        'SAVE
        If Not rsConfig Is Nothing Then
            rsConfig.UpdateBatch adAffectAllChapters
            MessageBox "Save data successfully...", "Konfigurasi Laporan", msgOkOnly, msgInfo
        End If
    
    Case 5
        If Trim(SelectNode.Root.Text) = "" Then
          MessageBox "Pilih data yang akan dihapus!", "Informasi", msgOkOnly, msgInfo
          Exit Sub
        End If
        
        If MessageBox(TreeView1.SelectedItem.Text & " Yakin Dihapus?", vbOKCancel) = vbOK Then
          SendDataToServer "delete from [Report_Filter] where report_id='" & FirstNode & "'"
          LoadTree
        End If

End Select
Exit Sub
Hell:
    If Err.Number = 35605 Then
        MessageBox "Baris ini sudah dihapus", "Informasi", msgOkOnly, msgExclamation
    Else
        MessageBox Err.Description, "Kontrol Konfigurasi", msgOkOnly, msgExclamation
    End If
  Err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

  If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
'  PassTengah Me, MainMenu
  LoadTree
  'SetDataConfig
  GridLayout
  GridConfig.Columns(4).width = 0
  CenterForm Me, Me
End Sub

Private Sub GridConfig_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo 1
  Select Case ColIndex

    Case 4
      If GridConfig.AllowUpdate = True Then
        FrmLookUp.TitleForm = "Object Type"
        Set FrmLookUp.FormCaller = Me
        Set FrmLookUp.FormContainer = GridConfig.Columns(ColIndex)
        '    Set FrmLookUp.FormContainer2 = GridConfig.Columns(ColIndex + 1)
        
        '   strSQL = "SELECT [Catalogue No], [Engineering Description], [Equipment ID] From [Equipment APLs Table]"
        strSQL = "SELECT Object, [Object Name] From Tools_ObjectView ORDER BY [Object Name]"
        FrmLookUp.SQLScript = strSQL
        FrmLookUp.ColRefNumber = 0
        FrmLookUp.ColRefNumber2 = 0
        FrmLookUp.ColRefNumber3 = 0
        Load FrmLookUp
        FrmLookUp.Show vbModal
        GridConfig.SetFocus
        '    GridConfig.Col = 4
        '    GridConfig.EditActive = True
      End If

    Case 5

      If GridConfig.AllowUpdate = True Then
        
        FrmLookUp.TitleForm = "Data Combo"
        Set FrmLookUp.FormCaller = Me
        Set FrmLookUp.FormContainer = GridConfig.Columns(ColIndex)
        '    Set FrmLookUp.FormContainer2 = GridConfig.Columns(ColIndex + 1)
        
        '   strSQL = "SELECT [Catalogue No], [Engineering Description], [Equipment ID] From [Equipment APLs Table]"
        strSQL = "SELECT Table_Name,Table_Type From kolom_object where table_type='VIEW' Group by Table_Name,Table_Type"
        FrmLookUp.SQLScript = strSQL
        FrmLookUp.ColRefNumber = 0
        FrmLookUp.ColRefNumber2 = 0
        FrmLookUp.ColRefNumber3 = 0
        Load FrmLookUp
        FrmLookUp.Show vbModal
        GridConfig.SetFocus
      End If

  End Select
Exit Sub
1:
MessageBox Err.Description, "frmconfig:gridconfig_buttonclick" & Err.Number, msgOkOnly, msgExclamation
End Sub



