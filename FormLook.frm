VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormLook 
   BackColor       =   &H00EEDAC1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cari Barang"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "FormLook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDAC1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   6210
      TabIndex        =   3
      Top             =   3930
      Width           =   6210
      Begin VB.CommandButton CmdFind 
         Caption         =   "&Cari"
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
         Left            =   4155
         TabIndex        =   10
         Top             =   98
         Width           =   660
      End
      Begin VB.OptionButton OptSearch 
         BackColor       =   &H00EEDAC1&
         Caption         =   "&Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   9
         Top             =   143
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton OptSearch 
         BackColor       =   &H00EEDAC1&
         Caption         =   "&Kode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   143
         Width           =   675
      End
      Begin VB.CommandButton CmdRefresh 
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
         Left            =   1470
         Picture         =   "FormLook.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresh"
         Top             =   91
         Width           =   330
      End
      Begin VB.Frame FrTombol 
         Height          =   30
         Left            =   -45
         TabIndex        =   6
         Top             =   0
         Width           =   6735
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Pilih"
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
         Left            =   4815
         TabIndex        =   5
         Top             =   98
         Width           =   660
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Batal"
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
         Left            =   5475
         TabIndex        =   4
         Top             =   98
         Width           =   660
      End
      Begin VB.TextBox TxtLook 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1815
         TabIndex        =   0
         Top             =   83
         Width           =   2280
      End
   End
   Begin VB.PictureBox PictCari 
      BackColor       =   &H00EAAF6F&
      Height          =   3735
      Left            =   30
      ScaleHeight     =   3675
      ScaleWidth      =   6045
      TabIndex        =   1
      Top             =   45
      Width           =   6105
      Begin MSDataGridLib.DataGrid GridLook 
         Height          =   3570
         Left            =   60
         TabIndex        =   2
         Top             =   45
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6297
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            MarqueeStyle    =   4
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FormLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String
Dim rsLook As ADODB.Recordset
Dim strSQL As String
Private mVarFormObject As Object 'local copy
Private mVarObjectContainer As Object 'local copy
Private mVarObjectContainer2 As Object 'local copy
Private mVarObjectContainer3 As Object 'local copy
Private mVarColData As Integer   'local copy
Private mVarSQLScript As String  'local copy
Private mCaption As String

Private Sub CmdCancel_Click()
mVarObjectContainer = ""
mVarObjectContainer2 = ""
Unload Me
End Sub

Private Sub cmdOk_Click()
GridLook_DblClick
End Sub

Private Sub Form_Activate()
'GridLook.SetFocus
'TxtLook.SetFocus
OptSearch_Click (1)
'Gelas False
End Sub

Private Sub Form_Load()
On Error GoTo 1
Set rsLook = New ADODB.Recordset
rsLook.CursorLocation = adUseClient
rsLook.Open mVarSQLScript, CNN, adOpenStatic, adLockReadOnly, adCmdText
Set GridLook.DataSource = rsLook
GridLayout
Screen.MousePointer = 0
Me.Caption = mCaption
Exit Sub
1:
MessageBox Err.Description, "formlook:form_load" & Err.Number, msgOkOnly, msgExclamation
End Sub
Private Sub GridLayout()
With GridLook
   .Columns(0).width = 1200
   .Columns(1).width = 4500     'NAMA
'   .Columns(2).Width = 1150     'TANGGAL
'   .Columns(2).NumberFormat = ShortDateForm
'   .Columns(3).Width = 1500     'STATUS
'   .Columns(3).Alignment = dbgLeft
'   .Height = 3575
   .HoldFields
End With

If rsLook.Recordcount > GridLook.VisibleRows Then
'   PictCari.Width = 6225
'   Me.Width = 6385
Else
'   PictCari.Width = 6225
'   Me.Width = 6390
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo KeyErr
Select Case KeyCode
   Case 13, vbKeyF9  'PRINT
      mVarObjectContainer = GridLook.Columns(mVarColData).Value
      mVarObjectContainer2 = GridLook.Columns(mVarColData + 1).Value
      Unload Me
   Case 27, vbKeyF10  'DESIGN
      Unload Me
   Case Else
      Exit Sub
End Select
Exit Sub

KeyErr:
   MessageBox Err.Description, vbExclamation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo OutErr
rsLook.Close
Set rsLook = Nothing

Exit Sub
OutErr:
   MessageBox Err.Description, vbExclamation
End Sub

Private Sub GridLook_DblClick()
On Error GoTo DblErr
Screen.MousePointer = 11
mVarObjectContainer = GridLook.Columns(mVarColData).Value
mVarObjectContainer2 = GridLook.Columns(mVarColData + 1).Value
mVarObjectContainer3 = GridLook.Columns(mVarColData + 2).Value
Unload Me
Exit Sub

DblErr:
   MessageBox Err.Description, vbExclamation

End Sub

Private Sub GridLook_HeadClick(ByVal ColIndex As Integer)
'rsLook.Sort = GridLook.Columns(ColIndex).DataField
End Sub

Public Property Set FormCaller(ByVal vData As Object)
    Set mVarFormObject = vData
End Property

Public Property Get FormCaller() As Object
    Set FormCaller = mVarFormObject
End Property

Public Property Set TextContainer(ByVal vData As Object)
    Set mVarObjectContainer = vData
End Property

Public Property Get TextContainer() As Object
    Set TextContainer = mVarObjectContainer
End Property
Public Property Set TextContainer2(ByVal vData As Object)
    Set mVarObjectContainer2 = vData
End Property

Public Property Get TextContainer2() As Object
    Set TextContainer2 = mVarObjectContainer2
End Property
Public Property Set TextContainer3(ByVal vData As Object)
    Set mVarObjectContainer3 = vData
End Property

Public Property Get TextContainer3() As Object
    Set TextContainer3 = mVarObjectContainer3
End Property
Public Property Let ColRefNumber(ByVal vData As Integer)
    mVarColData = vData
End Property

Public Property Let SQLScript(ByVal vData As String)
    mVarSQLScript = vData
End Property

Public Property Let JudulForm(ByVal vData As String)
    mCaption = vData
End Property
Private Sub OptSearch_Click(Index As Integer)
On Error GoTo 1
'If Index = 0 Then
'   strFilter = "NoInduk"
'Else
'   strFilter = "Nama"
'End If
strFilter = GridLook.Columns(Index).Caption
TxtLook.SetFocus
TxtLook.Text = ""
Exit Sub
1:
MessageBox Err.Description, "formlook:optsearch_click" & Err.Number, msgOkOnly, msgExclamation
End Sub


Private Sub TxtLook_Change()

'strFilter = "Nama "
On Error GoTo 1
If Len(TxtLook.Text) <> 0 Then
   rsLook.Filter = strFilter & " like '" & TxtLook.Text & "*'"
Else
   cmdRefresh.Value = True
End If
Exit Sub
1:
MessageBox Err.Description, "formlook" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub cmdRefresh_Click()
rsLook.Filter = adFilterNone
TxtLook.Text = ""
rsLook.Requery
GridLayout
TxtLook.SetFocus
End Sub


