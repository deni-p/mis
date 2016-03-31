VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmLookUp 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type of Analysis"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000000&
   Icon            =   "FrmLookUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicTombol 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDAC1&
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
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   3315
      Width           =   5355
      Begin VB.CommandButton CmdSelect 
         Caption         =   "F2 - &Select"
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
         Left            =   3225
         TabIndex        =   2
         Top             =   105
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "F3 - &Cancel"
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
         Left            =   4200
         TabIndex        =   3
         Top             =   105
         Width           =   975
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
         Left            =   90
         Picture         =   "FrmLookUp.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Refresh"
         Top             =   105
         Width           =   330
      End
      Begin VB.TextBox TxtCari 
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
         Left            =   420
         TabIndex        =   1
         Top             =   105
         Width           =   2475
      End
      Begin VB.Frame FrTombol 
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
         Left            =   -45
         TabIndex        =   7
         Top             =   0
         Width           =   5400
      End
   End
   Begin VB.PictureBox PictLook 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3175
      Left            =   60
      ScaleHeight     =   3120
      ScaleWidth      =   5160
      TabIndex        =   5
      Top             =   60
      Width           =   5220
      Begin MSDataGridLib.DataGrid GridLook 
         Height          =   3150
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   5155
         _ExtentX        =   9102
         _ExtentY        =   5556
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ForeColor       =   -2147483641
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
            MarqueeStyle    =   3
            RecordSelectors =   0   'False
            ScrollGroup     =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLook As ADODB.Recordset

Private mVarFormObject As Object

Private mVarObjectContainer As Object

Private mVarObjectContainer2 As Object

Private mVarObjectContainer3 As Object

Private mVarColData As Integer

Private mVarColData2 As Integer

Private mVarColData3 As Integer

Private mVarSQLScript As String

Private mTitleForm As String

Public Property Let ColRefNumber(ByVal vData As Integer)
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  mVarColData = vData
End Property

Public Property Let ColRefNumber2(ByVal vData As Integer)
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  mVarColData2 = vData
End Property

Public Property Let ColRefNumber3(ByVal vData As Integer)
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  mVarColData3 = vData
End Property

Public Property Let SQLScript(ByVal vData As String)
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  mVarSQLScript = vData
End Property

Public Property Let TitleForm(ByVal vData As String)
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  mTitleForm = vData
End Property

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdRefresh_Click()
  'Gelas True
  rsLook.Filter = adFilterNone
  rsLook.Requery
  txtCari.Text = ""
  txtCari.SetFocus
  GridLook.Columns(0).width = 2500
  GridLook.Columns(1).width = 3610
 ' Gelas False
End Sub

Private Sub CmdSelect_Click()
  On Error GoTo DblErr
  GridLook_DblClick
  Unload Me
  Exit Sub

DblErr:


  MessageBox Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
  GridLook.SetFocus
 ' Gelas False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
  On Error GoTo KeyErr

  Select Case KeyCode

    Case 13

      mVarObjectContainer = GridLook.Columns(mVarColData).Value

      If mVarColData2 <> 0 Then mVarObjectContainer2 = GridLook.Columns(mVarColData + mVarColData2).Text
      If mVarColData3 <> 0 Then mVarObjectContainer3 = GridLook.Columns(mVarColData + mVarColData3).Text

      Unload Me

      '      Me.Hide
    Case 27

      '      Me.Hide
      Unload Me
   
  End Select

  Exit Sub

KeyErr:


  MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
End Sub

Private Sub Form_Load()
  Dim I As Integer
  Dim strSQL As String
  Set rsLook = New ADODB.Recordset
  rsLook.CursorLocation = adUseClient
  rsLook.Open mVarSQLScript, CNN, adOpenKeyset, adLockReadOnly, adCmdText
  Set GridLook.DataSource = rsLook

  With GridLook
    '   .Height = 3125
    .Columns(0).width = 2500
    .Columns(1).width = 3610
    '   If .Columns.Count > 2 Then
    '      .Columns(2).Alignment = dbgCenter
    '      Set .Columns(2).DataFormat = fmtBooleanData
    '   End If
  End With

  If rsLook.Recordcount <= 12 Then
    GridLook.Columns(1).width = 3610 + 250
  End If

  Me.Caption = mTitleForm
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
  On Error GoTo OutErr
  rsLook.Close
  Set rsLook = Nothing

  Exit Sub
OutErr:


  MessageBox Err.Description, "Error", msgOkOnly, msgExclamation
End Sub

Private Sub GridLook_DblClick()
  On Error GoTo DblErr
  mVarObjectContainer = GridLook.Columns(mVarColData).Value

  If mVarColData2 <> 0 Then mVarObjectContainer2 = GridLook.Columns(mVarColData + mVarColData2).Text
  If mVarColData3 <> 0 Then mVarObjectContainer3 = GridLook.Columns(mVarColData + mVarColData3).Text
  Unload Me
  Exit Sub

DblErr:


  MessageBox Err.Description, "Error", msgOkOnly, msgExclamation

End Sub

Private Sub GridLook_HeadClick(ByVal ColIndex As Integer)
  rsLook.Sort = "[" & GridLook.Columns(ColIndex).DataField & "]"
End Sub

Private Sub TxtCari_Change()
  On Error Resume Next

  If Len(txtCari.Text) <> 0 Then
    rsLook.Filter = "[" & GridLook.Columns(1).Caption & "] like '" & txtCari.Text & "*'"
  Else
    cmdRefresh.Value = True
  End If

End Sub

Public Property Get FormCaller() As Object
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  Set FormCaller = mVarFormObject
End Property

'Private Sub GridLook_KeyPress(KeyAscii As Integer)
'rsLook.Find "ItemCode like  '" & Chr(KeyAscii) & "*'"
'If rsLook.EOF Then rsLook.MoveFirst
'End Sub

Public Property Set FormCaller(ByVal vData As Object)
  'used when assigning an Object to the property, on the left side of a Set statement.
  'Syntax: Set x.Database = Form1
  Set mVarFormObject = vData
End Property

Public Property Get FormContainer2() As Object
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  Set FormContainer2 = mVarObjectContainer2
End Property

Public Property Set FormContainer2(ByVal vData As Object)
  'used when assigning an Object to the property, on the left side of a Set statement.
  'Syntax: Set x.Database = Form1
  Set mVarObjectContainer2 = vData
End Property

Public Property Get FormContainer3() As Object
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  Set FormContainer3 = mVarObjectContainer3
End Property

Public Property Set FormContainer3(ByVal vData As Object)
  'used when assigning an Object to the property, on the left side of a Set statement.
  'Syntax: Set x.Database = Form1
  Set mVarObjectContainer3 = vData
End Property

Public Property Get FormContainer() As Object
  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.Database
  Set FormContainer = mVarObjectContainer
End Property

Public Property Set FormContainer(ByVal vData As Object)
  'used when assigning an Object to the property, on the left side of a Set statement.
  'Syntax: Set x.Database = Form1
  Set mVarObjectContainer = vData
End Property

