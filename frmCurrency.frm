VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Begin VB.Form frmCurrency 
   Caption         =   "Master Currency"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCurrency.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   10335
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   9990
      TabIndex        =   1
      Top             =   0
      Width           =   10050
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   3450
         Left            =   375
         ScaleHeight     =   3420
         ScaleWidth      =   7290
         TabIndex        =   2
         Top             =   660
         Width           =   7320
         Begin VB.TextBox txtBox 
            DataField       =   "Currency Name"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   4
            Tag             =   "Partner"
            Top             =   570
            Width           =   3045
         End
         Begin VB.TextBox txtBox 
            DataField       =   "CurrID"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1965
            MaxLength       =   5
            TabIndex        =   3
            Tag             =   "Partner"
            Top             =   210
            Width           =   1935
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2310
            Index           =   0
            Left            =   210
            TabIndex        =   5
            Tag             =   "Partner"
            Top             =   1005
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   4075
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   2
            RowHeight       =   16
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "CurrID"
               Caption         =   "Mata Uang"
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
               DataField       =   "Currency Name"
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
               MarqueeStyle    =   4
               BeginProperty Column00 
                  ColumnWidth     =   1814.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   4364.788
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
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
            Index           =   1
            Left            =   630
            TabIndex        =   7
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mata Uang"
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
            Index           =   0
            Left            =   705
            TabIndex        =   6
            Top             =   255
            Width           =   990
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   5745
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1217
      BindFormTAG     =   "Partner"
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'ScanKey KeyCode, Shift, MyDDE
'End Sub
'
'Private Sub Form_Load()
'OpenDB
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If MyDDE.CheckRecordPendinged = True Then
'   ScanKey vbKeyF5, 0, MyDDE
'   If MyDDE.IsSucces = True Then
'      Cancel = False
'      MyDDE.ClearRecordset
'   Else
'      Cancel = True
'   End If
'Else
'   MyDDE.ClearRecordset
'End If
'End Sub
'
'Private Sub Form_Resize()
'On Error Resume Next
'If Me.WindowState <> vbMaximized Then
'   Me.Height = MainMenu.ScaleHeight
'   Me.Width = MainMenu.ScaleWidth
'End If
'HiasForm Picture1, Me
'CenterForm Picture2
'Err.Clear
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''
'End Sub
'
'Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'Select Case AdReasonActiveDb
'       Case tmbAddNew:
'            mVarDataDc = True
'            txtBox(0).SetFocus
'       Case tmbEdit:
'            txtBox(0).Enabled = False
'            mVarDataDc = True
'            txtBox(1).SetFocus
'       Case tmbPrint:
'            CallRPTReport "Tabel Mata Uang.rpt"
'       Case Else: mVarDataDc = False
'End Select
'End Sub
'
'Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
'Select Case AdReasonActiveDb
'       Case tmbDelete:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               PrepareQuery
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
'       Case tmbSave:
'            If MyDDE.CheckEmptyControl = False Then
'               MyDDE.IsChildMemberReady = True
'               PrepareQuery
'            Else
'               MyDDE.IsChildMemberReady = False
'            End If
'End Select
'End Sub
'
'Private Sub txtBox_GotFocus(Index As Integer)
'Block txtBox(Index)
'End Sub
'
'Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then KeyEnter KeyCode
'End Sub
'
'Private Sub OpenDB()
'With MyDDE
'    .EditModeReplace = False
'    Set .BindForm = frmMataUang
'    .BindFormTAG = "Partner"
'    Set .ActiveConnection = Cnn
'    .PrepareQuery = "Select * from [Currency Table]"
'End With
'End Sub
'
'Private Sub PrepareQuery()
'With MyDDE
'    .PrepareAppend = " INSERT INTO [Currency Table] (CurrID, [Currency Name]) " & _
'                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "')"
'
'    .PrepareUpdate = " UPDATE [Currency Table] Set [Currency Name] = N'" & ValidString(txtBox(1)) & "' WHERE     (CurrID = N'" & ValidString(txtBox(0)) & "')"
'
'    .PrepareDelete = " DELETE FROM [Currency Table] WHERE   (CurrID = N'" & ValidString(txtBox(0)) & "') "
'End With
'End Sub
