VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D04E540-01A7-41AC-A49D-31A6AB39B954}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmRegional 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regional"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegional.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Tag             =   "Regional Setting"
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4170
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   0
      ScaleHeight     =   4170
      ScaleWidth      =   8865
      TabIndex        =   7
      Top             =   0
      Width           =   8865
      Begin VB.OptionButton OptCity 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   1
         Left            =   1980
         TabIndex        =   2
         Top             =   165
         Width           =   1080
      End
      Begin VB.OptionButton OptCity 
         BackColor       =   &H00EAAF6F&
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   165
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RG"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   0
         Left            =   1590
         MaxLength       =   16
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   570
         Width           =   1935
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Code RG"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   2
         Left            =   1590
         MaxLength       =   16
         TabIndex        =   5
         Tag             =   "Partner"
         Top             =   1260
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "RG Name"
         DataSource      =   "Adodc1"
         Height          =   330
         Index           =   1
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   915
         Width           =   3045
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2130
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   1740
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   3757
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   16
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "RG"
            Caption         =   "RG"
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
            DataField       =   "RG Name"
            Caption         =   "Nama RG"
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
            DataField       =   "CODE RG"
            Caption         =   "Kode RG"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   210
         X2              =   1770
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   210
         X2              =   1770
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   210
         X2              =   1770
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Negara"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   10
         Top             =   1328
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   9
         Top             =   983
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regional"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   8
         Top             =   645
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRegional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mVarType  As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
'HiasForm Picture1, Me
MyDDE.SetPermissions = aksess.MayDo("Regional")

HiasFormManTell Picture2, Me
OptCity(0).BackColor = Picture2.BackColor
OptCity(1).BackColor = Picture2.BackColor
'OptCity(0).ForeColor = Picture1.BackColor
'OptCity(1).ForeColor = Picture1.BackColor
OpenDB
GridLayout
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
End If
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
'OptCity(0).BackColor = Picture2.BackColor
'OptCity(1).BackColor = Picture2.BackColor
'OptCity(0).ForeColor = Picture1.BackColor
'OptCity(1).ForeColor = Picture1.BackColor
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmRegional = Nothing
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
            OptCity(0).Enabled = False
            OptCity(1).Enabled = False
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
            OptCity(0).Enabled = False
            OptCity(1).Enabled = False
       Case tmbSave, tmbCancel:
            OptCity(0).Enabled = True
            OptCity(1).Enabled = True
       Case tmbPrint:
            If mVarType = "CITY" Then
               CallRPTReport "Regional City.Rpt"
            Else
               CallRPTReport "Regional Country.Rpt"
            End If
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
PrepareQuery
Err.Clear
End Sub


Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterRegional) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub

Private Sub OptCity_Click(Index As Integer)
If Index = 0 Then
   OpenDB
   Label1(2).Caption = "Postal Code"
Else
   OpenDB True
   Label1(2).Caption = "State Code"
End If
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

Private Sub OpenDB(Optional ByVal Tipical As Boolean)
With MyDDE
    .EditModeReplace = False
    Set .BindForm = frmRegional
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    If Tipical = False Then
       .PrepareQuery = "Select * from  regional where [Type RG] ='CITY'"
       mVarType = "CITY"
    Else
       .PrepareQuery = "Select * from  regional where [Type RG] ='COUNTRY'"
       mVarType = "COUNTRY"
    End If
End With
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO regional (RG, [RG Name],[Code RG],[Type RG]) " & _
                     " VALUES (N'" & ValidString(txtBox(0)) & "', N'" & ValidString(txtBox(1)) & "', N'" & ValidString(txtBox(2)) & "','" & mVarType & "')"
                     
    .PrepareUpdate = " UPDATE regional Set [RG Name] = N'" & ValidString(txtBox(1)) & "',[Code RG] = N'" & ValidString(txtBox(2)) & "' WHERE     (RG = N'" & ValidString(txtBox(0)) & "')"
                     
    .PrepareDelete = " DELETE FROM regional WHERE   (RG = N'" & ValidString(txtBox(0)) & "') "
End With
End Sub

Private Sub GridLayout()
'DataGrid1(0).Height = 2010
DataGrid1(0).width = 7950
DataGrid1(0).Columns(0).width = 1934.929
DataGrid1(0).Columns(1).width = 3614.74
DataGrid1(0).Columns(2).width = 1844.787
End Sub
