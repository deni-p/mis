VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesTeam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Team"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSalesTeam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8685
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
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8625
      TabIndex        =   7
      Top             =   5865
      Width           =   8685
      Begin VB.CommandButton cmdRefresh 
         Appearance      =   0  'Flat
         Caption         =   "&Refresh"
         Height          =   550
         Left            =   6840
         Picture         =   "frmSalesTeam.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   800
      End
      Begin VB.TextBox txtCari 
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
         Left            =   75
         TabIndex        =   1
         Top             =   285
         Width           =   3000
      End
      Begin VB.CommandButton CmdExit 
         Appearance      =   0  'Flat
         Caption         =   "&Keluar"
         Height          =   550
         Left            =   7680
         Picture         =   "frmSalesTeam.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cari Kriteria"
         Height          =   195
         Left            =   75
         TabIndex        =   8
         Top             =   45
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5820
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   8685
      TabIndex        =   4
      Top             =   0
      Width           =   8685
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5640
         Left            =   75
         ScaleHeight     =   5580
         ScaleWidth      =   8400
         TabIndex        =   5
         Top             =   75
         Width           =   8460
         Begin MSDataGridLib.DataGrid GridLook 
            Height          =   5460
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   8400
            _ExtentX        =   14817
            _ExtentY        =   9631
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSMask.MaskEdBox MaskCari 
      Height          =   270
      Left            =   4950
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmSalesTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsYes As Integer
Private WithEvents mVarFormObject As Recordset
Attribute mVarFormObject.VB_VarHelpID = -1
Private mLoop As Integer
Private mvarTagForm As String
Private myStd As New StdDataFormat
Private mVarBoolean, mVarLoad As Boolean
Private MyFrm As Form
Private LookupCaller As Form
Event RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
Event ActionJackson(ByVal ColIndex As Integer)
Event CallLinkForm()
Event BeforeUnload()
Dim rsLook As ADODB.Recordset
Dim CON As New Connection

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
Dim sqltxt As String

Set rsLook = New ADODB.Recordset
sqltxt = "select * from q_karyawan_user order by empid"
rsLook.CursorLocation = adUseClient
rsLook.Open sqltxt, CON, adOpenKeyset, adLockReadOnly, adCmdText
Set GridLook.DataSource = rsLook
GridLook.Columns(2).Visible = False
GridLook.Columns(3).Visible = False
GridLook.Columns(6).Visible = False
GridLook.Columns(7).Visible = False
GridLook.Columns(8).Visible = False
GridLook.Columns(9).Visible = False
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim strSQL, sqltxt As String

HiasFormManTell Picture2, Me
Set CON = New ADODB.Connection
Set rsLook = New ADODB.Recordset

strSQL = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & GetSetting("Manufacturing Intelligent", "SERVER", "Server Name") & ";DATABASE=hris_asml;USER=bulirpadi;PASSWORD=bulirpadi;OPTION=3;"
'strSQL = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=servermmt;DATABASE=hris_asml;USER=bulirpadi;PASSWORD=bulirpadi;OPTION=3;"
CON.Open strSQL

sqltxt = "select * from q_karyawan_user order by empid"
rsLook.CursorLocation = adUseClient
rsLook.Open sqltxt, CON, adOpenKeyset, adLockReadOnly, adCmdText
Set GridLook.DataSource = rsLook

GridLook.Columns(2).Visible = False
GridLook.Columns(3).Visible = False
GridLook.Columns(6).Visible = False
GridLook.Columns(7).Visible = False
GridLook.Columns(8).Visible = False
GridLook.Columns(9).Visible = False

'Me.Caption = mTitleForm
End Sub


Private Sub GridLook_HeadClick(ByVal ColIndex As Integer)
'On Error Resume Next
If ColIndex <= 0 Then
   mLoop = 0
Else
   mLoop = ColIndex
End If
Label1 = "&Cari Kriteria by " & GridLook.Columns(ColIndex).DataField
GridLook.col = ColIndex
Err.Clear
End Sub

Private Sub TxtCari_Change()
Dim strcari As String
If txtCari <> "" Then
       If rsLook.Recordcount <> 0 Then
          strcari = "[" & GridLook.Columns(mLoop).DataField & "]" & " Like '" & txtCari & "%'"
          rsLook.Filter = strcari ', 0, adSearchForward, adBookmarkFirst
          If rsLook.Recordcount = 0 Then MsgBox "Kriteria Yang Dicari Tidak Ada..............!", vbCritical
       Else
          Call cmdRefresh_Click
       End If
Else
    Call cmdRefresh_Click
End If
End Sub

Public Property Let FindCaller(ByVal FindStringValue As String)
On Error GoTo Hell
If Not mVarFormObject Is Nothing Then
   If mVarFormObject.State = 1 Then
      If FindStringValue <> "" Then
         mVarFormObject.Find FindStringValue, 0, adSearchForward, adBookmarkFirst
         If mVarFormObject.EOF Then
            mVarFormObject.AbsolutePosition = 1
         End If
      End If
   End If
End If
Hell:
    Err.Clear
End Property


Public Property Get GetFieldByName(ByVal FldListFields As Variant) As Variant
On Error Resume Next
If mVarFormObject.Recordcount <> 0 Then
   GetFieldByName = IIf(Not IsNull(mVarFormObject.Fields(FldListFields).Value), mVarFormObject.Fields(FldListFields).Value, Empty)
Else
   GetFieldByName = Empty
End If
Err.Clear
End Property

Public Property Set FormData(ByVal vData As Recordset)
    Me.Caption = mvarTagForm
    Set mVarFormObject = vData
    Set GridLook.DataSource = mVarFormObject
    Tampil
End Property


Private Sub Tampil()
    Dim I As Integer
    Label1 = "&Cari Kriteria by " & mVarFormObject.Fields(I).Name
    GridLook.Height = (GridLook.RowHeight * 19) + 50
    For I = 0 To GridLook.Columns.Count - 1
        If (GridLook.Columns.Count - 1) >= 1 Then
            If (GridLook.RowHeight * mVarFormObject.Recordcount) > GridLook.Height Then
                If (GridLook.Columns.Count - 1) > 2 Then
                   GridLook.Columns(I).width = ((GridLook.width / 1.8) / 2) + 184
                ElseIf (GridLook.Columns.Count - 1) <= 2 Then
                   GridLook.Columns(I).width = (GridLook.width / 2) - 280
                End If
            Else
                If (GridLook.Columns.Count) = 2 Then
                    GridLook.Columns(I).width = (GridLook.width / 2) - 300
                ElseIf (GridLook.Columns.Count - 1) >= 2 Then
                   GridLook.Columns(I).width = ((GridLook.width / 1.8) / 2) + 250
                ElseIf (GridLook.Columns.Count - 1) <= 2 Then
                   GridLook.Columns(I).width = (GridLook.width / 2) - 190
                End If
            End If
        Else
           GridLook.Columns(I).width = GridLook.width - 600
        End If
        GridLook.Columns(I).DividerStyle = dbgLightGrayLine
        GridLook.Columns(I).Locked = True
        With mVarFormObject
            Select Case .Fields(I).Type
                   Case adBigInt, adCurrency, adDecimal, adDouble, adInteger
                        myStd.Type = fmtCustom
                        myStd.Format = "#,##0"
                        GridLook.Columns(I).NumberFormat = "#,##0"
                        GridLook.Columns(I).Alignment = dbgRight
                   Case adDate, adDBDate, adDBTime, adDBTimeStamp
'                        Debug.Print .Fields(I).Value
'                        myStd.Type = fmtCustom
'                        myStd.Format = "dd mm yyyy"
'                        Set gridlook.Columns(I).DataFormat = myStd
                        GridLook.Columns(I).Alignment = dbgLeft
                   Case 11:
                        myStd.Type = fmtBoolean
                        myStd.Format = "YES/NO"
                        Set GridLook.Columns(I).DataFormat = myStd
                        GridLook.Columns(I).Alignment = dbgRight
                   Case Else:
            End Select
        End With
    Next I
    GridLook.ToolTipText = Me.Caption
End Sub
