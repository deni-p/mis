VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCallerBaru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmcallerBaru.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8220
      TabIndex        =   1
      Top             =   5085
      Width           =   8280
      Begin VB.TextBox txtCari 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   5
         Top             =   285
         Width           =   3000
      End
      Begin VB.CommandButton CmdOK 
         Appearance      =   0  'Flat
         Caption         =   "P&ilih"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   6555
         Picture         =   "frmcallerBaru.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   800
      End
      Begin VB.CommandButton CmdExit 
         Appearance      =   0  'Flat
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7350
         Picture         =   "frmcallerBaru.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   800
      End
      Begin VB.CommandButton cmdRefresh 
         Appearance      =   0  'Flat
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5760
         Picture         =   "frmcallerBaru.frx":7366
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cari Kriteria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   6
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
      Height          =   5475
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   8280
      TabIndex        =   0
      Top             =   0
      Width           =   8280
      Begin VB.PictureBox Picture1 
         Height          =   4800
         Left            =   75
         ScaleHeight     =   4740
         ScaleWidth      =   8040
         TabIndex        =   8
         Top             =   75
         Width           =   8100
         Begin MSDataGridLib.DataGrid GridLook 
            Height          =   4740
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   8361
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
      Begin VB.Frame FrTombol 
         Height          =   30
         Left            =   -30
         TabIndex        =   7
         Top             =   4980
         Width           =   8370
      End
   End
End
Attribute VB_Name = "frmCallerBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLook As ADODB.Recordset
Private WithEvents mVarFormObject As Recordset
Attribute mVarFormObject.VB_VarHelpID = -1
Dim CON As New Connection
Dim yRow, xCol As Integer
Dim mLoop As Byte


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Tsample = True Then
    frmMPermintaanSample.Text1(1).Text = GridLook.Columns(1).Text
    frmMPermintaanSample.Text1(2).Text = GridLook.Columns(0).Text
ElseIf TSalesforcast = True Then
    FrmSalesForecast.txtBox(2).Text = GridLook.Columns(1).Text  'nama
    FrmSalesForecast.lblsalesID.Caption = GridLook.Columns(0).Text  'nik
Else
    FrmPolicy.txtBox(0).Text = GridLook.Columns(0).Text 'Nik
    FrmPolicy.txtBox(2).Text = GridLook.Columns(1).Text 'Nama
    FrmPolicy.txtBox(7).Text = GridLook.Columns(2).Text 'Kode Area
    FrmPolicy.txtkodeDept.Text = GridLook.Columns(4).Text 'Kode Dept
    FrmPolicy.txtNameDept.Text = GridLook.Columns(5).Text ' Nama Dept
End If
Tsample = False
TSalesforcast = False
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
On Error GoTo LoadErr
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

Me.Caption = mTitleForm
Exit Sub
LoadErr:
    MessageBox Err.Description, "Form_Load", msgOkOnly, msgExclamation
End Sub

Private Sub GridLook_Click()
On Error Resume Next
mLoop = xCol
Label1 = "&Cari kriteria  berdasarkan " & UCase(GridLook.Columns(mLoop).DataField)
GridLook.col = xCol
Err.Clear
End Sub

Private Sub GridLook_DblClick()
Call cmdOk_Click
End Sub

Private Sub GridLook_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex <= 0 Then
   mLoop = 0
Else
   mLoop = ColIndex
End If
Label1 = "&Cari kriteria  berdasarkan " & UCase(GridLook.Columns(ColIndex).DataField)
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


Public Property Set FormData(ByVal vData As Recordset)
    Me.Caption = mvarTagForm
    Set mVarFormObject = vData
    Set GridLook.DataSource = mVarFormObject
 '   AdjustDataGridColumns DgDetail, mVarFormObject, mVarFormObject.Recordcount, mVarFormObject.Fields.Count, True
   ' Tampil
End Property

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
