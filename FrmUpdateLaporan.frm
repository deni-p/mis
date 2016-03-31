VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUpdateLaporan 
   Caption         =   "Seting Laporan Tambahan"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUpdateLaporan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9825
   Begin VB.CommandButton cmdLink 
      Height          =   330
      Left            =   5880
      Picture         =   "FrmUpdateLaporan.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3270
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   225
      ScaleHeight     =   5295
      ScaleWidth      =   10020
      TabIndex        =   11
      Top             =   240
      Width           =   10050
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4560
         Left            =   360
         ScaleHeight     =   4530
         ScaleWidth      =   7290
         TabIndex        =   12
         Top             =   540
         Width           =   7320
         Begin VB.CommandButton SemeruButton1 
            Caption         =   "Keluar"
            Height          =   375
            Left            =   3660
            TabIndex        =   10
            Top             =   3135
            Width           =   1560
         End
         Begin VB.CommandButton CmdProses 
            Caption         =   "Buka File"
            Height          =   375
            Left            =   1665
            TabIndex        =   9
            Top             =   3135
            Width           =   1560
         End
         Begin VB.TextBox txtBox 
            DataField       =   "AliasReport"
            Height          =   315
            Index           =   2
            Left            =   1665
            MaxLength       =   200
            TabIndex        =   5
            Tag             =   "Design"
            Top             =   2115
            Width           =   3555
         End
         Begin VB.TextBox txtBox 
            DataField       =   "Notes"
            Height          =   315
            Index           =   1
            Left            =   1665
            MaxLength       =   200
            TabIndex        =   8
            Tag             =   "Design"
            Top             =   2775
            Width           =   3555
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Rekapitulasi"
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
            Index           =   4
            Left            =   1665
            TabIndex        =   4
            Top             =   1560
            Width           =   3000
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Inventory"
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
            Index           =   3
            Left            =   1665
            TabIndex        =   3
            Top             =   1224
            Width           =   3000
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Laporan Bulanan"
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
            Index           =   2
            Left            =   1665
            TabIndex        =   2
            Top             =   891
            Width           =   3000
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Transaksi"
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
            Left            =   1665
            TabIndex        =   1
            Top             =   558
            Width           =   3000
         End
         Begin VB.OptionButton chkOpt 
            Caption         =   "Master Data"
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
            Left            =   1665
            TabIndex        =   0
            Top             =   225
            Value           =   -1  'True
            Width           =   3000
         End
         Begin VB.TextBox txtBox 
            DataField       =   "FileNameReport"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1665
            MaxLength       =   500
            TabIndex        =   6
            Tag             =   "Design"
            Top             =   2445
            Width           =   3555
         End
         Begin MSComDlg.CommonDialog Dialog 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alias"
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
            Index           =   4
            Left            =   1155
            TabIndex        =   15
            Top             =   2160
            Width           =   420
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
            Index           =   3
            Left            =   510
            TabIndex        =   14
            Top             =   2820
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Laporan"
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
            Index           =   2
            Left            =   270
            TabIndex        =   13
            Top             =   2490
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "FrmUpdateLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVarLokationRPT As String
Private mVarLokationUDT As String
Private mVarRPTName, Midex As String
Private IdxOpt As Integer

Private Sub chkOpt_Click(Index As Integer)
IdxOpt = Index
End Sub

Private Sub cmdLink_Click()
On Error GoTo RepERR
With Dialog
    .InitDir = App.Path & "\Report"
    .Filter = "*.rpt|*.rpt" '"Crystal Report"
    .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
    .ShowOpen
    If .Filename = "" Then
    Else
        txtBox(0) = .FileTitle
        mVarLokationRPT = .Filename
        mVarLokationUDT = Trim(Left(.Filename, Len(.Filename) - 4)) & ".UDT"
    End If
End With
'txtBox(0).SetFocus
RepERR:
    If Err <> 0 Then
        MessageBox Err.Description & " - " & Err.Number, "Peringatan"
    End If
End Sub

Private Sub CmdProses_Click()
If FileExist(mVarLokationUDT) = True Then
    If txtBox(2) = "" Or txtBox(1) = "" Then
       MessageBox "Seting laporan tidak boleh kosong", "Peringatan", msgOkOnly
       Exit Sub
    End If
    Select Case IdxOpt
           Case 0: Midex = "Mst-" & CreateIdx
           Case 1: Midex = "Trs-" & CreateIdx
           Case 2: Midex = "Lbl-" & CreateIdx
           Case 3: Midex = "Lit-" & CreateIdx
           Case 4: Midex = "Rek-" & CreateIdx
    End Select
    
    'MessageBox BukafilePath(mVarLokationUDT)
    If CreateView(BukafilePath(mVarLokationUDT), txtBox(0), Cnn.ConnectionString) = True Then
       DOBackup mVarLokationRPT, App.Path & "\Report\" & txtBox(0)
       SendDataToServer (" INSERT INTO [Report Modules]  (IDReport, ModulesName, AliasReport, FileNameReport,Notes)" & _
                         " VALUES     (N'" & Midex & "', N'" & chkOpt(IdxOpt).Caption & "', N'" & txtBox(2) & "', N'" & txtBox(0) & "',N'" & txtBox(1) & "')")
    End If
Else
   MessageBox "Nama file belum ada.", "Peringatan", msgOkOnly
End If
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
HiasForm Picture1, Me
CenterForm Picture2, Me
chkOpt(0).BackColor = Picture2.BackColor
chkOpt(1).BackColor = Picture2.BackColor
chkOpt(2).BackColor = Picture2.BackColor
chkOpt(3).BackColor = Picture2.BackColor
chkOpt(4).BackColor = Picture2.BackColor
IdxOpt = 0
End Sub

Private Sub Form_Resize()

'HiasForm Picture1, Me
'CenterForm Picture2, Me
'chkOpt(0).BackColor = Picture2.BackColor
'chkOpt(1).BackColor = Picture2.BackColor
'chkOpt(2).BackColor = Picture2.BackColor
'chkOpt(3).BackColor = Picture2.BackColor
'chkOpt(4).BackColor = Picture2.BackColor
Err.Clear
End Sub

Private Function CreateView(ByVal CommandProcedure As String, ByVal CommandProcedureName As String, ByVal ConnectionString As String) As Boolean
On Error Resume Next
Dim Icom As New Command
Dim Iconn As New Connection
Iconn.CursorLocation = adUseClient
Iconn.Mode = adModeShareExclusive
Iconn.IsolationLevel = adXactIsolated
Iconn.ConnectionString = ConnectionString
Iconn.Open
Set Icom.ActiveConnection = Iconn
With Icom
     .CommandType = adCmdUnknown
     .CommandText = "Drop View [" & Replace(UCase(CommandProcedureName), ".RPT", "") & "]"
     .Execute
     Err.Clear
     .CommandText = "Create View [" & Replace(UCase(CommandProcedureName), ".RPT", "") & "] AS " & CommandProcedure
     .Execute
     CreateView = True
End With
Hell:
    Set Icom = Nothing
    Iconn.Close
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Warning"
    Err.Clear
End Function

Private Function BukafilePath(Filename As String) As String
  Dim NamaFile As String
  Open Filename For Input As #1
     Input #1, NamaFile
     BukafilePath = NamaFile
     'MessageBox NamaFile
  Close #1
End Function

Private Sub DOBackup(ByVal SourceFile As String, ByVal DestinationFile As String, Optional ByVal WarningMessageOff As Boolean)
On Error GoTo Hell
Dim f As New FileSystemObject
If (FileExist(DestinationFile)) Then Kill DestinationFile
If (FileExist(SourceFile)) Then
   f.CopyFile SourceFile, DestinationFile, True
   MessageBox "Update data laporan telah selesai", "Perhatian", msgOkOnly
Else
   MessageBox "File Sumber Tidak Ada. Harap Dicari Dulu File Yang Valid", "Perhatian", msgOkOnly
End If
Exit Sub
Hell:
    MessageBox Err.Number & " - " & Err.Description, "Fatal Error", msgOkOnly
End Sub

Private Function FileExist(ByVal Filename As String) As Boolean
  On Error GoTo FileDoesNotExist
  Call FileLen(Filename)
  FileExist = True
  Exit Function
FileDoesNotExist:
  FileExist = False
End Function

Private Function CreateIdx() As String
Dim RcIdx As New DBQuick
Dim mVarNo As Integer
RcIdx.DBOpen "SELECT MAX(RIGHT(IDReport, 5)) AS MaxNo FROM [Report Modules]", Cnn, lckLockReadOnly
With RcIdx
     If .Recordcount <> 0 Then
        mVarNo = IIf(Not IsNull(.Fields(0)), .Fields(0), 0)
     Else
        mVarNo = 0
     End If
     mVarNo = mVarNo + 1
     Select Case Len(Trim(Str(mVarNo)))
            Case 1: CreateIdx = "0000" & Trim(Str(mVarNo))
            Case 2: CreateIdx = "000" & Trim(Str(mVarNo))
            Case 3: CreateIdx = "00" & Trim(Str(mVarNo))
            Case 4: CreateIdx = "0" & Trim(Str(mVarNo))
            Case 5: CreateIdx = Trim(Str(mVarNo))
     End Select
End With
RcIdx.CloseDB
End Function

Private Sub SemeruButton1_Click()
Unload Me
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub txtBox_Validate(Index As Integer, Cancel As Boolean)
If txtBox(Index) = "" Then
   MessageBox "Data tidak boleh kosong.", "Peringatan", msgOkOnly
   Cancel = True
End If
End Sub
