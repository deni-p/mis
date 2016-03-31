VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5895
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   30
      ScaleHeight     =   3705
      ScaleWidth      =   5775
      TabIndex        =   19
      Top             =   45
      Width           =   5805
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
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
         Height          =   2850
         Left            =   135
         ScaleHeight     =   2820
         ScaleWidth      =   5490
         TabIndex        =   20
         Top             =   675
         Width           =   5520
         Begin VB.TextBox txtUser 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   0
            Left            =   1890
            TabIndex        =   6
            Text            =   "Neo"
            Top             =   1125
            Width           =   2685
         End
         Begin VB.ComboBox cmbData 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1890
            TabIndex        =   4
            Text            =   "DEVPC2\LOCAL"
            Top             =   765
            Width           =   2685
         End
         Begin VB.ComboBox cboServer 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1890
            TabIndex        =   2
            Text            =   "DEVPC2\LOCAL"
            Top             =   405
            Width           =   2685
         End
         Begin VB.TextBox txtUser 
            Appearance      =   0  'Flat
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1890
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   1485
            Width           =   2685
         End
         Begin VB.TextBox txtUser 
            Appearance      =   0  'Flat
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1890
            TabIndex        =   10
            Top             =   1845
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00EAAF6F&
            Caption         =   "&Advance"
            Height          =   345
            Index           =   2
            Left            =   3570
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2355
            Width           =   1020
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Windows Security"
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
            Height          =   270
            Left            =   210
            TabIndex        =   0
            Top             =   165
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00EAAF6F&
            Caption         =   "&Login"
            Height          =   345
            Index           =   0
            Left            =   1530
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2355
            Width           =   1020
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00EAAF6F&
            Caption         =   "&Cancel"
            Height          =   345
            Index           =   1
            Left            =   2550
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2355
            Width           =   1020
         End
         Begin VB.CommandButton CmdLookUp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4575
            Picture         =   "frmLogin.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1860
            Visible         =   0   'False
            Width           =   330
         End
         Begin MSComDlg.CommonDialog CmnDialog 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Server Name"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   690
            TabIndex        =   1
            Top             =   473
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&User Name"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   690
            TabIndex        =   5
            Top             =   1193
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Password"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   690
            TabIndex        =   7
            Top             =   1553
            Width           =   690
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   675
            X2              =   2355
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   675
            X2              =   2355
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Database"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   690
            TabIndex        =   3
            Top             =   833
            Width           =   690
         End
         Begin VB.Line Line1 
            Index           =   4
            Visible         =   0   'False
            X1              =   690
            X2              =   2370
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Report"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   690
            TabIndex        =   9
            Top             =   1913
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   690
            X2              =   2370
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   690
            X2              =   2370
            Y1              =   1065
            Y2              =   1065
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   5895
      TabIndex        =   15
      Top             =   3900
      Width           =   5895
      Begin VB.Frame FrProgram 
         BackColor       =   &H00C0FFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   18
         Top             =   -15
         Width           =   6045
      End
      Begin VB.Label LblProgram 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Version 5.20"
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
         Index           =   1
         Left            =   4875
         TabIndex        =   17
         Top             =   135
         Width           =   900
      End
      Begin VB.Label LblProgram 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Manufacturing Intelligent   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   135
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Dim myMenu As New clsMenu
Dim idtemp As Variant
Dim exeLoadDBlist As Boolean

'Dim IDtemp As Variant
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private oSQLServerApp As SQLDMO.Application
Private mVarNamaFile, mVarLokasi As String

Private Sub LoadSQLServer2005DBList()
Dim oSQLServer As New SQLServer
On Error GoTo AdoErr
   If exeLoadDBlist Then
      'Connect to selected SQL Server
      On Error GoTo AdoErr
      oSQLServer.Connect cboServer.Text, "sa", ""
      Dim nDatabase As Integer
         'Populate dropdown with list of DB's
      cmbData.Clear
      For nDatabase = 1 To oSQLServer.Databases.Count
         cmbData.AddItem oSQLServer.Databases(nDatabase).Name
      Next nDatabase
   End If
Exit Sub
AdoErr:
   Err.Clear
   'messagebox "Error : " & Err.Description, vbExclamation, "MS SQL Server"
End Sub

Private Sub cboServer_Click()
   LoadSQLServer2005DBList
End Sub

Private Sub CmdLookUp_Click()
On Error GoTo RepERR
With CmnDialog
    .InitDir = App.Path & "\Report"
    .Filter = "Crystal Report|*.rpt"
    .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
    .ShowOpen
    If .Filename <> "" Then
        txtUser(2).Text = Left$(.Filename, Len(.Filename) - Len(.FileTitle))
        SaveSetting "Manufacturing Intelligent", "Server", "Report Location", txtUser(2).Text
    End If
End With
txtUser(2).SetFocus

RepERR:
    If Err <> 0 Then
        MsgBox Err.Description & " - " & Err.Number, vbCritical, "FormRptDesign-Load"
    End If
End Sub

Private Sub Form_Activate()
If txtUser(0).Enabled = True Then txtUser(0).SetFocus
'cmdOk(0).Value = True
'cboServer.Text = "DEVPC2\LOCAL"
End Sub

Private Sub Form_Load()
On Error Resume Next
HiasForm Picture1, Me
Check1.BackColor = Picture2.BackColor
HiasFormManTell Picture2, Me
Picture2.BackColor = &HEAAF6F
Check1.BackColor = &HEAAF6F

mVarPassword = ""
StartFromIdle = False
'CloseMenuAll
Check1.Value = BoolToInt(CBool(GetSetting("Manufacturing Intelligent", "Lisence Profile", "SQL Security")))
exeLoadDBlist = False
BukaServer
exeLoadDBlist = True
cboServer.Text = GetSetting("Manufacturing Intelligent", "Server", "Server Name")
LoadSQLServer2005DBList
cmbData.Text = GetSetting("Manufacturing Intelligent", "Server", "Database")

txtUser(0).Text = GetSetting("Manufacturing Intelligent", "Server", "UserName Active")
ReportPath = GetSetting("Manufacturing Intelligent", "Server", "Report Location")
txtUser(2).Text = ReportPath
IDLELIMIT = GetSetting("Manufacturing Intelligent", "Data", "Idle")
txtUser(1).SetFocus
'If cboServer.ListCount <> 0 Then cboServer.ListIndex = 0

StayOnTop Me
If Check1.Value = 1 Then
   Username
   txtUser(1) = ""
Else
   'txtUser(0) = ""
   'txtUser(1) = "111111"
End If
   'txtUser(0) = "Neo"
   'txtUser(1) = "111111"
LblProgram(0).BackColor = &HC0FFFF
LblProgram(1).BackColor = &HC0FFFF
LblProgram(0).FontBold = True

Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ReleaseTop Me
End Sub

Private Sub Form_Resize()


Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oSQLServerApp = Nothing
Set frmLogin = Nothing
End Sub

Private Sub BukaServer()
On Error GoTo Hell
Dim I As Integer
  Set oSQLServerApp = New SQLDMO.Application
  Dim namX As NameList
  Set namX = oSQLServerApp.ListAvailableSQLServers
  cboServer.Clear
  For I = 1 To namX.Count
    If namX.Item(I) <> "" Then cboServer.AddItem namX.Item(I)
  Next
    If I = namX.Count + 1 Then cboServer.AddItem "(local)"
    cboServer.ListIndex = 0
    Exit Sub
Hell:
    cboServer.Clear
    Set oSQLServerApp = Nothing
    cboServer.AddItem "(local)"
    cboServer.ListIndex = 0
  
  'If Err.Number <> 0 Then messagebox Err.Description, vbCritical, "Warning"
  Err.Clear
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      LoadSQLServer2005DBList
      cmdOk(0).Value = True
   End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Username
   txtUser(1) = ""
Else
   'txtUser(0) = ""
   txtUser(1) = ""
   txtUser(0).BackColor = &H80000005
   txtUser(1).BackColor = &H80000005
End If
txtUser(0).Enabled = Not CBool(Check1.Value)
txtUser(1).Enabled = Not CBool(Check1.Value)
End Sub

Private Sub cmdOk_Click(Index As Integer)
On Error GoTo Hell
Dim txtDenkrip As String
Dim mVarOpn As Variant
Screen.MousePointer = vbHourglass
Select Case Index
       Case 0:
            If Not IsLogOff Then
                CloseCnn
                If Check1.Value = 1 Then
                   StrCnn = "PROVIDER=MSDataShape;Data Provider=SQLOLEDB.1;Password=;Persist Security Info=True;User ID=sa;Initial Catalog=" & cmbData.Text & ";Data Source=" & cboServer.Text
                Else
                  StrCnn = "PROVIDER=MSDataShape;Data Provider=SQLOLEDB.1;Password=;Persist Security Info=True;User ID=sa;Initial Catalog=" & cmbData.Text & ";Data Source=" & cboServer.Text
                End If
                'StrCnn = "PROVIDER=MSDataShape;Data Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ManTellDB;data Source =" & cboServer.Text
                mVarOpn = OpenCnn(StrCnn)
            End If
            
            aksess.Login txtUser(0), txtUser(1)
            aksess.GetID
            If aksess.GetID = 0 Then
                  MessageBox "Password Anda Salah Atau User Tdk Ada", "Login Gagal", msgOkOnly, msgCrtical
                  Screen.MousePointer = 0
            'If mVarOpn = True Then
            Else
               SaveSetting "Manufacturing Intelligent", "Server", "Server Name", cboServer
               SaveSetting "Manufacturing Intelligent", "Server", "Security", Check1.Value
               SaveSetting "Manufacturing Intelligent", "Server", "UserName Active", txtUser(0)
               SaveSetting "Manufacturing Intelligent", "Server", "Database", cmbData.Text
               
               'Script Sementara ntar diletakkan di procedure UserLogin dimasukkan ke validasi user
               'if user allow access frmSPH & SPPH
               Set frmSPPH = New FrmPurchaseOffer
                  frmSPPH.OperationMode = "SPPH"
               Set frmSPH = New FrmPurchaseOffer
                  frmSPH.OperationMode = "SPH"
               '--------------
               
               IsLoginSucces = True
               mVarPassword = txtUser(1)
'               aksess.Login txtUser(0), txtUser(1)
'               aksess.GetID'
               NoaktifMenu
               aksess.setMenuAktif aksess.GetID
              ' myMenu.CreateMenu "MASTER"  'di gunakan untuk menampilkan menu pertama kali
              
               MainMenu.SemeruTree1.Visible = False
                                            
                             
                                            
               If Not MainMenu Is Nothing Then
                  If txtUser(0) <> "" Then MainMenu.StatusBar1.Panels(1).Text = txtUser(0) Else MainMenu.StatusBar1.Panels(1).Text = "Admin"
                  mVarServerName = GetSetting(App.EXEName, "Lisence Profile", "Servername")
                  MainMenu.StatusBar1.Panels(3).Text = "Server : " & cboServer
                  MainMenu.StatusBar1.Panels(4).Text = "Database : " & cmbData.Text   ' & cboServer
                  mVarLoginActive = MainMenu.StatusBar1.Panels(1).Text
               End If
               MainMenu.Enabled = True
               Unload Me
            'ElseIf mVarOpn = -2147467259 Then
             '  IsLoginSucces = False
            'ElseIf mVarOpn = -2147217843 Then
            '   IsLoginSucces = False
            'Else
             '  IsLoginSucces = False '
            End If
       Case 1:
            'CloseCnn
            Screen.MousePointer = 0
            Unload Me
            Unload MainMenu
            
       Case 2:
            If cmdOk(2).Caption = "&Advance" Then
               cmdOk(2).Caption = "&Simple"
               Label1(4).Visible = True
               Line1(4).Visible = True
               txtUser(2).Visible = True
               CmdLookUp.Visible = True
            Else
               cmdOk(2).Caption = "&Advance"
               Label1(4).Visible = False
               Line1(4).Visible = False
               txtUser(2).Visible = False
               CmdLookUp.Visible = False
            End If
           
End Select
Set mVarOpn = Nothing
Screen.MousePointer = 0
Exit Sub
Hell:
    Screen.MousePointer = 0
    Err.Clear
End Sub


Private Sub CloseCnn()
If Not CNN Is Nothing Then
   If CNN.State = 1 Then
      CNN.Close
   End If
End If
Set CNN = Nothing
End Sub

Private Sub AttachDB()
Dim ofn As OPENFILENAME
Dim I As Integer
 ofn.lStructSize = Len(ofn)
 ofn.hwndOwner = frmLogin.hwnd
 ofn.hInstance = App.hInstance
 ofn.lpstrFilter = "Database Files (*.MDF)" + Chr$(0) + "*.MDF" + Chr$(0) + "Database Files (*.MDF)" + Chr$(0) + "*.MDF" + Chr$(0)
     ofn.lpstrFile = Space$(254)
     ofn.nMaxFile = 255
     ofn.lpstrFileTitle = Space$(254)
     ofn.nMaxFileTitle = 255
     ofn.lpstrInitialDir = CurDir
     ofn.lpstrTitle = "Open Database File"
     ofn.flags = 0
     Dim A As Boolean
     A = GetOpenFileName(ofn)

     If (A) Then
             mVarNamaFile = Left(Trim(ofn.lpstrFileTitle), Len(Trim(ofn.lpstrFileTitle)) - 1)
             
             mVarLokasi = Left(Trim(ofn.lpstrFile), Len(Trim(ofn.lpstrFile)) - 1)
             I = InStr(mVarLokasi, mVarNamaFile)
             If I <> 0 Then
                mVarLokasi = Left(mVarLokasi, I - 1)
             End If
             LetAttachDB
     Else
             mVarNamaFile = ""
     End If
End Sub

Private Sub LetAttachDB()
On Error GoTo Hell
Dim icnn As New Connection
Dim Icomm As New Command
Dim mVarStr  As String
Dim mVarI As String
Dim mVarPos As Integer
mVarI = InStr(mVarNamaFile, "_")
mVarPos = InStr(UCase(mVarNamaFile), ".MDF")
If mVarPos <> 0 Then
    'If mVarI <> 0 Then
       icnn.CursorLocation = adUseClient
       icnn.Mode = adModeShareExclusive
       '"Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=master;Data Source=DEVPC2"
       If Check1.Value = 0 Then
          icnn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & txtUser(1) & ";Persist Security Info=True;User ID=" & txtUser(0) & ";Initial Catalog=master;Data Source=" & cboServer
       Else
          icnn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=" & cboServer
       End If
       
       'icnn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=" & cboServer.Text
       icnn.Open
       mVarI = InStr(mVarNamaFile, "_")
       If mVarI <> 0 Then
          mVarNamaFile = Trim(Left(mVarNamaFile, mVarI - 1))
          mVarStr = " Exec sp_attach_db  N'" & mVarNamaFile & "', " & _
                    " N'" & mVarLokasi & mVarNamaFile & "_Data.mdf'," & _
                    " N'" & mVarLokasi & mVarNamaFile & "_log.ldf'"
       Else
          mVarI = InStr(UCase(mVarNamaFile), ".MDF")
          mVarNamaFile = Left(mVarNamaFile, mVarI - 1)
          mVarStr = " Exec sp_attach_db  N'" & mVarNamaFile & "', " & _
                    " N'" & mVarLokasi & mVarNamaFile & ".mdf'," & _
                    " N'" & mVarLokasi & mVarNamaFile & "_log.ldf'"
       End If
          mVarStr = " Exec sp_attach_db  N'" & mVarNamaFile & "', " & _
                    " N'" & mVarLokasi & mVarNamaFile & ".mdf'" '& _
                   ' " N'" & mVarLokasi & mVarNamaFile & "_log.ldf'"
                   
          mVarStr = " DBCC CHECKTABLE ('Master')"

                   
       With Icomm
            Set .ActiveConnection = icnn
'            messagebox icnn
            .CommandType = adCmdStoredProc
            .CommandText = "Gandeng"
'            .Parameters(0).Value = "manufacture"
'            .Parameters("dbname").Value = "C:\manufacture.Mdf"
            MessageBox mVarStr
            .Execute
       End With
       Set Icomm = Nothing
       icnn.Close
       Set icnn = Nothing
       MessageBox "Seting database ke server sukses." & vbCrLf & "Database sekarang bisa diakses melalui aplikasi.", "Peringatan", msgOkOnly, msgCrtical
Else
   MessageBox "Nama database tidak sesuai.Harap diulangi.", "Peringatan", msgOkOnly, msgCrtical
End If
Exit Sub
Hell:
    MessageBox Err.Description, "Warning", msgOkOnly, msgExclamation
    Set icnn = Nothing
    Set Icomm = Nothing
    Err.Clear
End Sub

Private Function Username() As String
On Error GoTo Hell
Dim Counter As Long
Dim s As String
Dim dl As Long
Counter = 200
s = String(Counter, 0)
dl = GetUserName(s, Counter)
Username = Trim(Left(s, Counter))
'txtUser(0).Text = Username
'StatusBar1.Panels(1).Text = Trim(Left(s, Counter))
Exit Function
Hell:
    Err.Clear
    'txtUser(0).Text = ""
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtUser_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdOk(0).Value = True ' KeyEnter KeyCode
End Sub

'Private Sub OpenToolMenu()
'Dim Rc As New DBQuick
'Rc.DBOpen "SELECT  [Detail User Table].Laporan,[Form Table].[Group Table] FROM [Detail User Table] INNER JOIN [Form Table] ON [Detail User Table].Idx = [Form Table].Idx WHERE     ([Detail User Table].[User ID] = " & mVarIDUser & ") GROUP BY [Detail User Table].Laporan, [Form Table].[Group Table] ORDER BY [Form Table].[Group Table] ", CNN, lckLockReadOnly
'
'With Rc.DBRecordset
'     If .Recordcount <> 0 Then
'        Do
'          If .EOF Then Exit Do
'             With MainMenu.Toolbar1
'                  Select Case Rc.DBRecordset.AbsolutePosition
'                         Case 1: 'ACCOUNTING
'                              .Buttons(7).Visible = CBool(Rc.DBRecordset.Fields(0))
'                              .Buttons(7).Caption = Rc.DBRecordset.Fields(1)
'                              .Buttons(8).Visible = .Buttons(7).Visible
'                              .Buttons(9).Visible = True
'                              .Buttons(10).Visible = True
'                         Case 2: 'DISTRIBUTION
'                              .Buttons(3).Visible = CBool(Rc.DBRecordset.Fields(0))
'                              .Buttons(3).Caption = Rc.DBRecordset.Fields(1)
'                              .Buttons(4).Visible = .Buttons(2).Visible
'                              'MainMenu.mnTRans.Enabled = .Buttons(2).Visible
'                              'MainMenu.mnTRans.Visible = .Buttons(2).Visible
'                         Case 3: 'MASTER DATA
'                              .Buttons(1).Visible = CBool(Rc.DBRecordset.Fields(0))
'                              .Buttons(1).Caption = Rc.DBRecordset.Fields(1)
'                              .Buttons(2).Visible = .Buttons(1).Visible
'                              MainMenu.mnMaster.Visible = .Buttons(1).Visible
'                              MainMenu.mnMaster.Enabled = .Buttons(1).Visible
'                         Case 4: 'PRODUCTION
'                              .Buttons(5).Visible = CBool(Rc.DBRecordset.Fields(0))
'                              .Buttons(5).Caption = Rc.DBRecordset.Fields(1)
'                              .Buttons(6).Visible = .Buttons(5).Visible
'
'
'                  End Select
'                  .Buttons(11).Visible = .Buttons(1).Visible + .Buttons(3).Visible + .Buttons(5).Visible + .Buttons(7).Visible
'                  .Buttons(12).Visible = .Buttons(11).Visible
'                  .Buttons(12).Visible = .Buttons(11).Visible
'                  .Refresh
'             End With
'          .MoveNext
'        Loop
'        .MoveLast
'     End If
'End With
'End Sub


Private Sub NoaktifMenu()
'MainMenu.mnMaster.Enabled = False
'MainMenu.mnPurchase.Enabled = False
'MainMenu.mnMarketing.Enabled = False
'MainMenu.MnGudang.Enabled = False
'MainMenu.mnLogistik.Enabled = False
'MainMenu.mnInventory.Enabled = False
'MainMenu.mnAkun.Enabled = False
'MainMenu.mnQuality.Enabled = False
'MainMenu.mnMaintenance.Enabled = False
'MainMenu.mnHrd.Enabled = False
'
'MainMenu.Toolbar1.Buttons(1).Enabled = False 'master data
'MainMenu.Toolbar1.Buttons(3).Enabled = False ' pembelian
'MainMenu.Toolbar1.Buttons(5).Enabled = False ' penjualan
'MainMenu.Toolbar1.Buttons(7).Enabled = False 'logistik
'MainMenu.Toolbar1.Buttons(9).Enabled = False 'gudang rl
'MainMenu.Toolbar1.Buttons(11).Enabled = False ' production
'MainMenu.Toolbar1.Buttons(13).Enabled = False 'accouting
'MainMenu.Toolbar1.Buttons(15).Enabled = False 'quality
'MainMenu.Toolbar1.Buttons(17).Enabled = False 'hr
'MainMenu.Toolbar1.Buttons(19).Enabled = False 'maintenance

MainMenu.mnMaster.Visible = False
MainMenu.mnPurchase.Visible = False
MainMenu.mnMarketing.Visible = False
MainMenu.MnGudang.Visible = False
MainMenu.mnLogistik.Visible = False
MainMenu.mnInventory.Visible = False
MainMenu.mnAkun.Visible = False
MainMenu.mnQuality.Visible = False
MainMenu.mnMaintenance.Visible = False
MainMenu.mnHrd.Visible = False


MainMenu.Toolbar1.Buttons(1).Visible = False 'master data
MainMenu.Toolbar1.Buttons(2).Visible = False 'garis

MainMenu.Toolbar1.Buttons(3).Visible = False 'pembelian
MainMenu.Toolbar1.Buttons(4).Visible = False 'garis pembelian

MainMenu.Toolbar1.Buttons(5).Visible = False 'penjualan
MainMenu.Toolbar1.Buttons(6).Visible = False 'garis penjualan

MainMenu.Toolbar1.Buttons(7).Visible = False 'logistik
MainMenu.Toolbar1.Buttons(8).Visible = False 'garis logistik

MainMenu.Toolbar1.Buttons(9).Visible = False 'gudang rl
MainMenu.Toolbar1.Buttons(10).Visible = False 'garis gudang rl

MainMenu.Toolbar1.Buttons(11).Visible = False 'production
MainMenu.Toolbar1.Buttons(12).Visible = False 'garis production

MainMenu.Toolbar1.Buttons(13).Visible = False 'accounting
MainMenu.Toolbar1.Buttons(14).Visible = False 'garis accounting

MainMenu.Toolbar1.Buttons(15).Visible = False 'quality
MainMenu.Toolbar1.Buttons(16).Visible = False 'garis quality

MainMenu.Toolbar1.Buttons(17).Visible = False 'hr
MainMenu.Toolbar1.Buttons(18).Visible = False 'garis hr

MainMenu.Toolbar1.Buttons(19).Visible = False 'maintenance
MainMenu.Toolbar1.Buttons(20).Visible = False 'garis maintenance
MainMenu.Toolbar1.Buttons(21).Visible = False 'report

mnBInventoryAdj = False
mnBInventoryBrowser = False

End Sub





