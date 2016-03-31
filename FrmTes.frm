VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmTes 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2220
      Left            =   300
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   255
      Width           =   4080
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   3285
      Top             =   2775
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AgroKomoditi;Data Source=BULIRCOMP"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AgroKomoditi;Data Source=BULIRCOMP"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1380
      Left            =   525
      TabIndex        =   0
      Top             =   3645
      Width           =   2460
   End
End
Attribute VB_Name = "FrmTes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


Public Sub SettingReport(ByVal CommandProcedure As String, ByVal CommandProcedureName As String, ByVal ConnectionString As String)
On Error Resume Next
Dim Icom As New Command
Dim Iconn As New Connection
Iconn.CursorLocation = adUseClient
Iconn.Mode = adModeShareExclusive
Iconn.IsolationLevel = adXactIsolated
Iconn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AgroKomoditi;Data Source=BULIRCOMP"
Iconn.Open
Set Icom.ActiveConnection = Iconn
With Icom
     .CommandType = adCmdUnknown
     .CommandText = "Drop View " & CommandProcedureName
     .Execute
     .CommandText = "Create View " & CommandProcedureName & " as " & CommandProcedure
     .Execute
End With
Hell:
    Set icomm = Nothing
    Iconn.Close
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Warning"
    Err.Clear
End Sub
