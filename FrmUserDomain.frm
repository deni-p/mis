VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUserDomain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Domain"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUserDomain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5190
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   9155
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16577005
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "User Domain"
         Caption         =   "User Domain"
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
            ColumnWidth     =   4500.284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4050
      Picture         =   "FrmUserDomain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5295
      Width           =   1230
   End
End
Attribute VB_Name = "FrmUserDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Rc As Recordset
Attribute Rc.VB_VarHelpID = -1

Private Type USER_INFO
    Name As String
    Comment As String
    UserComment As String
    FullName As String
End Type

Private Type USER_INFO_API
    Name As Long
    Comment As Long
    UserComment As Long
    FullName As Long
End Type


Private Declare Function NetUserEnum Lib "netapi32" _
  (lpServer As Any, ByVal Level As Long, _
   ByVal Filter As Long, lpBuffer As Long, _
   ByVal PrefMaxLen As Long, EntriesRead As Long, _
   TotalEntries As Long, ResumeHandle As Long) As Long
   
Private Declare Function NetApiBufferFree Lib "netapi32" _
   (ByVal pBuffer As Long) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias _
   "RtlMoveMemory" (pTo As Any, uFrom As Any, _
    ByVal lSize As Long)
    
Private Declare Function lstrlenW Lib "kernel32" _
 (ByVal lpString As Long) As Long

Private Const NERR_Success As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&

Private Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Private Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Private Const FILTER_PROXY_ACCOUNT As Long = &H4&
Private Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Private Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Private Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&

Private Function GetUsers(UserNames() As String, _
   Optional ServerName As String = "") As Boolean
    
     Dim lptrStrBuffer As Long
    Dim lRet As Long
    Dim lUsersRead As Long
    Dim lTotalUsers As Long
    Dim lHnd As Long
    Dim etUserInfo As USER_INFO_API
    Dim bytServerName() As Byte
    Dim lElement As Long
    Dim Users() As USER_INFO 'This function
    'is designed to return a string of username
    'but optionally, you can change it to
    'get this array of the UDT, which
    'will provide more information
    'about each user
    Dim I As Long
    
    ReDim Users(0) As USER_INFO
    ReDim UserNames(0) As String
    
    If Trim(ServerName) = "" Then
        'Local users
        bytServerName = vbNullString
    Else
        'Check the syntax of the ServerName string
        If InStr(ServerName, "\\") = 1 Then
            bytServerName = ServerName & vbNullChar
        Else
            bytServerName = "\\" & ServerName & vbNullChar
        End If
    End If
    lHnd = 0

 Do
         'Begin enumerating users
         If Trim(ServerName) = "" Then
             lRet = NetUserEnum(vbNullString, 10, _
              FILTER_NORMAL_ACCOUNT, lptrStrBuffer, 1, _
               lUsersRead, lTotalUsers, lHnd)
         Else
             lRet = NetUserEnum(bytServerName(0), 10, _
              FILTER_NORMAL_ACCOUNT, lptrStrBuffer, 1, _
                lUsersRead, lTotalUsers, lHnd)
         End If

         'Populate UserInfo Structure
         'If lRet = ERROR_MORE_DATA Then

         '  If lUsersRead  1 that why th for construct

         For I = 0 To lUsersRead - 1
           CopyMem etUserInfo, ByVal lptrStrBuffer + Len(etUserInfo) * I, _
 Len(etUserInfo)
           If Users(0).Name = "" Then
               lElement = 0
           Else
               lElement = UBound(Users) + 1
           End If
           'ReDim Preserve UserNames(lElement)
           ReDim Preserve Users(lElement) As USER_INFO

           'data of interest
           Users(lElement).Name = PtrToString(etUserInfo.Name)

 'If lRet = ERROR_MORE_DATA Then --  i removed because i lost the last
'entry while the result is NERR_Success

           'Other stuff you can get, but not
           'returned by this function
           'modify this function if you are interested

           Users(lElement).Comment = PtrToString(etUserInfo.Comment)
           Users(lElement).UserComment = PtrToString(etUserInfo.UserComment)
           Users(lElement).FullName = PtrToString(etUserInfo.FullName)
            ReDim Preserve UserNames(lElement)
           UserNames(lElement) = Users(lElement).Name
         Next

         If lptrStrBuffer Then
             Call NetApiBufferFree(lptrStrBuffer)
         End If
         DoEvents
         If lRet = NERR_Success Then Exit Do
     Loop While lRet = ERROR_MORE_DATA
 GetUsers = True
    Exit Function
ErrHandler:
On Error Resume Next
Call NetApiBufferFree(lptrStrBuffer)
End Function

Private Function PtrToString(lpString As Long) As String
    'Convert a windows pointer to a VB string
    Dim bytBuffer() As Byte
    Dim lLen As Long
    
    If lpString Then
        lLen = lstrlenW(lpString) * 2
        If lLen Then
            ReDim bytBuffer(0 To (lLen - 1)) As Byte
            CopyMem bytBuffer(0), ByVal lpString, lLen
            PtrToString = bytBuffer
        End If
    End If
End Function

Public Sub LookupUser()
On Error GoTo Hell
Dim sUsers() As String
Dim lctr
If Not Rc Is Nothing Then
   If Rc.State = 1 Then Rc.Close
End If
Set Rc = Nothing
Set Rc = New Recordset
With Rc
     .Fields.Append "User Domain", adBSTR
     .Open
     GetUsers sUsers, "bulirserver"
     For lctr = LBound(sUsers) To UBound(sUsers)
         .AddNew 0, sUsers(lctr)
     Next
     Set DataGrid1.DataSource = Rc
End With
Hell:
    If Err.Number <> 0 Then MessageBox Err.Number & vbCrLf & Err.Description, "Warning"
End Sub

Private Sub CmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
LookupUser
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not Rc Is Nothing Then
   If Rc.State = 1 Then Rc.Close
End If
Set Rc = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmUserDomain = Nothing
End Sub
