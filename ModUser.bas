Attribute VB_Name = "ModUser"
Option Explicit

Dim Rc As New Recordset

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
Public Function GetUsers(UserNames() As String, _
   Optional ServerName As String = "") As Boolean
    
    'PURPOSE: Get LoginNames of all users on the domain and
    'save in a string array
    
    'PARAMETERS: UserNames(): Empty String array, passed byref,
      'to hold user names
    
    'ServerName (Optional): Set to "" if you want user
      'names for current machine, otherwise, set to the server
      'you want (e.g., Domain Controller Name)
    
    'RETURNS: True if successful, false otherwise
    
    'EXAMPLE:
        'Dim sUsers() As String
        'dim lCtr as long
        'GetUsers sUsers, "MyDomainController"
        
        'OR FOR LOCAL MACHINE
        
        'GetUsers sUsers
   
    'For lCtr = LBound(sUsers) To UBound(sUsers)
     '   Debug.Print sUsers(lCtr)
    'Next
    
     'NOTES: WINDOWS NT/2000 only
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
    Dim i As Long
    
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

         For i = 0 To lUsersRead - 1
           CopyMem etUserInfo, ByVal lptrStrBuffer + Len(etUserInfo) * i, _
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

'Public Sub LookupUser(Byval )
'Dim sUsers() As String
'Dim lctr
'With Rc
'     .Fields.Append "User Domain", adBSTR
'     .Open
'   GetUsers sUsers, "BULIRSERVER"
'   For lctr = LBound(sUsers) To UBound(sUsers)
'        .AddNew 0, sUsers(lctr)
'   Next
'   Set DataGrid1.DataSource = Rc
'End With
'End Sub



