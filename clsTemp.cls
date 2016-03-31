VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As Guid) As Long
Private Const S_OK = 0

Public Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
   FileExist = (Dir(Fname) <> "")
End Function

'Sub Encrypt(secret$, PassWord$)
'   ' secret$ = the string you wish to encrypt or decrypt.
'   ' PassWord$ = the password with which to encrypt the string.
'   L = Len(PassWord$)
'   For X = 1 To Len(secret$)
'      Char = Asc(Mid$(PassWord$, (X Mod L) - L * ((X Mod L) = 0), 1))
'      Mid$(secret$, X, 1) = Chr$(Asc(Mid$(secret$, X, 1)) Xor Char)
'   Next
'End Sub

Public Function BukafilePath(Filename As String) As String
  Dim NamaFile As String
  Open Filename For Input As #1
     Input #1, NamaFile
     BukafilePath = NamaFile
  Close #1
End Function

Public Function FlashRemuse(ByVal InlineString As String, ByVal KeyRemuse As String) As String
FlashRemuse = EncryptString(InlineString, KeyRemuse)
FlashRemuse = EncryptString(InlineString, KeyRemuse)
FlashRemuse = EncryptString(InlineString, KeyRemuse)
End Function

Public Function SlowRemuse(ByVal InlineString As String, ByVal KeyRemuse As String) As String
SlowRemuse = DecryptString(InlineString, KeyRemuse)
SlowRemuse = DecryptString(InlineString, KeyRemuse)
SlowRemuse = DecryptString(InlineString, KeyRemuse)
End Function

Public Sub SimpanfilePath(Filename, ByVal mykey As String, ByVal TanggalAct As String)
  Open Filename For Output As #1
     Write #1, FlashRemuse(ISIDB & "$" & TanggalAct, mykey)
  Close #1
End Sub
Private Function EncryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey As String
 Dim OutString As String
 Dim OldChar As Long
 Dim NewChar As Long
 Dim CryptChar As Long
 Dim I As Integer
 
 I = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 Do
  I = I + 1
  OldChar = Asc(Mid(InString, I, 1))
  CryptChar = Asc(Mid(TempKey, I, 1))
  Select Case I Mod 2
   Case 0
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
   Case Else
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
  End Select
  If NewChar < 35 Then
   OutString = OutString + "!" + Chr(NewChar + 40)
  Else
   OutString = OutString + Chr(NewChar)
  End If
 Loop Until I = Len(InString)
 
 EncryptString = OutString

End Function

Private Function DecryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey As String
 Dim OutString As String
 Dim OldChar As Long
 Dim NewChar As Long
 Dim CryptChar As Long
 Dim I As Integer
 Dim C As Integer
 On Error GoTo Hell
 C = 0
 I = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 Do
  I = I + 1
  C = C + 1
  OldChar = Asc(Mid(InString, C, 1))
  If OldChar = 33 Then
   C = C + 1
   OldChar = Asc(Mid(InString, C, 1))
   OldChar = OldChar - 40
  End If
  CryptChar = Asc(Mid(TempKey, I, 1))
  Select Case I Mod 2
   Case 0
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
   Case Else
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
  End Select
  OutString = OutString + Chr(NewChar)
 Loop Until C = Len(InString)
 
 DecryptString = OutString
Hell:
    Err.Clear
End Function



Private Function EncryptStr(ByVal OrigStr As String, ByVal Key As String) As String
Dim Counter As Integer
Dim EncCode As Integer
    EncCode = CreateEncryptCode(Key)
    EncryptStr = ""
    For Counter = 1 To Len(OrigStr)
        Select Case Counter Mod 2
            Case 0
                EncryptStr = EncryptStr + Chr(Asc(Mid(OrigStr, Counter, 1)) + EncCode)
            Case Else
                EncryptStr = EncryptStr + Chr(Asc(Mid(OrigStr, Counter, 1)) - EncCode)
        End Select
    Next Counter
End Function


Private Function DecryptStr(ByVal EncStr As String, ByVal Key As String) As String
Dim Counter As Integer
Dim EncCode As Integer

    EncCode = CreateEncryptCode(Key)
    DecryptStr = ""
    For Counter = 1 To Len(EncStr)
        Select Case Counter Mod 2
            Case 0
                DecryptStr = DecryptStr + Chr(Asc(Mid(EncStr, Counter, 1)) - EncCode)
            Case Else
                DecryptStr = DecryptStr + Chr(Asc(Mid(EncStr, Counter, 1)) + EncCode)
        End Select
    Next Counter
End Function


Private Function CreateEncryptCode(Key As String) As Integer
Dim Total(1 To 2) As Integer
Dim NbChars(1 To 2) As Integer
Dim Counter As Integer

    Total(1) = 0
    Total(2) = 0
    NbChars(1) = 0
    NbChars(2) = 0
    For Counter = 1 To Len(Key)
        Select Case Counter Mod 2
            Case 0
                Total(1) = Total(1) + Asc(Mid(Key, Counter, 1))
                NbChars(1) = NbChars(1) + 1
            Case Else
                Total(2) = Total(2) + Asc(Mid(Key, Counter, 1))
                NbChars(2) = NbChars(2) + 1
        End Select
    Next Counter
    If NbChars(1) > 0 And NbChars(2) > 0 Then
        CreateEncryptCode = Abs((Total(1) / NbChars(1)) _
                              - (Total(2) / NbChars(2)))
    Else
        CreateEncryptCode = 1
    End If
End Function

Public Function GetGUID() As String
    
    Dim lResult As Long
    Dim lguid As GUID
    Dim MyguidString As String
    Dim MyGuidString1 As String
    Dim MyGuidString2 As String
    Dim MyGuidString3 As String
    Dim DataLen As Integer
    Dim StringLen As Integer
    Dim I%
    
    On Error GoTo error_olemsg
    
    'lResult = CoCreateGuid(lguid)
    
    If lResult = S_OK Then
        MyGuidString1 = Hex$(lguid.Data1)
        StringLen = Len(MyGuidString1)
        DataLen = Len(lguid.Data1)
        MyGuidString1 = LeadingZeros(2 * DataLen, StringLen) & MyGuidString1        'First 4 bytes (8 hex digits)
        
        MyGuidString2 = Hex$(lguid.Data2)
        StringLen = Len(MyGuidString2)
        DataLen = Len(lguid.Data2)
        MyGuidString2 = LeadingZeros(2 * DataLen, StringLen) & Trim$(MyGuidString2)        'Next 2 bytes (4 hex digits)
        
        MyGuidString3 = Hex$(lguid.Data3)
        StringLen = Len(MyGuidString3)
        DataLen = Len(lguid.Data3)
        MyGuidString3 = LeadingZeros(2 * DataLen, StringLen) & Trim$(MyGuidString3)        'Next 2 bytes (4 hex digits)
        
        GetGUID = MyGuidString1 & MyGuidString2 & MyGuidString3
        
        For I% = 0 To 7
           MyguidString = MyguidString & Format$(Hex$(lguid.Data4(I%)), "00")
        Next I%
        GetGUID = GetGUID & MyguidString
    Else
        GetGUID = "00000000"
    End If
    
    Exit Function

error_olemsg:
         MsgBox "Error " & Str(Err) & ": " & Error$(Err)
         GetGUID = "00000000"
         Exit Function

End Function

Private Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
    LeadingZeros = String$(ExpectedLen - ActualLen, "0")
End Function

Private Function ISIDB() As String
ISIDB = "[Package|Standard Setup Package 1|Files In Project]" & _
"C:\WINNT\system32\MSVBVM60.DLL=Yes" & _
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\OLEAUT32.DLL=Yes" & _
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\OLEPRO32.DLL=Yes" & _
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\ASYCFILT.DLL=Yes" & _
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\STDOLE2.TLB=Yes" & _
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\COMCAT.DLL=Yes" & _
"C:\Program Files\Common Files\System\ado\msado25.tlb=Yes" & _
"C:\WINNT\system32\MSBIND.DLL=Yes" & _
"C:\WINNT\system32\msstdfmt.dll=Yes" & _
"C:\Program Files\Seagate Software\Viewers\ActiveXViewer\crviewer.dll=Yes" & _
"C:\Program Files\Seagate Software\Report Designer Component\craxdrt.dll=Yes" & _
"C:\Program Files\Seagate Software\Report Designer Component\craxddrt.dll=Yes" & _
"C:\WINNT\system32\Implode.dll=Yes" & _
"C:\WINNT\system32\Crpaig80.dll=Yes" & _
"C:\WINNT\system32\msvcrt.dll=Yes" & _
"C:\WINNT\system32\mfc42.dll=Yes" & _
"C:\WINNT\system32\P2smon.dll=Yes" & _
"C:\WINNT\system32\MSCOMCTL.OCX=Yes" & _
"C:\WINNT\system32\mscomct2.ocx=Yes" & _
"C:\WINNT\system32\comct332.ocx=Yes"
End Function

