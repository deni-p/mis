Attribute VB_Name = "Patch"
Private Declare Function CopyFile Lib "kernel32" Alias _
"CopyFileA" (ByVal lpExistingFileName As String, ByVal _
lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias _
"MoveFileA" (ByVal lpExistingFileName As String, ByVal _
lpNewFileName As String) As Long

Dim NameSource As String
Dim PathSource As String
Dim strSource As String
Dim NameTarget As String
Dim PathTarget As String
Dim strTarget As String
Dim lngRetVal As Long

Const VS_FFI_SIGNATURE = &HFEEF04BD
Const VS_FFI_STRUCVERSION = &H10000
Const VS_FFI_FILEFLAGSMASK = &H3F&
Const VS_FF_DEBUG = &H1
Const VS_FF_PRERELEASE = &H2
Const VS_FF_PATCHED = &H4
Const VS_FF_PRIVATEBUILD = &H8
Const VS_FF_INFOINFERRED = &H10
Const VS_FF_SPECIALBUILD = &H20
Const VOS_UNKNOWN = &H0
Const VOS_DOS = &H10000
Const VOS_OS216 = &H20000
Const VOS_OS232 = &H30000
Const VOS_NT = &H40000
Const VOS__BASE = &H0
Const VOS__WINDOWS16 = &H1
Const VOS__PM16 = &H2
Const VOS__PM32 = &H3
Const VOS__WINDOWS32 = &H4
Const VOS_DOS_WINDOWS16 = &H10001
Const VOS_DOS_WINDOWS32 = &H10004
Const VOS_OS216_PM16 = &H20002
Const VOS_OS232_PM32 = &H30003
Const VOS_NT_WINDOWS32 = &H40004
Const VFT_UNKNOWN = &H0
Const VFT_APP = &H1
Const VFT_DLL = &H2
Const VFT_DRV = &H3
Const VFT_FONT = &H4
Const VFT_VXD = &H5
Const VFT_STATIC_LIB = &H7
Const VFT2_UNKNOWN = &H0
Const VFT2_DRV_PRINTER = &H1
Const VFT2_DRV_KEYBOARD = &H2
Const VFT2_DRV_LANGUAGE = &H3
Const VFT2_DRV_DISPLAY = &H4
Const VFT2_DRV_MOUSE = &H5
Const VFT2_DRV_NETWORK = &H6
Const VFT2_DRV_SYSTEM = &H7
Const VFT2_DRV_INSTALLABLE = &H8
Const VFT2_DRV_SOUND = &H9
Const VFT2_DRV_COMM = &HA

Private Type VS_FIXEDFILEINFO
dwSignature As Long
dwStrucVersionl As Integer ' e.g. = &h0000 = 0
dwStrucVersionh As Integer ' e.g. = &h0042 = .42
dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
dwFileFlagsMask As Long ' = &h3F for version "0.42"
dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
dwFileType As Long ' e.g. VFT_DRIVER
dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
dwFileDateMS As Long ' e.g. 0
dwFileDateLS As Long ' e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Dim Filename As String, Directory As String, FullFileName As String
Dim StrucVer As String, FileVer As String, ProdVer As String
Dim FileFlags As String, FileOS As String, FileType As String, FileSubType As String

Private Sub DisplayVerInfo(FullFileName As String)
Dim rc As Long, lDummy As Long, sBuffer() As Byte
Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
Dim lVerbufferLen As Long

'*** Get size ****
lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
If lBufferLen < 1 Then
MsgBox "No Version Info available!"
Exit Sub
End If

'**** Store info to udtVerBuffer struct ****
ReDim sBuffer(lBufferLen)
rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
rc = VerQueryValue(sBuffer(0), "\\", lVerPointer, lVerbufferLen)
MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

'**** Determine Structure Version number - NOT USED ****
StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)

'**** Determine File Version number ****
FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)

'**** Determine Product Version number ****
ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)

'**** Determine Boolean attributes of File ****
FileFlags = ""
If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then FileFlags = "Debug "
If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE Then FileFlags = FileFlags & "PreRel "
If udtVerBuffer.dwFileFlags And VS_FF_PATCHED Then FileFlags = FileFlags & "Patched "
If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD Then FileFlags = FileFlags & "Private "
If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRE Then FileFlags = FileFlags & "Info "
If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD Then FileFlags = FileFlags & "Special "
If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN Then FileFlags = FileFlags + "Unknown "

'**** Determine OS for which file was designed ****
Select Case udtVerBuffer.dwFileOS
Case VOS_DOS_WINDOWS16
    FileOS = "DOS-Win16"
Case VOS_DOS_WINDOWS32
    FileOS = "DOS-Win32"
Case VOS_OS216_PM16
    FileOS = "OS/2-16 PM-16"
Case VOS_OS232_PM32
    FileOS = "OS/2-16 PM-32"
Case VOS_NT_WINDOWS32
    FileOS = "NT-Win32"
Case other
    FileOS = "Unknown"
End Select

Select Case udtVerBuffer.dwFileType
Case VFT_APP
    FileType = "App"
Case VFT_DLL
    FileType = "DLL"
Case VFT_DRV
    FileType = "Driver"
    
    Select Case udtVerBuffer.dwFileSubtype
    Case VFT2_DRV_PRINTER
        FileSubType = "Printer drv"
    Case VFT2_DRV_KEYBOARD
        FileSubType = "Keyboard drv"
    Case VFT2_DRV_LANGUAGE
        FileSubType = "Language drv"
    Case VFT2_DRV_DISPLAY
        FileSubType = "Display drv"
    Case VFT2_DRV_MOUSE
        FileSubType = "Mouse drv"
    Case VFT2_DRV_NETWORK
        FileSubType = "Network drv"
    Case VFT2_DRV_SYSTEM
        FileSubType = "System drv"
    Case VFT2_DRV_INSTALLABLE
        FileSubType = "Installable"
    Case VFT2_DRV_SOUND
        FileSubType = "Sound drv"
    Case VFT2_DRV_COMM
        FileSubType = "Comm drv"
    Case VFT2_UNKNOWN
        FileSubType = "Unknown"
    End Select
    
Case VFT_FONT
    FileType = "Font"
        
    Select Case udtVerBuffer.dwFileSubtype
    Case VFT_FONT_RASTER
        FileSubType = "Raster Font"
    Case VFT_FONT_VECTOR
        FileSubType = "Vector Font"
    Case VFT_FONT_TRUETYPE
        FileSubType = "TrueType Font"
    End Select
    
Case VFT_VXD
    FileType = "VxD"
Case VFT_STATIC_LIB
    FileType = "Lib"
Case Else
    FileType = "Unknown"
End Select
End Sub

Private Sub Command1_Click()

''strSource = "C:\Myfile.txt"
'NameSource = App.EXEName & ".exe"
''PathSource = "\\server\Mantell\EXECUTABLE\"
'PathSource = CekChar(GetSetting("Mantell", "LOGIN", "PATHSERVER")) & "\"
'strSource = PathSource & NameSource
'
''strTarget = "D:\MyFolder\Myfile.txt"
'NameTarget = App.EXEName & ".exe"
''PathTarget = "C:\Program Files\Mantell\"
'PathTarget = App.Path & "\"
'strTarget = PathTarget & NameTarget

strSource = Label3.Caption
FileVer = ""
DisplayVerInfo (strSource)
Label5.Caption = FileVer
strTarget = Label4.Caption
FileVer = ""
DisplayVerInfo (strTarget)
Label6.Caption = FileVer

If Label5.Caption > Label6.Caption Then
    If FileExist(strTarget) = True Then
        ChangeName (strTarget)
    End If
    
    '// Copy File
    lngRetVal = CopyFile(Trim$(strSource), Trim(strTarget), True)

Else
    MsgBox "Cancel Update File..."
    Exit Sub
End If

'If FileExist(strTarget) = True Then
'    ChangeName (strTarget)
'End If


If lngRetVal Then
    MsgBox "File copied!"
    MsgBox "Sucessfull Update File..."
Else
    MsgBox "Error. File not copied!"
    MsgBox "Failed Update File..."
End If
End Sub

Private Function ChangeName(PathInFile As String) As String
Dim CreatedFileTime As Date
Dim AccessedFileTime As Date
Dim ModifiedFileTime As Date
Dim DateModified As String

GetFileTimes PathInFile, CreatedFileTime, AccessedFileTime, ModifiedFileTime
DateModified = Day(ModifiedFileTime) & Month(ModifiedFileTime) & Year(ModifiedFileTime) & " " & strTime(ModifiedFileTime)
'MsgBox "Folder created " & CreatedFileTime

Name PathInFile As Left(PathInFile, Len(PathInFile) - Len(Ekstensi(PathInFile))) & " " & DateModified & Ekstensi(PathInFile)
'FileCopy PathInFile, Left(PathInFile, Len(PathInFile) - Len(Ekstensi(PathInFile))) & " " & DateModified & Ekstensi(PathInFile)
'Kill PathInFile
End Function

Private Function FileExist(PathFile As String) As Boolean
Dim LengthFile As Long
    On Error GoTo ErrorHandler
    LengthFile = FileLen(PathFile)
    Call FileLen(PathFile): FileExist = True: Exit Function
ErrorHandler:
    FileExist = False
End Function

Private Function Ekstensi(NameInFile As String) As String
Dim CountExt As Byte
Dim N As Integer
    CountExt = 0
    For N = Len(NameInFile) To 1 Step -1
        If Mid(NameInFile, N, 1) <> "." Then
            CountExt = CountExt + 1
        Else
            N = 1
        End If
    Next N
    Ekstensi = "." + Right(NameInFile, CountExt)
End Function

Private Function strTime(inDate As Date) As String
Dim Jam, Menit, Detik, TZone As String
Dim strWaktu, fixWaktu As String
strWaktu = inDate

strWaktu = Right(strWaktu, Len(strWaktu) - (Len(Format(inDate, "dd/MM/yyyy")) + Len(" ")))
fixWaktu = ""
For I = 1 To Len(strWaktu)
    If Mid(strWaktu, I, 1) <> ":" Then
        If Mid(strWaktu, I, 1) <> " " Then
            fixWaktu = fixWaktu & Mid(strWaktu, I, 1)
        End If
    End If
Next I
strTime = fixWaktu
End Function

Private Function CekChar(VIn)
Dim d As String
    If IsNull(VIn) Or VIn = "" Then
       d = ""
    Else
       d = VIn
    End If
    CekChar = d
End Function


