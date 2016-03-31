Attribute VB_Name = "patchLib"
'Here is the module code:

' © 2005 Kevin M. Jones

Option Explicit

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000

Private Type tFileTime
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type tSystemTime
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Declare Function SetFileTime Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpCreatedTime As tFileTime, _
      lpAccessedTime As tFileTime, _
      lpLastWriteTime As tFileTime _
   ) As Long
   
Private Declare Function GetFileTime Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpCreatedTime As tFileTime, _
      lpAccessedTime As tFileTime, _
      lpLastWriteTime As tFileTime _
   ) As Long
   
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
      ByVal lpFileName As String, _
      ByVal dwDesiredAccess As Long, _
      ByVal dwShareMode As Long, _
      ByVal lpSecurityAttributes As Long, _
      ByVal dwCreatedDisposition As Long, _
      ByVal dwFlagsAndAttributes As Long, _
      ByVal hTemplateFile As Long _
   ) As Long
   
Private Declare Function CloseHandle Lib "kernel32" ( _
      ByVal hObject As Long _
   ) As Long
   
Private Declare Function SystemTimeToFileTime Lib "kernel32" ( _
      lpSystemTime As tSystemTime, _
      lpFileTime As tFileTime _
   ) As Long
   
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" ( _
      lpLocalFileTime As tFileTime, _
      lpFileTime As tFileTime _
   ) As Long
   
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
      lpFileTime As tFileTime, _
      lpLocalFileTime As tFileTime _
   ) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" ( _
      lpFileTime As tFileTime, _
      lpSystemTime As tSystemTime _
   ) As Long

Private Sub ConvertVBTimeToFileTime( _
      ByVal VBTime As Date, _
      ByRef FileTime As tFileTime _
   )
   
' Converts a VB date/time value to a GMT file time value.

   Dim SystemTime As tSystemTime
   Dim LocalTime As tFileTime
   
   ' Load system time structure
   SystemTime.wYear = Year(VBTime)
   SystemTime.wMonth = Month(VBTime)
   SystemTime.wDay = Day(VBTime)
   SystemTime.wDayOfWeek = Weekday(VBTime) - 1
   SystemTime.wHour = Hour(VBTime)
   SystemTime.wMinute = Minute(VBTime)
   SystemTime.wSecond = Second(VBTime)
   SystemTime.wMilliseconds = 0
   
   ' Convert system time format to local time format
   SystemTimeToFileTime SystemTime, LocalTime
   
   ' Convert local time to GMT
   LocalFileTimeToFileTime LocalTime, FileTime
   
End Sub

Private Sub ConvertFileTimeToVBTime( _
      ByRef FileTime As tFileTime, _
      ByRef VBTime As Date _
   )
   
' Converts a GMT file time value to a VB date/time value.

   Dim SystemTime As tSystemTime
   Dim LocalTime As tFileTime
   
   ' Convert GMT to local time
   FileTimeToLocalFileTime FileTime, LocalTime
   
   ' Convert local time format to system time format
   FileTimeToSystemTime LocalTime, SystemTime
   
   ' Pull VB time from system time structure
   VBTime = DateSerial(SystemTime.wYear, SystemTime.wMonth, SystemTime.wDay) + TimeSerial(SystemTime.wHour, SystemTime.wMinute, SystemTime.wSecond)
   
End Sub

Public Function GetFileTimes( _
      ByVal FilePath As String, _
      ByRef CreatedTime As Date, _
      ByRef AccessedTime As Date, _
      ByRef ModifiedTime As Date _
   ) As Boolean

' Gets the created, last accessed, and/or modified times for the specified file.
   
   Dim FileHandle As Long
   Dim CreatedFileTime As tFileTime
   Dim AccessedFileTime As tFileTime
   Dim ModifiedFileTime As tFileTime
   Dim LocalTime As tFileTime
   Dim SystemTime As tSystemTime
   Dim Result As Long
   
   ' Open file object
   FileHandle = CreateFile(FilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
   If FileHandle = 0 Then Exit Function
   
   ' Get file times
   Result = GetFileTime(FileHandle, CreatedFileTime, AccessedFileTime, ModifiedFileTime)
   If Result = 0 Then Exit Function
   
   ConvertFileTimeToVBTime CreatedFileTime, CreatedTime
   ConvertFileTimeToVBTime AccessedFileTime, AccessedTime
   ConvertFileTimeToVBTime ModifiedFileTime, ModifiedTime
   
   ' Close file
   CloseHandle FileHandle
   
   ' Return success
   GetFileTimes = True
   
End Function

Public Function SetFileTimes( _
      ByVal FilePath As String, _
      Optional ByVal CreatedTime As Date, _
      Optional ByVal AccessedTime As Date, _
      Optional ByVal ModifiedTime As Date _
   ) As Boolean

' Sets the created, last accessed, and/or modified times for the specified file.
   
   Dim FileHandle As Long
   Dim CreatedFileTime As tFileTime
   Dim AccessedFileTime As tFileTime
   Dim ModifiedFileTime As tFileTime
   Dim LocalTime As tFileTime
   Dim SystemTime As tSystemTime
   Dim Result As Long
   
   ' Open file
   FileHandle = CreateFile(FilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
   If FileHandle = 0 Then Exit Function
   
   ' Get file times
   Result = GetFileTime(FileHandle, CreatedFileTime, AccessedFileTime, ModifiedFileTime)
   If Result = 0 Then Exit Function
   
   If CreatedTime > 0 Then
      ConvertVBTimeToFileTime CreatedTime, CreatedFileTime
   End If
   
   If AccessedTime > 0 Then
      ConvertVBTimeToFileTime AccessedTime, AccessedFileTime
   End If
   
   If ModifiedTime > 0 Then
      ConvertVBTimeToFileTime ModifiedTime, ModifiedFileTime
   End If
   
   ' Set file times
   Result = SetFileTime(FileHandle, CreatedFileTime, AccessedFileTime, ModifiedFileTime)
   If Result = 0 Then Exit Function
   
   ' Close file
   CloseHandle FileHandle
   
   ' Return success
   SetFileTimes = True
   
End Function



