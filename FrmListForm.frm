VERSION 5.00
Begin VB.Form FrmListForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Form"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmListForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4050
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
      Left            =   2775
      Picture         =   "FrmListForm.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4410
      Width           =   1230
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      ItemData        =   "FrmListForm.frx":0C54
      Left            =   15
      List            =   "FrmListForm.frx":0CBE
      TabIndex        =   0
      Top             =   0
      Width           =   4005
   End
End
Attribute VB_Name = "FrmListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'      Private Const GW_CHILD = 5
'      Private Const GW_HWNDNEXT = 2
'
'    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
'              ByVal wCmd As Long) As Long
'    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
'              (ByVal hwnd As Long, ByVal lpString As String, _
'              ByVal cch As Long) As Long
'    Private Declare Function GetTopWindow Lib "user32" _
'              (ByVal hwnd As Long) As Long
'    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
'              (ByVal hwnd As Long, ByVal lpClassName As String, _
'              ByVal nMaxCount As Long) As Long
'
'Private Sub AddChildWindows(ByVal hwndParent As Long, ByVal Level As Long)
'Dim WT As String, CN As String, Length As Long, hwnd As Long
'  If Level = 0 Then
'    hwnd = hwndParent
'  Else
'    hwnd = GetWindow(hwndParent, GW_CHILD)
'  End If
'  Do While hwnd <> 0
'    WT = Space(256)
'    Length = GetWindowText(hwnd, WT, 255)
'    WT = Left$(WT, Length)
'    CN = Space(256)
'    Length = GetClassName(hwnd, CN, 255)
'    CN = Left$(CN, Length)
'    List1.AddItem String(2 * Level, ".") & WT & " (" & CN & ")"
'    AddChildWindows hwnd, Level + 1
'    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
'  Loop
'End Sub


Private Sub CmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Dim hwnd As Long
'hwnd = GetWindow(MainMenu.hwnd, GW_CHILD)
'AddChildWindows MainMenu.hwnd, 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmListForm = Nothing
End Sub
