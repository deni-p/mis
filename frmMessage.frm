VERSION 5.00
Begin VB.Form frmMessage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   6300
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
   Icon            =   "frmMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOK 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   5010
      TabIndex        =   2
      Top             =   2805
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   3795
      TabIndex        =   1
      Top             =   2805
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   2580
      TabIndex        =   0
      Top             =   2805
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   870
      Width           =   6165
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1170
      TabIndex        =   4
      Top             =   300
      Width           =   855
   End
   Begin VB.Image ImInfo 
      Height          =   720
      Left            =   100
      Picture         =   "frmMessage.frx":6852
      Top             =   100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImQuestion 
      Height          =   720
      Left            =   100
      Picture         =   "frmMessage.frx":6E5B
      Top             =   100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImExclamation 
      Height          =   720
      Left            =   100
      Picture         =   "frmMessage.frx":756A
      Top             =   100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImCritical 
      Height          =   720
      Left            =   100
      Picture         =   "frmMessage.frx":7B59
      Top             =   100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "frmMessage.frx":80E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6300
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean

Private Sub cmdOk_Click(Index As Integer)
Select Case Index
       Case 0:
       Case 1: KuBox = 1
       Case 2: KuBox = 2
End Select
Unload Me
End Sub

Private Sub cmdOk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmdOk_Click(Index)
End Sub

Private Sub Form_Activate()
'txtMessage.Enabled = True
If CmdOK(0).Visible = True Then
   CmdOK(0).SetFocus
ElseIf CmdOK(1).Visible = True Then
   CmdOK(1).SetFocus
ElseIf CmdOK(2).Visible = True Then
   CmdOK(2).SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)

'Me.Caption = "Warning " & lblMessage
'Me.Caption = lblMessage
StayOnTop Me
'txtMessage.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ReleaseTop Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMessage = Nothing
End Sub

Private Sub txtMessage_GotFocus()
If mFirst = False Then
   KeyEnter 13
   mFirst = True
End If
End Sub
