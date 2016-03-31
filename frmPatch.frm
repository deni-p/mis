VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Patch"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6375
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   1770
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   450
         Left            =   4995
         TabIndex        =   4
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   1455
         Top             =   1185
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   5730
         TabIndex        =   5
         Top             =   615
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   225
         TabIndex        =   2
         Top             =   600
         Width           =   5490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Update Server Path"
         Height          =   420
         Left            =   255
         TabIndex        =   3
         Top             =   360
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()

With CmnDialog
    .InitDir = App.Path
    .Filter = "Application|*.exe"
    .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
    .ShowOpen
    If .Filename <> "" Then
        Text1.Text = .Filename
        SaveSetting "Manufacturing Intelligent", "Data", "Patch Location", Text1.Text
    End If
End With

End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   Text1.Text = GetSetting("Manufacturing Intelligent", "Data", "Patch Location")
End Sub

