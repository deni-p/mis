VERSION 5.00
Begin VB.Form frmMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9570
   Begin Man.SemeruForm Linux1 
      Height          =   6585
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   11615
      BackColor       =   16777215
      Caption         =   ":::Supplier..."
      Begin Man.SemeruGrid UserControl11 
         Height          =   2070
         Left            =   120
         TabIndex        =   13
         Top             =   2490
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   3651
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   645
         Width           =   4020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   165
         TabIndex        =   5
         Top             =   1305
         Width           =   4020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   1980
         Width           =   4020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   4350
         TabIndex        =   3
         Top             =   1965
         Width           =   4020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   4350
         TabIndex        =   2
         Top             =   1290
         Width           =   4020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   4350
         TabIndex        =   1
         Top             =   630
         Width           =   4020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No ID"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   12
         Top             =   390
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   210
         Index           =   2
         Left            =   165
         TabIndex        =   10
         Top             =   1695
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   210
         Index           =   3
         Left            =   4350
         TabIndex        =   9
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   210
         Index           =   4
         Left            =   4350
         TabIndex        =   8
         Top             =   1020
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No ID"
         Height          =   210
         Index           =   5
         Left            =   4350
         TabIndex        =   7
         Top             =   375
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
Linux1.Top = 0
Linux1.Left = 0
Linux1.Height = Me.ScaleHeight
Linux1.Width = Me.ScaleWidth
End Sub

