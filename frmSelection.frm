VERSION 5.00
Begin VB.Form frmSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilih Salah Satu"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   4380
      TabIndex        =   16
      Top             =   4785
      Width           =   4380
      Begin VB.CommandButton CmdTombol 
         Caption         =   "Preview"
         Height          =   555
         Index           =   0
         Left            =   105
         Picture         =   "frmSelection.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   -45
         TabIndex        =   17
         Top             =   0
         Width           =   4575
      End
      Begin VB.CommandButton CmdTombol 
         Caption         =   "&Keluar"
         Height          =   555
         Index           =   1
         Left            =   825
         Picture         =   "frmSelection.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   4380
      TabIndex        =   14
      Top             =   0
      Width           =   4380
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         Top             =   120
         Width           =   2115
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EAAF6F&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Caption         =   " Seleksi Layout "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   15
         Top             =   525
         Width           =   4095
         Begin VB.OptionButton opLay 
            Caption         =   "10"
            Height          =   630
            Index           =   10
            Left            =   2115
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3105
            Width           =   1770
         End
         Begin VB.OptionButton opLay 
            Caption         =   "9"
            Height          =   630
            Index           =   9
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3105
            Width           =   1725
         End
         Begin VB.OptionButton opLay 
            Caption         =   "8"
            Height          =   630
            Index           =   8
            Left            =   2115
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2400
            Width           =   1770
         End
         Begin VB.OptionButton opLay 
            Caption         =   "7"
            Height          =   630
            Index           =   7
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2400
            Width           =   1725
         End
         Begin VB.OptionButton opLay 
            Caption         =   "6"
            Height          =   630
            Index           =   6
            Left            =   2115
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1710
            Width           =   1770
         End
         Begin VB.OptionButton opLay 
            Caption         =   "5"
            Height          =   630
            Index           =   5
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1710
            Width           =   1725
         End
         Begin VB.OptionButton opLay 
            Caption         =   "4"
            Height          =   630
            Index           =   4
            Left            =   2115
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1020
            Width           =   1770
         End
         Begin VB.OptionButton opLay 
            Caption         =   "3"
            Height          =   630
            Index           =   3
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1020
            Width           =   1725
         End
         Begin VB.OptionButton opLay 
            Caption         =   "2"
            Height          =   630
            Index           =   2
            Left            =   2115
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   315
            Width           =   1770
         End
         Begin VB.OptionButton opLay 
            Caption         =   "1"
            Height          =   630
            Index           =   1
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   315
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private varID As String
Private VarSQL As String
Private VarReport As String
Private VarJenis As String
Private varBerat As String
Private varSupplier As String
Dim x1 As Integer

Public Property Let ID(ByVal Value As String)
       varID = Value
End Property

Public Property Let sql(ByVal Value As String)
       VarSQL = Value
End Property

Public Property Let ReportFile(ByVal Value As String)
       VarReport = Value
End Property

Public Property Let JenisRL(ByVal Value As String)
       VarJenis = Value
End Property

Public Property Let BeratRL(ByVal Value As String)
       varBerat = Value
End Property

Public Property Let Suppplier(ByVal Value As String)
       varSupplier = Value
End Property


Private Sub Command1_Click()
   Dim flg As Boolean
   flg = Option1.Value
   Unload Me
   If flg = True Then
       frmBarcodeLogistik.PrintBarcode varID, VarJenis, varBerat, varSupplier, x1
    Else
       Dim aReport As New utility
       aReport.CallReportView VarSQL, VarReport, ReportPath, "Tanda Terima Rumput Laut"
       Set aReport = Nothing
   End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub CmdTombol_Click(Index As Integer)
   If Index = 0 Then
      Dim flg As Boolean
      flg = Option1.Value
      Unload Me
      If flg = True Then
          frmBarcodeLogistik.PrintBarcode varID, VarJenis, varBerat, varSupplier, x1
       Else
          Dim aReport As New utility
          aReport.CallReportView VarSQL, VarReport, ReportPath, "Tanda Terima Rumput Laut"
          Set aReport = Nothing
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   HiasFormManTell Picture1, Me
   'x1 = 1
End Sub

Private Sub opLay_Click(Index As Integer)
      x1 = Index
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
      Frame1.Enabled = True
   Else
      Frame1.Enabled = False
   End If
End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then
      Frame1.Enabled = False
   Else
      Frame1.Enabled = True
   End If
End Sub

