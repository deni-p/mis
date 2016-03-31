VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFilterPeriode 
   BackColor       =   &H00FCF1ED&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Periode"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFilterPeriode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Lanjut"
      Height          =   390
      Left            =   3645
      TabIndex        =   3
      Top             =   1425
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   5790
      TabIndex        =   4
      Top             =   0
      Width           =   5790
      Begin VB.ComboBox Combo2 
         Height          =   330
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   30
         Width           =   2445
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAAF6F&
         Height          =   135
         Left            =   30
         TabIndex        =   5
         Top             =   1125
         Width           =   5505
      End
      Begin MSComCtl2.DTPicker DtFilter 
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   810
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   52035587
         CurrentDate     =   38383
      End
      Begin MSComCtl2.DTPicker DtFilter 
         Height          =   315
         Index           =   1
         Left            =   3030
         TabIndex        =   2
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   52035587
         CurrentDate     =   38383
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Selesai"
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   2
         Left            =   3030
         TabIndex        =   8
         Top             =   510
         Width           =   1245
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Mulai"
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   510
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   75
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmFilterPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Combo2_Click()

OpenPeriodeBerjalan
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmdOk_Click
End Sub

Private Sub DtFilter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmdOk_Click
End Sub

Private Sub Form_Load()
Set Me.Picture = MainMenu.Picture
On Error Resume Next
Dim I As Integer
Combo2.Clear
For I = 1 To 12
    Combo2.AddItem Format(DateSerial(Year(Date), I, 1), "MMMM")
Next I
Combo2.ListIndex = mVarPeriode - 1
Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFilterPeriode = Nothing
End Sub

Private Sub OpenPeriodeBerjalan()
Dim rcPer As New DBQuick
rcPer.DBOpen " SELECT     BeginDate, EndDate FROM     SettingPeriod WHERE     (LEFT(GlFile, 4) = N'" & TahunFiskalYear & "') AND (Periode = " & Combo2.ListIndex + 1 & ")", CNN, lckLockReadOnly
With rcPer.DBRecordset
     If .Recordcount <> 0 Then
        DtFilter(0).Value = Format(IIf(Not IsNull(.Fields("BeginDate")), .Fields("BeginDate"), Date), "dd/mm/yyyy")
        DtFilter(1).Value = Format(IIf(Not IsNull(.Fields("EndDate")), .Fields("EndDate"), Date), "dd/mm/yyyy")
     End If
End With
mVarTempPeriode = Combo2.ListIndex + 1
rcPer.CloseDB
End Sub
