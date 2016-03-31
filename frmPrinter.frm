VERSION 5.00
Object = "{341455FA-3231-4678-9675-13EA48167D30}#2.0#0"; "SemeruDC.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmPrinter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   7410
   Begin SemeruDC.SemeruForm SemeruForm1 
      Height          =   7350
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   12965
      BackColor       =   16777215
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   5730
         Left            =   120
         TabIndex        =   1
         Top             =   315
         Width           =   4380
         Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
            Height          =   1815
            Left            =   90
            TabIndex        =   2
            Top             =   255
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   3201
            SectionData     =   "frmPrinter.frx":0000
         End
      End
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
''
End Sub

Private Sub Form_Resize()
SemeruForm1.Left = 0
SemeruForm1.Top = 0
SemeruForm1.Width = Me.ScaleWidth
SemeruForm1.Height = Me.ScaleHeight
Frame1.Left = 0
Frame1.Top = 0
Frame1.Width = Me.ScaleWidth
Frame1.Height = Me.ScaleHeight
'SemeruPanels1.Width = SemeruForm1.Width
'SemeruPanels1.Height = SemeruForm1.Height
'ARViewer21.Left = 0
'ARViewer21.Top = 0
'ARViewer21.Height = SemeruForm1.Height
'ARViewer21.Width = SemeruForm1.Width
'CenterForm SemeruPanels1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub
