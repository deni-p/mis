VERSION 5.00
Object = "{594F23A7-88F5-4C02-866B-8E877A62F75C}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form frmBackFlushing 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13170
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   13170
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   360
      ScaleHeight     =   5145
      ScaleWidth      =   9165
      TabIndex        =   1
      Top             =   120
      Width           =   9195
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   5475
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "frmBackFlushing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub
