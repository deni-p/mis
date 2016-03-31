VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{43E6F32B-2B03-46D3-9276-69426FE6D51B}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmPFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Press"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   9795
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   9795
      TabIndex        =   1
      Top             =   0
      Width           =   9825
      Begin VB.CommandButton cmdLink 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3930
         MaskColor       =   &H000000C0&
         Picture         =   "FrmFilter.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "SPPH"
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "No_ekstrasi"
         DataSource      =   "DDE"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "alkali"
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         DataField       =   "Grup"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   2
         Tag             =   "alkali"
         Top             =   1245
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "AT_tanggal"
         DataSource      =   "DDE"
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   5
         Tag             =   "alkali"
         Top             =   945
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   63438851
         CurrentDate     =   39365
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2235
         X2              =   195
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2190
         X2              =   150
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2205
         X2              =   165
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   1305
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "No Ekstrasi"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   735
         Width           =   2055
      End
   End
   Begin SemeruDC.SemeruOleDC mydc 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Tag             =   "minta_sampel"
      Top             =   2775
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   1005
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmPFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
