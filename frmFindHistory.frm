VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFindHistory 
   BackColor       =   &H00EAAF6F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Criteria History"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmFindHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBox1 
      Appearance      =   0  'Flat
      DataField       =   "PurchaseID"
      Height          =   330
      Left            =   1365
      MaxLength       =   15
      TabIndex        =   7
      Tag             =   "PO"
      Top             =   855
      Width           =   4545
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      DataField       =   "PurchaseID"
      Height          =   330
      Left            =   1365
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "PO"
      Top             =   135
      Width           =   4545
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6075
      TabIndex        =   10
      Top             =   1320
      Width           =   6075
      Begin VB.CommandButton cmd 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   825
         Picture         =   "frmFindHistory.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   720
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   105
         Picture         =   "frmFindHistory.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   720
      End
      Begin VB.Frame FrTombol 
         BackColor       =   &H00C0FFFF&
         Height          =   30
         Left            =   -30
         TabIndex        =   11
         Top             =   0
         Width           =   9945
      End
   End
   Begin MSComCtl2.DTPicker DTP 
      DataField       =   "DatePurchase"
      Height          =   315
      Index           =   0
      Left            =   1365
      TabIndex        =   3
      Tag             =   "PO"
      Top             =   495
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   58916867
      CurrentDate     =   38272
   End
   Begin MSComCtl2.DTPicker DTP 
      DataField       =   "DatePurchase"
      Height          =   315
      Index           =   1
      Left            =   3885
      TabIndex        =   5
      Tag             =   "PO"
      Top             =   495
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   58916867
      CurrentDate     =   38272
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   165
      X2              =   2085
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   930
      Width           =   540
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   180
      X2              =   2085
      Y1              =   435
      Y2              =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   555
      Width           =   570
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   165
      X2              =   2085
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s / d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   4
      Left            =   3465
      TabIndex        =   4
      Top             =   570
      Width           =   315
   End
End
Attribute VB_Name = "frmFindHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsSelected As Boolean
Private aID, aDate, aPartner As String
Private lHistName, lPartnerType As String

Public Property Let HistoryName(vData As String)
   lHistName = vData
End Property

Public Property Let PartnerType(vData As String)
   lPartnerType = vData
End Property

Public Property Get ID() As String
   ID = aID
End Property

Public Property Get DateRange() As String
   DateRange = aDate
End Property

Public Property Get Partner() As String
   Partner = aPartner
End Property


Private Sub cmd_Click(Index As Integer)
   If Index = 0 Then
      aID = txtBox.Text
      aDate = " between '" & Format(frmFindHistory.DTP(0).Value, "yyyy-MM-dd") & "' and '" & Format(frmFindHistory.DTP(1).Value, "yyyy-MM-dd") & "' "
      aPartner = txtBox1.Text
      IsSelected = True
   Else
      IsSelected = False
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   'HiasForm Picture1, Me
   HiasFormManTell Picture2, Me
   DTP(0).Value = Now
   DTP(1).Value = Now
   IsSelected = False
   
'   lbl(0).Caption = lHistName & " ID"
'   lbl(1).Caption = "Tanggal " & lHistName
    Me.Caption = "History - " & lHistName
   lbl(2).Caption = lPartnerType
   
   If lPartnerType = "" Then
      Line1(2).Visible = False
      txtBox1.Visible = False
   Else
      Line1(2).Visible = True
      txtBox1.Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   lPartnerType = ""
   lHistName = ""
End Sub

