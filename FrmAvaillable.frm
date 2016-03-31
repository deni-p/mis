VERSION 5.00
Begin VB.Form FrmAvaillable 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quantity Availlable"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAvaillable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Quantity Availlable"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
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
      Height          =   2340
      Left            =   0
      ScaleHeight     =   2340
      ScaleWidth      =   5280
      TabIndex        =   6
      Top             =   0
      Width           =   5280
      Begin VB.TextBox txtUser 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   435
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "0"
         Top             =   105
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   765
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   1095
         Width           =   2895
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1845
         Width           =   975
      End
      Begin VB.TextBox txtUser 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1425
         Width           =   2895
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   390
         X2              =   3225
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   390
         X2              =   3060
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   390
         X2              =   3255
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity make produksi"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   11
         Top             =   810
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity on purchase"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   10
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity on hand"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   9
         Top             =   150
         Width           =   1260
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   390
         X2              =   3225
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Available"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   390
         TabIndex        =   8
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LeadTime"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   7
         Top             =   1470
         Width           =   675
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   375
         X2              =   3210
         Y1              =   1725
         Y2              =   1725
      End
   End
End
Attribute VB_Name = "FrmAvaillable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mParam As String

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmAvaillable = Nothing
End Sub

Public Property Let SetNoItem(ByVal NamaItem As String, ByVal vNewValue As String)
mParam = vNewValue
Me.Tag = mParam & " - " & NamaItem
'HiasForm Picture1, Me
OpenQuantity
End Property

Private Sub OpenQuantity()
On Error GoTo 1
Dim RcQty As New DBQuick
RcQty.DBOpen " SELECT     SUM([QTY On Hand]) AS [QTY On Hand], SUM([Qty ON Purchase]) AS [Qty ON Purchase], SUM([Qty ON Production]) AS [Qty ON Production], NoItem,  SUM(LeadTime) AS LeadTime FROM         [QTY Availlable] GROUP BY NoItem HAVING      (NoItem = N'" & mParam & "')", CNN, lckLockReadOnly
If RcQty.Recordcount <> 0 Then
   txtUser(0) = FormatNumber(IIf(Not IsNull(RcQty.DBRecordset.Fields(0)), RcQty.DBRecordset.Fields(0), 0), 0)
   txtUser(1) = FormatNumber(IIf(Not IsNull(RcQty.DBRecordset.Fields(1)), RcQty.DBRecordset.Fields(1), 0), 0)
   txtUser(2) = FormatNumber(IIf(Not IsNull(RcQty.DBRecordset.Fields(2)), RcQty.DBRecordset.Fields(2), 0), 0)
   txtUser(4) = FormatNumber(IIf(Not IsNull(RcQty.DBRecordset.Fields(4)), RcQty.DBRecordset.Fields(4), 0), 0)
Else
   txtUser(0) = 0
   txtUser(1) = 0
   txtUser(2) = 0
   txtUser(4) = 0
End If
txtUser(3) = FormatNumber(CDbl(txtUser(0)) + CDbl(txtUser(1)) + CDbl(txtUser(2)), 0)
RcQty.CloseDB
Set RcQty = Nothing
Exit Sub
1:
MessageBox Err.Description, "frmavailable_openquantity" & Err.Number, msgOkOnly, msgExclamation
End Sub

