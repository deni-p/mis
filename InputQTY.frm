VERSION 5.00
Begin VB.Form InputQTY 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Quantity Adjustment"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InputQTY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   3615
      Begin VB.Frame Frame4 
         Caption         =   "Stock Sesudah Adjustment"
         Height          =   900
         Left            =   120
         TabIndex        =   5
         Top             =   1035
         Width           =   3390
         Begin VB.Label StockEnd 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1590
            TabIndex        =   9
            Top             =   525
            Width           =   1620
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   3
            X1              =   1260
            X2              =   3210
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label LblQTY 
            BackColor       =   &H8000000C&
            Caption         =   "Current Qty"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   3
            Left            =   150
            TabIndex        =   8
            Top             =   525
            Width           =   1350
         End
         Begin VB.Label LblQTY 
            BackColor       =   &H8000000C&
            Caption         =   "Adjustment Qty"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   6
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label lblStock 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1590
            TabIndex        =   7
            Top             =   240
            Width           =   1620
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   2
            X1              =   1260
            X2              =   3210
            Y1              =   465
            Y2              =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Stock Sebelum Adjustment"
         Height          =   900
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   3390
         Begin VB.TextBox QuantityStock 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1635
            TabIndex        =   4
            Text            =   "0000"
            Top             =   510
            Width           =   1605
         End
         Begin VB.Label LblQTY 
            BackColor       =   &H8000000C&
            Caption         =   "Current Qty"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label StockAkhir 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0;(#.##0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1620
            TabIndex        =   2
            Top             =   255
            Width           =   1620
         End
         Begin VB.Label LblQTY 
            BackColor       =   &H8000000C&
            Caption         =   "Actual Qty"
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   540
            Width           =   1350
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   0
            X1              =   1275
            X2              =   3225
            Y1              =   465
            Y2              =   465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   1
            X1              =   1275
            X2              =   3225
            Y1              =   735
            Y2              =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   585
         Left            =   120
         TabIndex        =   13
         Top             =   1935
         Width           =   3390
         Begin VB.CommandButton CmdAdj 
            Caption         =   "&Update"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1425
            TabIndex        =   10
            Top             =   150
            Width           =   945
         End
         Begin VB.CommandButton CmdAdj 
            Caption         =   "C&ancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2370
            TabIndex        =   11
            Top             =   150
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "InputQTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iLoad As Boolean

Private Sub CmdAdj_Click(Index As Integer)
On Error GoTo InputErr
Dim I As Integer
Select Case Index
       Case 1:
         If CDbl(QuantityStock) <> 0 Then
            If IsNumeric(StockEnd) = True And IsNumeric(lblStock(0).Caption) = True Then
               If CDbl(lblStock(0).Caption) < 0 Then
                  'MASUK IN BILA KAKEHAN
'                  mStokAdjustment = CDbl(lblStock(0).Caption)   'Abs(CDbl(lblStock(0)))
'                  'KOLOM (DEBET)
'                  FrmClosing.DgDetailADJ.Columns(6) = mStokAdjustment 'FrmClosing.mStokAdjustment
                  FrmClosing.DgDetailADJ.Columns(6).Text = CDbl(lblStock(0).Caption)
'                  MsgBox "TRAP1", vbCritical
               Else
                  'MASUK OUT BILA KURANG
'                  FrmClosing.mStokAdjustment = CDbl(lblStock(0).Caption)
                  'KOLOM KELUAR/DEBIT
                  FrmClosing.DgDetailADJ.Columns(6).Text = CDbl(lblStock(0).Caption)
               End If
               FrmClosing.DgDetailADJ.Columns(7) = CDbl(StockEnd)
            End If
         Else
            MsgBox "Tidak ada adjustment ?", vbQuestion, "Adjustment Stock"
            FrmClosing.CancelAdjustment = True
         End If
       Case 2:
            If MsgBox("Anda yakin untuk membatalkan data Adjustment.", vbQuestion + vbYesNo, "Warning") = vbYes Then
            'If I = 6 Then
               FrmClosing.CancelAdjustment = True
            Else
               Exit Sub
            End If
End Select

Unload Me
'MsgBox "TRAP1", vbCritical
Exit Sub
InputErr:
   MsgBox Err.Description, vbCritical, "Input Qty Adjustment"
End Sub

Private Sub Form_Activate()
iLoad = True
QuantityStock.SetFocus
'FrmClosing.mStokAdjustment = 0
End Sub

Private Sub Form_Load()
iLoad = False
'FrmClosing.mStokAdjustment = 0
End Sub

Private Sub QuantityStock_Change()
If iLoad = True Then
   If QuantityStock = "" And IsNumeric(QuantityStock) = False Then
      QuantityStock = "0"
   End If
End If
If IsNumeric(QuantityStock) = True Then
   lblStock(0) = FormatNumber(CDbl(StockAkhir) - CDbl(QuantityStock), 0)
   StockEnd = FormatNumber(CDbl(QuantityStock), 0)
Else
   lblStock(0) = "0"
   StockEnd = "0"
End If
End Sub

Private Sub QuantityStock_GotFocus()
Block QuantityStock
End Sub

Private Sub QuantityStock_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then CmdAdj(1).Value = True
End Sub

Private Sub QuantityStock_KeyPress(KeyAscii As Integer)
ValidNum KeyAscii
End Sub
