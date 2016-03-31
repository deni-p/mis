VERSION 5.00
Begin VB.Form frmBarcodeLogistik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Barcode"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBarcode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9390
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0FFFF&
      Height          =   710
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9330
      TabIndex        =   1
      Top             =   6480
      Width           =   9390
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak"
         Height          =   555
         Left            =   45
         Picture         =   "FormBarcode.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   720
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Batal"
         Height          =   555
         Left            =   765
         Picture         =   "FormBarcode.frx":D0A4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   15585
      Left            =   -15
      ScaleHeight     =   15585
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   15
      Width           =   9435
   End
End
Attribute VB_Name = "frmBarcodeLogistik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32.dll" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN = &H10& ' Draw all visible child
Private Const PRF_OWNED = &H20&    ' Draw all owned windows
Private varRL As String
Private varBerat As String
Private varSupplier As String
Private varLayout As Integer

Public Sub PrintBarcode(TxtBarcode As String, RL As String, Berat As String, Supplier As String, Optional noLayout As Integer)
On Error GoTo 1
   varRL = RL
   varBerat = Berat
   varSupplier = Supplier
   varLayout = noLayout
   If noLayout > 4 Then
      Picture1.Top = -4700
   Else
      Picture1.Top = 15
   End If
   Call DrawBarcode(TxtBarcode, Picture1)
Exit Sub
1:
MessageBox Err.Description, "frmbarcodelogistik:printbarcode" & Err.Number, msgOkOnly, msgExclamation
End Sub


Private Sub Command1_Click()

   PrintPictureBox Picture1, 10, 10
   Unload Me
End Sub

Private Sub PrintPictureBox(Box As PictureBox, _
                        Optional x As Single = 0, _
                        Optional Y As Single = 0)
On Error GoTo 2
Dim rv As Long
Dim ar As Boolean
  
On Error GoTo Exit_Sub
  
    With Box
      'Save ReDraw value
      ar = .AutoRedraw

      'Set persistance
      .AutoRedraw = True

      'Wake up printer
      Printer.Print

      'Draw controls to picture box
      rv = SendMessage(.hwnd, WM_PAINT, .hdc, 0)
      rv = SendMessage(.hwnd, WM_PRINT, .hdc, _
          PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED)

      'Refresh image to picture property
      .Picture = .Image

      'Copy picture to Printer
      Printer.PaintPicture .Picture, x, Y
      Printer.EndDoc

      'Restore backcolor  (Re-load picture if picture was used)
      Box.Line (0, 0)-(.ScaleWidth, .ScaleHeight), .BackColor, BF

      'Restore ReDraw
      .AutoRedraw = ar
    End With
  
Exit_Sub:
    If Err.Number Then MessageBox Err.Description, vbOKOnly, "Printer Error!"
2:
MessageBox Err.Description, "frmbarcodelogistik:_printpicturebox" & Err.Number, msgOkOnly, msgExclamation
End Sub

Sub DrawBarcode(ByVal bc_string As String, _
                obj As Object)
On Error GoTo 1
  Dim xpos!
  Dim Y1!
  Dim Y2!
  Dim dw%
  Dim TH!
  Dim tw
  Dim new_string$
  Dim N As Integer
  Dim C As Integer
  Dim I As Integer
  
  Dim bc_pattern As String
  
  If bc_string = "" Then
    obj.Cls
    Exit Sub
  End If

  Dim bc(90) As String
  bc(1) = "1 1221"
  bc(2) = "1 1221"
  bc(48) = "11 221"
  bc(49) = "21 112"
  bc(50) = "12 112"
  bc(51) = "22 111"
  bc(52) = "11 212"
  bc(53) = "21 211"
  bc(54) = "12 211"
  bc(55) = "11 122"
  bc(56) = "21 121"
  bc(57) = "12 121"
  bc(65) = "211 12"
  bc(66) = "121 12"
  bc(67) = "221 11"
  bc(68) = "112 12"
  bc(69) = "212 11"
  bc(70) = "122 11"
  bc(71) = "111 22"
  bc(72) = "211 21"
  bc(73) = "121 21"
  bc(74) = "112 21"
  bc(75) = "2111 2"
  bc(76) = "1211 2"
  bc(77) = "2211 1"
  bc(78) = "1121 2"
  bc(79) = "2121 1"
  bc(80) = "1221 1"
  bc(81) = "1112 2"
  bc(82) = "2112 1"
  bc(83) = "1212 1"
  bc(84) = "1122 1"
  bc(85) = "2 1112"
  bc(86) = "1 2112"
  bc(87) = "2 2111"
  bc(88) = "1 1212"
  bc(89) = "2 1211"
  bc(90) = "1 2211"
  bc(32) = "1 2121"
  bc(35) = ""
  bc(36) = "1 1 1 11"
  bc(37) = "11 1 1 1"
  bc(43) = "1 11 1 1"
  bc(45) = "1 1122"
  bc(47) = "1 1 11 1"
  bc(46) = "2 1121"
  bc(64) = ""
  bc(42) = "1 1221"
  bc_string = UCase(bc_string)
  obj.ScaleMode = 3
  obj.Cls
  obj.Picture = Nothing
  
'  '-------------------------
'  obj.CurrentX = 10
'  obj.CurrentY = 10
'  obj.FontSize = 15
'  obj.Print "RUMPUT LAUT : " & varRL
'  '-------------------------
'
'  '-------------------------
'  obj.CurrentX = 10
'  obj.CurrentY = 45
'  obj.FontSize = 15
'  obj.Print "BERAT : " & varBerat
'  '-------------------------
'
'  '-------------------------
'  obj.CurrentX = 10
'  obj.CurrentY = 75
'  obj.FontSize = 15
'  obj.Print "SUPPLIER : " & varSupplier
'  '-------------------------
  
 'dw = CInt(obj.ScaleHeight / 255)
  dw = CInt(obj.ScaleHeight / 700)  '300 untuk mengatur tebal bar
  If dw < 1 Then dw = 1
  TH = obj.TextHeight(bc_string) + 50
  tw = obj.TextWidth(bc_string)
  new_string = Chr$(1) & bc_string & Chr$(2)
  Y1 = obj.ScaleTop + GetPosYLayout(varLayout)
' Y2 = obj.ScaleTop + obj.ScaleHeight - 1.5 * TH + 50 original version

  Y2 = obj.ScaleTop + obj.ScaleHeight - 1.5 * TH - 870 + GetPosYLayout(varLayout)   ' untuk mengatur tinggi bar
  
  'obj.Width = 1.1 * Len(new_string) * (15 * dw) * obj.Width / obj.ScaleWidth
  'xpos = obj.ScaleLeft + 250  original version
  xpos = obj.ScaleLeft + GetPosXLayout(varLayout)

  For N = 1 To Len(new_string)
    C = Asc(Mid$(new_string, N, 1))

    If C > 90 Then C = 0
    bc_pattern$ = bc(C)

    For I = 1 To Len(bc_pattern$)

      Select Case Mid$(bc_pattern$, I, 1)

        Case " "
          obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
          xpos = xpos + dw

        Case "1"
          obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
          xpos = xpos + dw
          obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, BF
          xpos = xpos + dw

        Case "2"
          obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
          xpos = xpos + dw
          obj.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, BF
          xpos = xpos + 2 * dw
      End Select

    Next
  Next

  obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
  xpos = xpos + dw
  'obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
  obj.CurrentX = GetPosXLayout(varLayout) + 70  '(obj.ScaleWidth - tw) / 2
  obj.CurrentY = Y2 + 10 '+ 0.25 * Th
  obj.FontSize = 8
  obj.Print bc_string
Exit Sub
1:
MessageBox Err.Description, "frmbarcodelogistik:drawbarcode" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub


Private Function GetPosXLayout(lNo As Integer) As Integer
   Select Case lNo
      Case 1: GetPosXLayout = 30
      Case 2: GetPosXLayout = 335
      Case 3: GetPosXLayout = 30
      Case 4: GetPosXLayout = 335
      Case 5: GetPosXLayout = 30
      Case 6: GetPosXLayout = 335
      Case 7: GetPosXLayout = 30
      Case 8: GetPosXLayout = 335
      Case 9: GetPosXLayout = 30
      Case 10: GetPosXLayout = 335
   End Select
End Function

Private Function GetPosYLayout(lNo As Integer) As Integer
   Select Case lNo
      Case 1: GetPosYLayout = 30
      Case 2: GetPosYLayout = 30
      Case 3: GetPosYLayout = 180
      Case 4: GetPosYLayout = 180
      Case 5: GetPosYLayout = 330
      Case 6: GetPosYLayout = 330
      Case 7: GetPosYLayout = 480
      Case 8: GetPosYLayout = 480
      Case 9: GetPosYLayout = 630
      Case 10: GetPosYLayout = 630
   End Select
End Function


