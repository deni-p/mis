VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAbount 
   Caption         =   "About"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
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
      Height          =   6225
      Left            =   45
      ScaleHeight     =   6195
      ScaleWidth      =   10860
      TabIndex        =   10
      Top             =   60
      Width           =   10890
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   5805
         Left            =   165
         ScaleHeight     =   5775
         ScaleWidth      =   10530
         TabIndex        =   11
         Top             =   150
         Width           =   10560
         Begin TabDlg.SSTab SSTab1 
            Height          =   5430
            Left            =   255
            TabIndex        =   0
            Top             =   135
            Width           =   10080
            _ExtentX        =   17780
            _ExtentY        =   9578
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            Tab             =   1
            TabsPerRow      =   4
            TabHeight       =   520
            BackColor       =   15380335
            TabCaption(0)   =   "ProMan"
            TabPicture(0)   =   "frmAbount.frx":08CA
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame1"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Lisence To"
            TabPicture(1)   =   "frmAbount.frx":08E6
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Frame3"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Version History"
            TabPicture(2)   =   "frmAbount.frx":0902
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "Register ProMan"
            TabPicture(3)   =   "frmAbount.frx":091E
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Frame2"
            Tab(3).ControlCount=   1
            Begin VB.Frame Frame1 
               Height          =   4905
               Left            =   -74895
               TabIndex        =   20
               Top             =   405
               Width           =   9900
               Begin VB.TextBox txtLegal 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Height          =   1935
                  Left            =   1425
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   1
                  Top             =   1605
                  Width           =   7545
               End
               Begin VB.Image Image1 
                  Height          =   480
                  Left            =   180
                  Picture         =   "frmAbount.frx":093A
                  Top             =   375
                  Width           =   480
               End
               Begin VB.Shape Shape1 
                  FillColor       =   &H8000000D&
                  FillStyle       =   0  'Solid
                  Height          =   525
                  Left            =   1305
                  Top             =   240
                  Width           =   6480
               End
               Begin VB.Label lblApp 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "LblApp"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   1425
                  TabIndex        =   24
                  Top             =   825
                  Width           =   6240
               End
               Begin VB.Label lblCompany 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000E&
                  Height          =   210
                  Left            =   1425
                  TabIndex        =   23
                  Top             =   405
                  Width           =   6240
               End
               Begin VB.Label lblPart 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Part Of Bulir Software Technologies."
                  Height          =   210
                  Left            =   1425
                  TabIndex        =   22
                  Top             =   1305
                  Width           =   6240
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Powered By Idoer And Q Script."
                  Height          =   210
                  Left            =   1425
                  TabIndex        =   21
                  Top             =   3645
                  Width           =   2655
               End
            End
            Begin VB.Frame Frame2 
               Height          =   4905
               Left            =   -74895
               TabIndex        =   18
               Top             =   405
               Width           =   9900
               Begin VB.TextBox txtAct 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   720
                  MaxLength       =   20
                  TabIndex        =   8
                  Top             =   795
                  Width           =   6615
               End
               Begin VB.CommandButton cmdReg 
                  Caption         =   "Activate Number"
                  Height          =   360
                  Left            =   720
                  TabIndex        =   9
                  Top             =   1140
                  Width           =   2550
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Activation  Number"
                  Height          =   210
                  Index           =   2
                  Left            =   165
                  TabIndex        =   19
                  Top             =   360
                  Width           =   1575
               End
            End
            Begin VB.Frame Frame3 
               Height          =   4905
               Left            =   105
               TabIndex        =   12
               Top             =   405
               Width           =   9900
               Begin VB.TextBox txtLisence 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   0
                  Left            =   1755
                  TabIndex        =   2
                  Top             =   1245
                  Width           =   5940
               End
               Begin VB.TextBox txtLisence 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   1
                  Left            =   1755
                  TabIndex        =   3
                  Top             =   1575
                  Width           =   5940
               End
               Begin VB.TextBox txtLisence 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   2
                  Left            =   1755
                  TabIndex        =   4
                  Top             =   1905
                  Width           =   5940
               End
               Begin VB.TextBox txtLisence 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Index           =   3
                  Left            =   1755
                  TabIndex        =   5
                  Top             =   2235
                  Width           =   5940
               End
               Begin VB.CommandButton cmdComp 
                  Caption         =   "Update Company"
                  Height          =   360
                  Left            =   1755
                  TabIndex        =   7
                  Top             =   2940
                  Width           =   2550
               End
               Begin MSMask.MaskEdBox MaskEdBox1 
                  Height          =   330
                  Left            =   1755
                  TabIndex        =   6
                  Top             =   2565
                  Width           =   2550
                  _ExtentX        =   4498
                  _ExtentY        =   582
                  _Version        =   393216
                  Appearance      =   0
                  MaxLength       =   20
                  Format          =   "##.###.###.#-###.###"
                  Mask            =   "##.###.###.#-###.###"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Company Name"
                  Height          =   210
                  Index           =   0
                  Left            =   165
                  TabIndex        =   17
                  Top             =   1275
                  Width           =   1275
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
                  Height          =   210
                  Index           =   1
                  Left            =   165
                  TabIndex        =   16
                  Top             =   1620
                  Width           =   645
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Phone"
                  Height          =   210
                  Index           =   2
                  Left            =   165
                  TabIndex        =   15
                  Top             =   2310
                  Width           =   525
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "N.P.W.P"
                  Height          =   210
                  Index           =   3
                  Left            =   165
                  TabIndex        =   14
                  Top             =   2640
                  Width           =   690
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "City"
                  Height          =   210
                  Index           =   4
                  Left            =   165
                  TabIndex        =   13
                  Top             =   1950
                  Width           =   300
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmAbount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyEx As New clsTemp
Private MyNCmpstr, mVarString As String

Private Sub cmdComp_Click()
ReadCompany True
End Sub

Private Sub cmdComp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then SSTab1.SetFocus
End Sub

Private Sub cmdComp_LostFocus()
If SSTab1.Enabled = True Then SSTab1.SetFocus
End Sub

Private Sub cmdReg_Click()
'Select Case Index
'       Case 0:
''            txtSN(0) = MyEx.FlashRemuse(txtLisence(0), MyNCmpstr)
''            txtSN(1) = MyEx.FlashRemuse(txtLisence(1), MyNCmpstr)
''            txtSN(2) = MyEx.FlashRemuse(txtLisence(2), MyNCmpstr)
''            txtSN(3) = MyEx.FlashRemuse(txtLisence(3), MyNCmpstr)
'''            txtSN(4) = MyEx.FlashRemuse(MaskEdBox1, MyNCmpstr)
''             txtReg(0) = mTelo
'       Case 1:
'            'MsgBox Splitmom
'       Case 2:
'       Case 3:
'End Select
End Sub

Private Sub cmdReg_LostFocus()
If SSTab1.Enabled = True Then SSTab1.SetFocus
End Sub

'Private Sub Command1_Click()
'MessageBox MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "SID"), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
'End Sub

Private Sub Form_Load()
'Set Picture1.Picture = LoadResPicture(101, 0)
SSTab1.Tab = 0
Me.Caption = "About " & App.Title & " Ver. " & App.Major & "." & App.Minor & "." & App.Revision
lblCompany = App.CompanyName
lblApp.Caption = App.Comments
txtLegal = App.LegalCopyright
ReadCompany False
mVarString = txtLisence(0) & vbCrLf & txtLisence(1) & vbCrLf & txtLisence(2) & vbCrLf & txtLisence(3) & vbCrLf & MaskEdBox1
MyNCmpstr = txtLisence(0) & "-" & txtLisence(1) & "-" & txtLisence(2) & "-" & txtLisence(3) & "-" & MaskEdBox1 & "-" & mVarString
SSTab1.TabCaption(0) = App.Title
SSTab1.TabCaption(3) = "Register " & App.Title
SSTab1.Tab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
End Sub

Private Sub Form_Resize()

HiasForm Picture1, Me
CenterForm Picture2, Me
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAbount = Nothing
End Sub

Private Sub ReadCompany(ByVal Tipical As Boolean)
On Error Resume Next
Dim mCmp As String
Dim Dataku As String
mCmp = txtLisence(0) & txtLisence(1) & txtLisence(2) & txtLisence(3) & MaskEdBox1

If Tipical = False Then
   txtLisence(0) = GetSetting(App.EXEName, "Lisence Profile", "Company Name")
   txtLisence(1) = GetSetting(App.EXEName, "Lisence Profile", "Address")
   txtLisence(2) = GetSetting(App.EXEName, "Lisence Profile", "City")
   txtLisence(3) = GetSetting(App.EXEName, "Lisence Profile", "Phone")
   MaskEdBox1 = GetSetting(App.EXEName, "Lisence Profile", "NPWP")
   mVarString = GetSetting(App.EXEName, "Lisence Profile", "SID")
Else
   If mCmp <> "" Then
        SaveSetting App.EXEName, "Lisence Profile", "Company Name", UCase(txtLisence(0))
        SaveSetting App.EXEName, "Lisence Profile", "Address", UCase(txtLisence(1))
        SaveSetting App.EXEName, "Lisence Profile", "City", UCase(txtLisence(2))
        SaveSetting App.EXEName, "Lisence Profile", "Phone", UCase(txtLisence(3))
        SaveSetting App.EXEName, "Lisence Profile", "NPWP", UCase(MaskEdBox1)
        SaveSetting App.EXEName, "Lisence Profile", "SID", MyEx.FlashRemuse(Left(MyEx.GetGUID, 20), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
        Dataku = MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "CLSID KEY"), mCmp)
        If Dataku <> "TUTUP APLIKASI" Then
           SaveSetting App.EXEName, "Lisence Profile", "CLSID KEY", MyEx.FlashRemuse("TUTUP APLIKASI", UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
           SaveSetting App.EXEName, "Lisence Profile", "CLSID CMP", MyEx.FlashRemuse(Format(Date, "dd/mm/yyyy"), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
           SaveSetting App.EXEName, "Lisence Profile", "CLSID DMN", MyEx.FlashRemuse(Format(Date + 20, "dd/mm/yyyy"), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
        End If
        MyEx.SimpanfilePath App.Path & "\" & App.EXEName & ".Arj", UCase(mCmp), Format(Date, "dd/mm/yyyy")
        MessageBox "Update Data Lisence Sukses.Kemudian Restart Aplikasi " & App.EXEName, "Restart", msgOkOnly
        End
   End If
End If
Err.Clear
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)
'
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case SSTab1.Tab
       Case 0: If KeyCode = 13 Or KeyCode = vbKeyTab Then txtLegal.SetFocus
       Case 1: If KeyCode = 13 Or KeyCode = vbKeyTab Then txtLisence(0).SetFocus
       Case 2: 'If KeyCode = 13 Or KeyCode = vbTab Then txtLisence(0).SetFocus
       Case 3: If KeyCode = 13 Or KeyCode = vbKeyTab Then txtAct(0).SetFocus
End Select
End Sub

Private Sub txtLegal_Change()
'
End Sub

Private Sub txtLegal_GotFocus()
'
End Sub

Private Sub txtLegal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or vbTab Then SSTab1.SetFocus
End Sub

Private Sub txtLegal_LostFocus()
If SSTab1.Enabled = True Then SSTab1.SetFocus
End Sub

Private Sub txtLisence_GotFocus(Index As Integer)
Block txtLisence(Index)
End Sub

Private Sub txtLisence_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KeyEnter (KeyCode)
End Sub

Private Sub txtLisence_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then ValidNum KeyAscii
End Sub
