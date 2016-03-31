VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Tag             =   "About"
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   -15
      ScaleHeight     =   6015
      ScaleWidth      =   10590
      TabIndex        =   0
      Top             =   0
      Width           =   10590
      Begin TabDlg.SSTab SSTab1 
         Height          =   4485
         Left            =   180
         TabIndex        =   1
         Top             =   165
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   7911
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   15380335
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ProMan"
         TabPicture(0)   =   "frmAbout.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Lisence To"
         TabPicture(1)   =   "frmAbout.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Version History"
         TabPicture(2)   =   "frmAbout.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Register ProMan"
         TabPicture(3)   =   "frmAbout.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame2"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   4050
            Left            =   -74895
            TabIndex        =   14
            Top             =   315
            Width           =   7905
            Begin VB.TextBox txtLisence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   330
               Index           =   4
               Left            =   1710
               TabIndex        =   27
               Top             =   2145
               Width           =   5940
            End
            Begin VB.TextBox txtLisence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   330
               Index           =   0
               Left            =   1710
               TabIndex        =   26
               Top             =   345
               Width           =   5940
            End
            Begin VB.TextBox txtLisence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   330
               Index           =   1
               Left            =   1710
               TabIndex        =   25
               Top             =   705
               Width           =   5940
            End
            Begin VB.TextBox txtLisence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   330
               Index           =   2
               Left            =   1710
               TabIndex        =   24
               Top             =   1065
               Width           =   5940
            End
            Begin VB.TextBox txtLisence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   330
               Index           =   3
               Left            =   1710
               TabIndex        =   23
               Top             =   1425
               Width           =   5940
            End
            Begin VB.CommandButton cmdComp 
               Caption         =   "Update Company"
               Height          =   360
               Left            =   1710
               TabIndex        =   16
               Top             =   2985
               Width           =   2550
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Windows NT Security (Login To SQL Database)"
               Height          =   240
               Left            =   1710
               TabIndex        =   15
               Top             =   2580
               Width           =   4470
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   330
               Left            =   1710
               TabIndex        =   28
               Top             =   1785
               Width           =   2550
               _ExtentX        =   4498
               _ExtentY        =   582
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               MaxLength       =   20
               Format          =   "##.###.###.#-###.###"
               Mask            =   "##.###.###.#-###.###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "City"
               Height          =   210
               Index           =   4
               Left            =   165
               TabIndex        =   22
               Top             =   1125
               Width           =   300
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N.P.W.P"
               Height          =   210
               Index           =   3
               Left            =   165
               TabIndex        =   21
               Top             =   1845
               Width           =   690
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               Height          =   210
               Index           =   2
               Left            =   165
               TabIndex        =   20
               Top             =   1485
               Width           =   525
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   210
               Index           =   1
               Left            =   165
               TabIndex        =   19
               Top             =   765
               Width           =   645
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company Name"
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   18
               Top             =   405
               Width           =   1275
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Server Name"
               Height          =   210
               Index           =   5
               Left            =   165
               TabIndex        =   17
               Top             =   2205
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4050
            Left            =   -74895
            TabIndex        =   10
            Top             =   315
            Width           =   7905
            Begin VB.TextBox txtAct 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
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
               Left            =   315
               MaxLength       =   20
               TabIndex        =   11
               Top             =   840
               Width           =   5805
            End
            Begin VB.CommandButton cmdReg 
               Caption         =   "Activate Number"
               Height          =   360
               Left            =   315
               TabIndex        =   12
               Top             =   1230
               Width           =   2550
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Activation  Number"
               Height          =   210
               Index           =   2
               Left            =   315
               TabIndex        =   13
               Top             =   555
               Width           =   1575
            End
         End
         Begin VB.Frame Frame1 
            Height          =   4050
            Left            =   105
            TabIndex        =   4
            Top             =   315
            Width           =   7905
            Begin VB.PictureBox Picture1 
               Height          =   780
               Left            =   150
               Picture         =   "frmAbout.frx":68C2
               ScaleHeight     =   720
               ScaleWidth      =   705
               TabIndex        =   29
               Top             =   225
               Width           =   765
            End
            Begin VB.TextBox txtLegal 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   1935
               Left            =   180
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   1605
               Width           =   7545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Powered By MMT System Team"
               Height          =   195
               Left            =   195
               TabIndex        =   9
               Top             =   3645
               Width           =   2235
            End
            Begin VB.Label lblPart 
               BackStyle       =   0  'Transparent
               Caption         =   "Part of MMT Software Technologies."
               Height          =   210
               Left            =   195
               TabIndex        =   8
               Top             =   1305
               Width           =   6240
            End
            Begin VB.Label lblCompany 
               BackStyle       =   0  'Transparent
               Caption         =   "PT. Mulia Makmur Teknologi"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   240
               Left            =   1185
               TabIndex        =   7
               Top             =   375
               Width           =   6240
            End
            Begin VB.Label lblApp 
               BackStyle       =   0  'Transparent
               Caption         =   "LblApp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1050
               TabIndex        =   6
               Top             =   825
               Width           =   6360
            End
            Begin VB.Shape Shape1 
               FillColor       =   &H8000000D&
               FillStyle       =   0  'Solid
               Height          =   525
               Left            =   1005
               Shape           =   4  'Rounded Rectangle
               Top             =   240
               Width           =   6705
            End
         End
         Begin VB.Frame Frame4 
            Height          =   4050
            Left            =   -74895
            TabIndex        =   2
            Top             =   315
            Width           =   7905
            Begin VB.TextBox Text1 
               Height          =   3675
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "frmAbout.frx":8404
               Top             =   255
               Width           =   7650
            End
         End
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyEx As New clsTemp
Private MyNCmpstr, mVarString As String

Private Type CmpInfo
       CompanyInfoStr As String
       AddressInfoStr As String
       City As String
       PhoneInfo As String
       NPWP As String
       ServerNameInfo As String
       SQLLogin As Byte
End Type

Private CompanyInfoData As CmpInfo

Private Sub cmd_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdComp_Click()
ReadCompany True
End Sub

Private Sub cmdComp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then SSTab1.SetFocus
End Sub

Private Sub cmdComp_LostFocus()
If SSTab1.Enabled = True Then SSTab1.SetFocus
End Sub

Private Sub cmdReg_LostFocus()
If SSTab1.Enabled = True Then SSTab1.SetFocus
End Sub


'Private Sub Command2_Click()
'Dim CM As New Command
'Set CM.ActiveConnection = Cnn
'CM.CommandType = adCmdStoredProc
'CM.CommandText = "@Attack"
'End Sub

'Private Sub Command1_Click()
'MessageBox MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "SID"), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
'End Sub

Private Sub Form_Load()
On Error GoTo 1
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
SSTab1.Tab = 0
Me.Tag = "About " & App.Title & " Ver. " & App.Major & "." & App.Minor & "." & App.Revision
lblCompany = App.CompanyName
lblCompany.FontBold = True
lblCompany.FontSize = 12
lblCompany.ForeColor = &H8000000E
lblApp.Caption = App.Comments
lblApp.FontBold = True
lblApp.FontSize = 12
txtLegal = App.LegalCopyright
Check1.BackColor = &H8000000F
Label2(2).BackColor = &H8000000F
'BacaCompany
ReadCompany False
mVarString = txtLisence(0) & vbCrLf & txtLisence(1) & vbCrLf & txtLisence(2) & vbCrLf & txtLisence(3) & vbCrLf & MaskEdBox1
MyNCmpstr = txtLisence(0) & "-" & txtLisence(1) & "-" & txtLisence(2) & "-" & txtLisence(3) & "-" & MaskEdBox1 & "-" & mVarString
SSTab1.TabCaption(0) = App.Title
SSTab1.TabCaption(3) = "Register " & App.Title
SSTab1.Tab = 0
'Text1.Text = "Ver " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & _
" 1.  Update activex Dataobject for Transactional Database. " & Chr(13) & _
" 2.  Update system manufacture for Accounting Purpose. " & Chr(13) & _
" 3.  Update Otorisation User Personal Access. " & Chr(13) & _
" 4.  Value And Format Report Bug Fix. " & Chr(13) & _
" 5.  Create Customize Report Viewer. " & Chr(13) & _
" 6.  Checked For All Bug System And Technical Update. " & Chr(13) & _
" 7.  Validasi Journal For Accounting Purpose. " & Chr(13) & _
" 8.  Increase Dataobject For business Purpose. " & Chr(13) & _
" " & _
" ************************************************************************* " & _
" " & _
" Please Report All Bug In This Software Manufacture To Mulia Makmur Teknologi. " & Chr(13) & _
" MMT System Team "

Exit Sub
1:
MessageBox Err.Description, "frmabout:form_load" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAbout = Nothing
End Sub

Private Sub ReadCompany(ByVal Tipical As Boolean)
On Error GoTo 1
Dim mCmp As String
Dim Dataku As String
mCmp = txtLisence(0) & txtLisence(1) & txtLisence(2) & txtLisence(3) & MaskEdBox1

If Tipical = False Then
   txtLisence(0) = GetSetting(App.EXEName, "Lisence Profile", "Company Name")
   txtLisence(1) = GetSetting(App.EXEName, "Lisence Profile", "Address")
   txtLisence(2) = GetSetting(App.EXEName, "Lisence Profile", "City")
   txtLisence(3) = GetSetting(App.EXEName, "Lisence Profile", "Phone")
   txtLisence(4) = GetSetting(App.EXEName, "Lisence Profile", "Servername")
   Check1.Value = BoolToInt(CBool(GetSetting(App.EXEName, "Lisence Profile", "SQL Security")))
   MaskEdBox1 = GetSetting(App.EXEName, "Lisence Profile", "NPWP")
   mVarString = GetSetting(App.EXEName, "Lisence Profile", "SID")
Else
   If mCmp <> "" Then
        SaveSetting App.EXEName, "Lisence Profile", "Company Name", UCase(txtLisence(0))
        SaveSetting App.EXEName, "Lisence Profile", "Address", UCase(txtLisence(1))
        SaveSetting App.EXEName, "Lisence Profile", "City", UCase(txtLisence(2))
        SaveSetting App.EXEName, "Lisence Profile", "Phone", UCase(txtLisence(3))
        SaveSetting App.EXEName, "Lisence Profile", "ServerName", UCase(txtLisence(4))
        SaveSetting App.EXEName, "Lisence Profile", "NPWP", UCase(MaskEdBox1)
        SaveSetting App.EXEName, "Lisence Profile", "SQL Security", CBool(Check1.Value)
        SaveSetting App.EXEName, "Lisence Profile", "SID", MyEx.FlashRemuse(Left(MyEx.GetGUID, 20), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
        Dataku = MyEx.SlowRemuse(GetSetting(App.EXEName, "Lisence Profile", "CLSID KEY"), mCmp)
        If Dataku <> "TUTUP APLIKASI" Then
           SaveSetting App.EXEName, "Lisence Profile", "CLSID KEY", MyEx.FlashRemuse("TUTUP APLIKASI", UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
           SaveSetting App.EXEName, "Lisence Profile", "CLSID CMP", MyEx.FlashRemuse(Format(Date, "dd/mm/yyyy"), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
           SaveSetting App.EXEName, "Lisence Profile", "CLSID DMN", MyEx.FlashRemuse(Format(Date + 20, "dd/mm/yyyy"), UCase(txtLisence(0)) & UCase(txtLisence(1)) & UCase(txtLisence(2)) & UCase(txtLisence(3)) & UCase(MaskEdBox1))
        End If
        MyEx.SimpanfilePath App.Path & "\" & App.EXEName & ".dll", UCase(mCmp), Format(Date, "dd/mm/yyyy")
        MessageBox "Update Data Lisence Sukses.Kemudian Restart Aplikasi " & App.EXEName, "Restart", msgOkOnly, msgInfo
        End
   End If
End If
Err.Clear
Exit Sub
1:
MessageBox Err.Description, "formabout_readcompany" & Err.Number, msgOkOnly, msgExclamation
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)
'
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo 1
If SSTab1.Tab = 1 Then
    If UCase(MainMenu.StatusBar1.Panels(1)) = "ADMINISTRATOR" Or UCase(MainMenu.StatusBar1.Panels(1)) = "SA" Then
        txtLisence(0).Enabled = True
        txtLisence(1).Enabled = True
        txtLisence(2).Enabled = True
        txtLisence(3).Enabled = True
        txtLisence(4).Enabled = True
        MaskEdBox1.Enabled = True
        Check1.Enabled = True
        cmdComp.Enabled = True
    Else
        txtLisence(0).Enabled = False
        txtLisence(1).Enabled = False
        txtLisence(2).Enabled = False
        txtLisence(3).Enabled = False
        txtLisence(4).Enabled = False
        MaskEdBox1.Enabled = False
        Check1.Enabled = False
        cmdComp.Enabled = False
        MessageBox "User " & MainMenu.StatusBar1.Panels(1) & " tidak memepunyai hak access untuk merubah Lisence", "Informasi", msgOkOnly, msgInfo
    End If
End If
Exit Sub
1:

End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case SSTab1.Tab
       Case 0: If KeyCode = 13 Or KeyCode = vbKeyTab Then txtLegal.SetFocus
       Case 1: If KeyCode = 13 Or KeyCode = vbKeyTab Then txtLisence(0).SetFocus
       Case 2: 'If KeyCode = 13 Or KeyCode = vbTab Then txtLisence(0).SetFocus
       Case 3: If KeyCode = 13 Or KeyCode = vbKeyTab Then txtAct(0).SetFocus
End Select
End Sub

Private Sub txtLegal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyTab Then SSTab1.SetFocus
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

'Private Sub BacaCompany()
'Dim Rc As New DBQuick
'Rc.DBOpen "Select * from [Company Info]", CNN, lckLockReadOnly
'With Rc.DBRecordset
'''     If .Recordcount <> 0 Then
'''        CompanyInfoData.CompanyInfoStr = IIf(Not IsNull(.Fields(0)), .Fields(0), "")
'''        CompanyInfoData.AddressInfoStr = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
'''        CompanyInfoData.PhoneInfo = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
'''        CompanyInfoData.City = IIf(Not IsNull(.Fields(1)), .Fields(1), "")
'''        txtLisence(3) = GetSetting(App.EXEName, "Lisence Profile", "Phone")
'''        txtLisence(4) = GetSetting(App.EXEName, "Lisence Profile", "Servername")
'''
'''     End If
'End With
'End Sub
