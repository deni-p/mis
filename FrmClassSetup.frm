VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmClassSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Class Setup"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Tag             =   "Setup Class"
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
      Height          =   2415
      Left            =   30
      ScaleHeight     =   2385
      ScaleWidth      =   9855
      TabIndex        =   10
      Tag             =   "Class Setup"
      Top             =   90
      Width           =   9885
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   120
         ScaleHeight     =   1620
         ScaleWidth      =   9525
         TabIndex        =   11
         Top             =   345
         Width           =   9555
         Begin VB.TextBox TxtBook 
            BorderStyle     =   0  'None
            DataField       =   "ClassID"
            Height          =   315
            Index           =   0
            Left            =   1935
            MaxLength       =   5
            TabIndex        =   1
            Tag             =   "ASM"
            Text            =   "2005"
            Top             =   105
            Width           =   1020
         End
         Begin VB.TextBox TxtBook 
            BorderStyle     =   0  'None
            DataField       =   "Desc"
            Height          =   315
            Index           =   1
            Left            =   1935
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "ASM"
            Text            =   "2005"
            Top             =   465
            Width           =   6060
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "AccGroupID"
            Height          =   315
            Index           =   0
            Left            =   1935
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   810
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "AccGroupID"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "InsurID"
            Height          =   315
            Index           =   1
            Left            =   1935
            TabIndex        =   8
            Tag             =   "ASM"
            Top             =   1155
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "InsurID"
            Text            =   "DataCombo1"
         End
         Begin VB.Line Line2 
            X1              =   5550
            X2              =   7995
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Shape Shape1 
            Height          =   675
            Left            =   5550
            Top             =   810
            Width           =   2460
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   5640
            TabIndex        =   9
            Top             =   1185
            Width           =   2280
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   5640
            TabIndex        =   6
            Top             =   870
            Width           =   2280
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   375
            X2              =   2010
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   375
            TabIndex        =   2
            Top             =   525
            Width           =   795
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   375
            X2              =   2205
            Y1              =   1455
            Y2              =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Class ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   7
            Top             =   1215
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   0
            Top             =   165
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   375
            TabIndex        =   4
            Top             =   870
            Width           =   1275
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   375
            X2              =   2205
            Y1              =   405
            Y2              =   405
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   375
            X2              =   2205
            Y1              =   1110
            Y2              =   1110
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   2550
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmClassSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RcGroup As New DBQuick
Dim RcIns As New DBQuick

Private Sub DataCombo1_Change(Index As Integer)
LblAccount(Index) = DataCombo1(Index).BoundText
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
LblAccount(Index) = DataCombo1(Index).BoundText
End Sub

Private Sub DataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
RcGroup.DBOpen "SELECT NoAccount AS AccGroupID, AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(0).RowSource = RcGroup.DBRecordset

RcIns.DBOpen "SELECT     NoAccount AS InsurID, AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(1).RowSource = RcIns.DBRecordset

With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmClassSetup
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  * From SetupClassData"
End With
'Check1.BackColor = &HEAAF6F
'Check1.ForeColor = &H80000005
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcIns.CloseDB
Set RcIns = Nothing

RcGroup.CloseDB
Set RcGroup = Nothing
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmClassSetup = Nothing
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " INSERT INTO SetupClassData" & _
                      " ([ClassID], [Desc], AccGroupID, InsurID)" & _
                      " VALUES (N'" & TxtBook(0) & "', N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "' , N'" & DataCombo1(0).BoundText & "', N'" & DataCombo1(1).BoundText & "')"
MyDDE.PrepareUpdate = " UPDATE SetupClassData " & _
                      " SET  [Desc] =N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "', AccGroupID = N'" & DataCombo1(0).BoundText & "', InsurID = N'" & DataCombo1(1).BoundText & "'" & _
                      " WHERE ([ClassID]= '" & TxtBook(0) & "')"
MyDDE.PrepareDelete = " DELETE FROM SetupClassData WHERE  ([ClassID] = '" & TxtBook(0) & "')"
Err.Clear
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               TxtBook(0).SetFocus
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbEdit:
       
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               TxtBook(0).Enabled = False
               TxtBook(1).SetFocus
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
'               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
'               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
''                  PrepareQuery
'               Else
'                  MyDDE.CancelTrans = True
''                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
'                  MyDDE.IsChildMemberReady = False
'               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
'               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If

End Select
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

'Private Sub Text1_Change()
'If Text1 = "" Then Text1 = Format(Year(Date), "000#")
'End Sub

Private Sub TxtBook_Change(Index As Integer)
If TxtBook(Index) = "" Then TxtBook(Index) = "-"
End Sub

Private Sub TxtBook_GotFocus(Index As Integer)
Block TxtBook(Index)
End Sub

Private Sub TxtBook_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

