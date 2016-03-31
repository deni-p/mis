VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmBookSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Setup"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Tag             =   "Book Setup"
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
      Height          =   2970
      Left            =   60
      ScaleHeight     =   2940
      ScaleWidth      =   9855
      TabIndex        =   10
      Top             =   60
      Width           =   9885
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   120
         ScaleHeight     =   2160
         ScaleWidth      =   9525
         TabIndex        =   11
         Top             =   345
         Width           =   9555
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "CurrFisYear"
            Height          =   315
            Left            =   1935
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   810
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "CurrFisYear"
            BoundColumn     =   "CurrFisYear"
            Text            =   "DataCombo1"
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00EAAF6F&
            Caption         =   "Auto Add Book Info"
            DataField       =   "StatusAdd"
            ForeColor       =   &H80000005&
            Height          =   195
            Left            =   375
            TabIndex        =   8
            Tag             =   "ASM"
            Top             =   1530
            Width           =   2010
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "DeprePeriod"
            Height          =   315
            ItemData        =   "FrmBookSetup.frx":0000
            Left            =   1935
            List            =   "FrmBookSetup.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   1155
            Width           =   2235
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
            Width           =   3660
         End
         Begin VB.TextBox TxtBook 
            BorderStyle     =   0  'None
            DataField       =   "BookID"
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
         Begin VB.Line Line1 
            Index           =   1
            X1              =   375
            X2              =   2205
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   375
            X2              =   2205
            Y1              =   405
            Y2              =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fiscal Year"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   375
            TabIndex        =   4
            Top             =   870
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Book ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   0
            Top             =   165
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depreciation Period"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   6
            Top             =   1215
            Width           =   1395
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   375
            X2              =   2205
            Y1              =   1455
            Y2              =   1455
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   375
            X2              =   2430
            Y1              =   1770
            Y2              =   1770
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
            Index           =   4
            X1              =   375
            X2              =   2010
            Y1              =   765
            Y2              =   765
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmBookSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RcYear As New DBQuick

Private Sub DTPicker1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
RcYear.DBOpen "SELECT     LEFT(GlFile, 4) AS [CurrFisYear] FROM         SettingPeriod GROUP BY LEFT(GlFile, 4) ORDER BY LEFT(GlFile, 4)", CNN, lckLockReadOnly
Set DataCombo1.RowSource = RcYear.DBRecordset
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmBookSetup
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  * From SetupBookData"
End With
Check1.BackColor = &HEAAF6F
Check1.ForeColor = &H80000005
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcYear.CloseDB
Set RcYear = Nothing
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmQuarter = Nothing
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " INSERT INTO SetupBookData" & _
                      " (BookID, [Desc], CurrFisYear, DeprePeriod, StatusAdd)" & _
                      " VALUES (N'" & TxtBook(0) & "', N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "' , " & DataCombo1.Text & ", N'" & Combo1 & "', " & BoolToInt(CBool(Check1.Value)) & ")"
MyDDE.PrepareUpdate = " UPDATE SetupBookData " & _
                      " SET  [Desc] =N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "', CurrFisYear = " & DataCombo1.Text & ", DeprePeriod = N'" & Combo1 & "'," & _
                      " StatusAdd = " & BoolToInt(CBool(Check1.Value)) & " WHERE (BookID = '" & TxtBook(0) & "')"
MyDDE.PrepareDelete = " DELETE FROM SetupBookData WHERE     (BookID = '" & TxtBook(0) & "')"
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
