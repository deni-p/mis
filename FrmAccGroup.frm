VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{11D78E78-0CB5-48CD-ADB4-348FD684EE87}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmAccGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acc Group"
   ClientHeight    =   5085
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
   ScaleHeight     =   5085
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Tag             =   "Account Group"
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
      Height          =   4410
      Left            =   30
      ScaleHeight     =   4380
      ScaleWidth      =   9855
      TabIndex        =   21
      Tag             =   "Class Setup"
      Top             =   60
      Width           =   9885
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   120
         ScaleHeight     =   3765
         ScaleWidth      =   9525
         TabIndex        =   22
         Top             =   345
         Width           =   9555
         Begin VB.TextBox TxtBook 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   2085
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "ASM"
            Text            =   "2005"
            Top             =   465
            Width           =   6060
         End
         Begin VB.TextBox TxtBook 
            BorderStyle     =   0  'None
            DataField       =   "GroupID"
            Height          =   315
            Index           =   0
            Left            =   2085
            MaxLength       =   5
            TabIndex        =   1
            Tag             =   "ASM"
            Text            =   "2005"
            Top             =   105
            Width           =   1020
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "DepreciationExpense"
            Height          =   315
            Index           =   0
            Left            =   2085
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
            BoundColumn     =   "DepreciationExpense"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "Depreciation Reserve"
            Height          =   315
            Index           =   1
            Left            =   2085
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   1140
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "Depreciation Reserve"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "PriorYearDepreciation"
            Height          =   315
            Index           =   2
            Left            =   2085
            TabIndex        =   9
            Tag             =   "ASM"
            Top             =   1470
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "PriorYearDepreciation"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "AssetCost"
            Height          =   315
            Index           =   3
            Left            =   2085
            TabIndex        =   11
            Tag             =   "ASM"
            Top             =   1800
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "AssetCost"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "Proceeds"
            Height          =   315
            Index           =   4
            Left            =   2085
            TabIndex        =   13
            Tag             =   "ASM"
            Top             =   2130
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "Proceeds"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "RecognizedGL"
            Height          =   315
            Index           =   5
            Left            =   2085
            TabIndex        =   15
            Tag             =   "ASM"
            Top             =   2460
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "RecognizedGL"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "Non RecognizedGL"
            Height          =   315
            Index           =   6
            Left            =   2085
            TabIndex        =   17
            Tag             =   "ASM"
            Top             =   2790
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "Non RecognizedGL"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataField       =   "Clearing"
            Height          =   315
            Index           =   7
            Left            =   2085
            TabIndex        =   19
            Tag             =   "ASM"
            Top             =   3120
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "Clearing"
            Text            =   "DataCombo1"
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   375
            X2              =   2205
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Non Recognized G/L"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   9
            Left            =   375
            TabIndex        =   16
            Top             =   2850
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clearing"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   8
            Left            =   375
            TabIndex        =   18
            Top             =   3180
            Width           =   585
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   375
            X2              =   2205
            Y1              =   3420
            Y2              =   3420
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   6
            Left            =   5760
            TabIndex        =   30
            Top             =   2850
            Width           =   2280
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   7
            Left            =   5760
            TabIndex        =   29
            Top             =   3180
            Width           =   2280
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   390
            X2              =   2220
            Y1              =   2430
            Y2              =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proceeds"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   7
            Left            =   375
            TabIndex        =   12
            Top             =   2190
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recognized G/L"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   6
            Left            =   375
            TabIndex        =   14
            Top             =   2520
            Width           =   1110
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   375
            X2              =   2205
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   5760
            TabIndex        =   28
            Top             =   2190
            Width           =   2280
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   5
            Left            =   5760
            TabIndex        =   27
            Top             =   2520
            Width           =   2280
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   375
            X2              =   2205
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prior Year Depreciation"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   5
            Left            =   375
            TabIndex        =   8
            Top             =   1530
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Assets Cost"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   375
            TabIndex        =   10
            Top             =   1860
            Width           =   855
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   375
            X2              =   2205
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   5760
            TabIndex        =   26
            Top             =   1530
            Width           =   2280
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   5760
            TabIndex        =   25
            Top             =   1860
            Width           =   2280
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
            Caption         =   "Depreciation Expenses"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   375
            TabIndex        =   4
            Top             =   870
            Width           =   1635
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
            Caption         =   "Depreciation Reserve"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   6
            Top             =   1200
            Width           =   1545
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   375
            X2              =   2205
            Y1              =   1440
            Y2              =   1440
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
            X2              =   2355
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   5760
            TabIndex        =   24
            Top             =   870
            Width           =   2280
         End
         Begin VB.Label LblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Group ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   5760
            TabIndex        =   23
            Top             =   1200
            Width           =   2280
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   20
      Top             =   4515
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmAccGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RcGroup As New DBQuick
Dim RcIns As New DBQuick
Dim RcGroupA As New DBQuick
Dim RcInsA As New DBQuick
Dim RcGroupB As New DBQuick
Dim RcInsB As New DBQuick
Dim RcGroupC As New DBQuick
Dim RcInsC As New DBQuick

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
RcGroup.DBOpen "SELECT NoAccount AS DepreciationExpense, AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(0).RowSource = RcGroup.DBRecordset

RcIns.DBOpen "SELECT     NoAccount AS [Depreciation Reserve], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(1).RowSource = RcIns.DBRecordset

RcGroupA.DBOpen "SELECT NoAccount AS PriorYearDepreciation, AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(2).RowSource = RcGroupA.DBRecordset

RcInsA.DBOpen "SELECT     NoAccount AS AssetCost, AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(3).RowSource = RcInsA.DBRecordset

RcGroupB.DBOpen "SELECT NoAccount AS Proceeds, AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(4).RowSource = RcGroupB.DBRecordset

RcInsB.DBOpen "SELECT     NoAccount AS RecognizedGL, AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(5).RowSource = RcInsB.DBRecordset

RcGroupC.DBOpen "SELECT NoAccount AS [Non RecognizedGL], AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(6).RowSource = RcGroupC.DBRecordset

RcInsC.DBOpen "SELECT     NoAccount AS [Clearing], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
Set DataCombo1(7).RowSource = RcInsC.DBRecordset

With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmAccGroup
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT  * From AcctGroup"
End With
'Check1.BackColor = &HEAAF6F
'Check1.ForeColor = &H80000005
End Sub

Private Sub DataCombo1_Change(Index As Integer)
LblAccount(Index) = DataCombo1(Index).BoundText
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
LblAccount(Index) = DataCombo1(Index).BoundText
End Sub

Private Sub DataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcIns.CloseDB
Set RcIns = Nothing

RcGroup.CloseDB
RcGroupA.CloseDB
RcInsA.CloseDB
RcGroupB.CloseDB
RcInsB.CloseDB
RcGroupC.CloseDB
RcInsC.CloseDB

Set RcGroupA = Nothing
Set RcInsA = Nothing
Set RcGroupB = Nothing
Set RcInsB = Nothing
Set RcGroupC = Nothing
Set RcInsC = Nothing
Set RcGroup = Nothing
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmAccGroup = Nothing
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " INSERT INTO AcctGroup" & _
                      " (GroupID, [Desc], DepreciationExpense, [Depreciation Reserve], PriorYearDepreciation, AssetCost, Proceeds, RecognizedGL, [Non RecognizedGL], Clearing)" & _
                      " VALUES (N'" & TxtBook(0) & "', N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "' , N'" & DataCombo1(0).BoundText & "', N'" & DataCombo1(1).BoundText & "', N'" & DataCombo1(2).BoundText & "', N'" & DataCombo1(3).BoundText & "', N'" & DataCombo1(4).BoundText & "', N'" & DataCombo1(5).BoundText & "', N'" & DataCombo1(6).BoundText & "', N'" & DataCombo1(7).BoundText & "')"
MyDDE.PrepareUpdate = " UPDATE AcctGroup " & _
                      " SET  [Desc] =N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "'," & _
                      " DepreciationExpense = N'" & DataCombo1(0).BoundText & "'," & _
                      " [Depreciation Reserve] = N'" & DataCombo1(1).BoundText & "'," & _
                      " PriorYearDepreciation = N'" & DataCombo1(2).BoundText & "'," & _
                      " AssetCost = N'" & DataCombo1(3).BoundText & "'," & _
                      " Proceeds = N'" & DataCombo1(4).BoundText & "'," & _
                      " RecognizedGL = N'" & DataCombo1(5).BoundText & "'," & _
                      " [Non RecognizedGL] = N'" & DataCombo1(6).BoundText & "'," & _
                      " [Clearing] = N'" & DataCombo1(7).BoundText & "'" & _
                      " WHERE ([GroupID]= N'" & TxtBook(0) & "')"
MyDDE.PrepareDelete = " DELETE FROM AcctGroup WHERE  ([GroupID] = N'" & TxtBook(0) & "')"
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



