VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D7BB8F75-AC9E-4E80-A526-70EA20ACFD16}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmCurrencyAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Currency Account"
   ClientHeight    =   4890
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Tag             =   "Currency Posting Account Setup"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   0
      ScaleHeight     =   4365
      ScaleWidth      =   9960
      TabIndex        =   1
      Top             =   0
      Width           =   9960
      Begin VB.TextBox TxtBook 
         BorderStyle     =   0  'None
         DataField       =   "CurrID"
         Height          =   315
         Index           =   0
         Left            =   2085
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "ASM"
         Top             =   105
         Width           =   1020
      End
      Begin VB.TextBox TxtBook 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2085
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "ASM"
         Top             =   465
         Width           =   4830
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "RGainAcc"
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   4
         Tag             =   "ASM"
         Top             =   960
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Realized Gain"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "RLossAcc"
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   5
         Tag             =   "ASM"
         Top             =   1290
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Realized Loss"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "UGainAcc"
         Height          =   315
         Index           =   2
         Left            =   2070
         TabIndex        =   6
         Tag             =   "ASM"
         Top             =   1620
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "UnRealized Gain"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "ULossAcc"
         Height          =   315
         Index           =   3
         Left            =   2070
         TabIndex        =   7
         Tag             =   "ASM"
         Top             =   1950
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "UnRealized Loss"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "FinOffset"
         Height          =   315
         Index           =   4
         Left            =   2070
         TabIndex        =   8
         Tag             =   "ASM"
         Top             =   2280
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Financial Offset"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "SalesOffset"
         Height          =   315
         Index           =   5
         Left            =   2070
         TabIndex        =   9
         Tag             =   "ASM"
         Top             =   2610
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Sales Offset"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "PurOffset"
         Height          =   315
         Index           =   6
         Left            =   2070
         TabIndex        =   10
         Tag             =   "ASM"
         Top             =   2940
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Purchase Offset"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "RoundWrite"
         Height          =   315
         Index           =   7
         Left            =   2070
         TabIndex        =   11
         Tag             =   "ASM"
         Top             =   3270
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Rounding WriteOff"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "RoundDiff"
         Height          =   315
         Index           =   8
         Left            =   2070
         TabIndex        =   12
         Tag             =   "ASM"
         Top             =   3615
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "Rounding Difference"
         Text            =   "DataCombo1"
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   5745
         TabIndex        =   32
         Top             =   1365
         Width           =   2280
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   5745
         TabIndex        =   31
         Top             =   1020
         Width           =   2280
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   375
         X2              =   2355
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
         TabIndex        =   30
         Top             =   525
         Width           =   795
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   360
         X2              =   2190
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realized Loss"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   29
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   28
         Top             =   165
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realized Gain"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   27
         Top             =   1020
         Width           =   960
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
         X1              =   360
         X2              =   2190
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   5745
         TabIndex        =   26
         Top             =   2010
         Width           =   2280
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   5745
         TabIndex        =   25
         Top             =   1680
         Width           =   2280
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   360
         X2              =   2190
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unrealized Loss"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   24
         Top             =   2010
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unrealized Gain"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   23
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   360
         X2              =   2190
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   5
         Left            =   5745
         TabIndex        =   22
         Top             =   2670
         Width           =   2280
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   4
         Left            =   5745
         TabIndex        =   21
         Top             =   2340
         Width           =   2280
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   360
         X2              =   2190
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Offset"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   20
         Top             =   2670
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Offset"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   19
         Top             =   2340
         Width           =   1125
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   375
         X2              =   2205
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   7
         Left            =   5745
         TabIndex        =   18
         Top             =   3330
         Width           =   2280
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   6
         Left            =   5745
         TabIndex        =   17
         Top             =   3000
         Width           =   2280
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   360
         X2              =   2190
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rounding Writeoff"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   16
         Top             =   3330
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchasing Offset"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   15
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   360
         X2              =   2190
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rounding Difference"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   14
         Top             =   3675
         Width           =   1470
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   360
         X2              =   2190
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Label LblAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   8
         Left            =   5745
         TabIndex        =   13
         Top             =   3675
         Width           =   2280
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmCurrencyAccount"
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
Dim RcGroupD As New DBQuick
Dim RcInsD As New DBQuick
Dim RcGroupE As New DBQuick
Dim RcInsE As New DBQuick

Private Sub Form_Load()
    'HiasForm Picture1, Me

    'MyDDE.SetPermissions = aksess.MayDo("Posting Account Setup", aksess.GetID)  'Set hak Akses
    MyDDE.SetPermissions = aksess.MayDo("Posting Account Setup")  'Set hak Akses

    HiasFormManTell Picture2, Me
    RcGroup.DBOpen "SELECT NoAccount AS [Realized Gain], AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(0).RowSource = RcGroup.DBRecordset

    RcIns.DBOpen "SELECT     NoAccount AS [Realized Loss], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(1).RowSource = RcIns.DBRecordset

    RcGroupA.DBOpen "SELECT NoAccount AS [Unrealized Gain], AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(2).RowSource = RcGroupA.DBRecordset

    RcInsA.DBOpen "SELECT     NoAccount AS [Unrealized Loss], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(3).RowSource = RcInsA.DBRecordset

    RcGroupB.DBOpen "SELECT NoAccount AS [Financial Offset], AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(4).RowSource = RcGroupB.DBRecordset

    RcInsB.DBOpen "SELECT     NoAccount AS [Sales Offset], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(5).RowSource = RcInsB.DBRecordset

    RcGroupC.DBOpen "SELECT NoAccount AS [Purchase Offset], AccountName AS Description FROM GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(6).RowSource = RcGroupC.DBRecordset

    RcInsD.DBOpen "SELECT     NoAccount AS [Rounding WriteOff], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(7).RowSource = RcInsD.DBRecordset

    RcInsE.DBOpen "SELECT     NoAccount AS [Rounding Difference], AccountName AS Description FROM         GLAccount WHERE     ([Group] = N'Detail List Account') ORDER BY NoAccount", CNN, lckLockReadOnly
    Set DataCombo1(8).RowSource = RcInsE.DBRecordset

    With MyDDE
        .EditModeReplace = False
        Set .BindForm = FrmCurrencyAccount
        .BindFormTAG = "ASM"
        Set .ActiveConnection = CNN
        .PrepareQuery = "SELECT  * From CurrencyAccount"
    End With

    'Check1.BackColor = &HEAAF6F
    'Check1.ForeColor = &H80000005
End Sub

Private Sub DataCombo1_Change(Index As Integer)
    'Debug.Print DataCombo1(Index).BoundText
    LblAccount(Index).Caption = DataCombo1(Index).BoundText
End Sub

Private Sub DataCombo1_Click(Index As Integer, _
                             Area As Integer)
    LblAccount(Index) = DataCombo1(Index).BoundText
End Sub

Private Sub DataCombo1_KeyDown(Index As Integer, _
                               KeyCode As Integer, _
                               Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
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
    'Set FrmAccGroup = Nothing
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
    On Error Resume Next
    MyDDE.PrepareAppend = " INSERT INTO CurrencyAccount" & " (CurrID, [Realized Gain], [[Realized Loss]], " & " [Unrealized Gain], [Unrealized Loss], [Financial Offset], [Sales Offset], " & " [Purchase Offset], [Rounding WriteOff], [Rounding Difference])" & " VALUES (N'" & TxtBook(0) & "', N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "' , " & " N'" & DataCombo1(0).BoundText & "', N'" & DataCombo1(1).BoundText & "', " & " N'" & DataCombo1(2).BoundText & "', N'" & DataCombo1(3).BoundText & "', " & " N'" & DataCombo1(4).BoundText & "', N'" & DataCombo1(5).BoundText & "', " & " N'" & DataCombo1(6).BoundText & "', N'" & DataCombo1(7).BoundText & "', " & " N'" & DataCombo1(8).BoundText & "')"
                      
    MyDDE.PrepareUpdate = " UPDATE CurrencyAccount " & " SET  [Desc] =N'" & IIf(TxtBook(1).Text <> "", TxtBook(1).Text, "-") & "'," & " [Realized Gain] = N'" & DataCombo1(0).BoundText & "'," & " [[Realized Loss]] = N'" & DataCombo1(1).BoundText & "'," & " [Unrealized Gain] = N'" & DataCombo1(2).BoundText & "'," & " [Unrealized Loss] = N'" & DataCombo1(3).BoundText & "'," & " [Financial Offset] = N'" & DataCombo1(4).BoundText & "'," & " [Sales Offset] = N'" & DataCombo1(5).BoundText & "'," & " [Purchase Offset] = N'" & DataCombo1(6).BoundText & "'," & " [Rounding WriteOff] = N'" & DataCombo1(7).BoundText & "'" & " [Rounding Difference] = N'" & DataCombo1(8).BoundText & "'" & " WHERE ([CurrID]= N'" & TxtBook(0) & "')"
    MyDDE.PrepareDelete = " DELETE FROM CurrencyAccount WHERE  ([CurrID] = N'" & TxtBook(0) & "')"
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

Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    'MoveForm Picture1.Parent.hwnd
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

Private Sub TxtBook_KeyDown(Index As Integer, _
                            KeyCode As Integer, _
                            Shift As Integer)
    KeyEnter KeyCode
End Sub

