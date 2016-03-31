VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmTransferMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer Maintenance"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
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
   ScaleHeight     =   5850
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Tag             =   "Transfer Maintenance"
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
      Height          =   4845
      Left            =   30
      ScaleHeight     =   4815
      ScaleWidth      =   9750
      TabIndex        =   5
      Tag             =   "Class Setup"
      Top             =   120
      Width           =   9780
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   120
         ScaleHeight     =   4125
         ScaleWidth      =   9525
         TabIndex        =   6
         Top             =   405
         Width           =   9555
         Begin VB.PictureBox Picture3 
            Height          =   3075
            Left            =   75
            ScaleHeight     =   3015
            ScaleWidth      =   9330
            TabIndex        =   7
            Top             =   930
            Width           =   9390
            Begin MSDataGridLib.DataGrid DataGrid1 
               Bindings        =   "FrmTransferMaintenance.frx":0000
               Height          =   2970
               Left            =   45
               TabIndex        =   4
               Top             =   30
               Width           =   9255
               _ExtentX        =   16325
               _ExtentY        =   5239
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   7
               BeginProperty Column00 
                  DataField       =   "Quantity"
                  Caption         =   "Quantity"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "#.##0;(#.##0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "Cost"
                  Caption         =   "Cost"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "#.##0;(#.##0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Percent"
                  Caption         =   "Percent"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "#.##0;(#.##0)"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "NoAccount"
                  Caption         =   "NoAccount"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "LocID"
                  Caption         =   "LocID"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "LocFisID"
                  Caption         =   "LocFisID"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "StrucID"
                  Caption         =   "StrucID"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column03 
                  EndProperty
                  BeginProperty Column04 
                  EndProperty
                  BeginProperty Column05 
                  EndProperty
                  BeginProperty Column06 
                  EndProperty
               EndProperty
            End
         End
         Begin MSDataListLib.DataCombo cboAssets 
            DataField       =   "Assets ID"
            Height          =   330
            Left            =   1950
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   90
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Description"
            BoundColumn     =   "Assets ID"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   1935
            TabIndex        =   3
            Top             =   435
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   582
            _Version        =   393216
            Format          =   20774913
            CurrentDate     =   38612
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
            Caption         =   "Assets ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   0
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trasfer Date                                                            Transfer Event"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   375
            TabIndex        =   2
            Top             =   525
            Width           =   4695
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
      TabIndex        =   8
      Top             =   5280
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmTransferMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents mCall As frmCaller
Attribute mCall.VB_VarHelpID = -1
Dim RcPartner As New DBQuick
Dim RcInsC As New DBQuick
Dim RcAst As New DBQuick

Private Sub cboAssets_Click(Area As Integer)
'lblAssets = cboAssets.BoundText
End Sub

Private Sub cboAssets_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Columns(ColIndex) = "" Then DataGrid1.Columns(ColIndex) = 0
End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
OpenPartner
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DataGrid1.AllowUpdate = DTPicker1.Enabled
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_Load()
HiasForm Picture1, Me
CenterForm Picture2, Me
Set mCall = New frmCaller

'RcInsC.DBOpen "SELECT     CurrID, [Currency Name] FROM         [Currency Setup] ORDER BY CurrID", Cnn, lckLockReadOnly
'Set CboUang.RowSource = RcInsC.DBRecordset
'
RcAst.DBOpen "SELECT     [No Aktiva] AS [Assets ID], [Nama Aktiva] AS [Description] FROM         [Tabel Aktiva Tetap] ORDER BY [No Aktiva]", CNN, lckLockReadOnly
Set cboAssets.RowSource = RcAst.DBRecordset

With MyDDE
    .EditModeReplace = False
    .SetPermissions = UserDeleteDenied
    Set .BindForm = FrmTransferMaintenance
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT [No Aktiva] as [Assets ID], [Nama Aktiva] as [Description], [Transfer Date], [Transfer Event] FROM         [Tabel Aktiva Tetap] "
End With
'Check1.BackColor = &HEAAF6F
'Check1.ForeColor = &H80000005
End Sub

'Private Sub DataCombo1_Change(Index As Integer)
'LblAccount(Index) = DataCombo1(Index).BoundText
'End Sub
'
'Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
'LblAccount(Index) = DataCombo1(Index).BoundText
'End Sub

'Private Sub DataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'KeyEnter KeyCode
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RcInsC.CloseDB
Set RcInsC = Nothing
End Sub

Private Sub Form_Resize()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmTransferMaintenance = Nothing
End Sub

Private Sub mCall_RowColChange(ByVal TagForm As String, ByVal pRecordset As ADODB.Recordset)
With MyDDE.ChildRecordset
   Select Case TagForm
          Case "GENERAL ACCOUNT"
               .Fields("NoAccount") = pRecordset.Fields(0)
          Case "LOCATION"
               .Fields("LocID") = pRecordset.Fields(0)
          Case "PHYSICAL LOCATION"
               .Fields("LocFisID") = pRecordset.Fields(0)
          Case "STRUCTUR"
               .Fields("StrucID") = pRecordset.Fields(0)
   End Select
End With
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " UPDATE   [Detail Transfer] " & _
                      " set [No Aktiva] =  N'" & cboAssets.BoundText & "'" & _
                      " WHERE ([No Aktiva]= N'xxxxxx')"

MyDDE.PrepareUpdate = " UPDATE   [Detail Transfer] " & _
                      " set [No Aktiva] =  N'" & cboAssets.BoundText & "'" & _
                      " WHERE ([No Aktiva]= N'xxxxxx')"
                      
MyDDE.PrepareDelete = " DELETE FROM   [Detail Transfer] WHERE  ([No Aktiva] = N'" & cboAssets.BoundText & "') "
Err.Clear
End Sub

Private Sub MyDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            If MyDDE.IsChildMemberReady = True Then
               MyDDE.IsChildMemberReady = True
               cboAssets.SetFocus
            'Else
               'MyDDE.IsChildMemberReady = False
            End If
       Case tmbEdit:
       
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               cboAssets.Enabled = False
               DTPicker1.SetFocus
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
       Case tmbSave:
               With MyDDE.ChildRecordset
                    If .Recordcount <> 0 Then
                       If SendDataToServer("DELETE FROM   [Detail Transfer] WHERE     ([No Aktiva] = N'" & cboAssets.BoundText & "')") = True Then
                            .MoveFirst
                            Do
                              If .EOF Then Exit Do
                              SendDataToServer (" INSERT INTO   [Detail Transfer]" & _
                                                " (Quantity, Cost, [Percent], NoAccount, LocID, LocFisID, StrucID, [No Aktiva])" & _
                                                " VALUES (N'" & CDbl(.Fields("Quantity")) & "', " & CDbl(.Fields("Cost")) & ", " & CDbl(.Fields("Percent")) & ", N'" & .Fields("NoAccount") & "', N'" & .Fields("LocID") & "', N'" & .Fields("LocFisID") & "', N'" & .Fields("StrucID") & "',N'" & cboAssets.BoundText & "')")
                              .MoveNext
                            Loop
                            .MoveLast
                       End If
                    End If
               End With
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
'               With MyDDE.ChildRecordset
'                    .Fields("Id Retire") = txtRetirement(1)
'                    .Fields("Quantity") = 0
'                    .Fields("Cost") = 0
'                    .Fields("Percent") = 0
'                    .Fields("CashProceed") = 0
'                    .Fields("NonCashProceed") = 0
'                    .Fields("SalesExpenses") = 0
'               End With
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("Assets ID")
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
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True

            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Picture1.Parent.hwnd
End Sub

Private Sub OpenDetail(ByVal Param As String)
Dim Rc As New DBQuick
Rc.DBOpen " SELECT     Quantity, Cost, [Percent], NoAccount, LocID, LocFisID, StrucID , [Idx Transfer] FROM         [Detail Transfer] WHERE     ([No Aktiva] = N'" & Param & "')", CNN, lckLockBatch
Set MyDDE.ChildRecordset = Rc.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1.DataSource = MyDDE.ChildRecordset
DataGrid1.Columns(3).Button = True
DataGrid1.Columns(4).Button = True
DataGrid1.Columns(5).Button = True
DataGrid1.Columns(6).Button = True
Rc.CloseDB
Set Rc = Nothing
End Sub

Private Sub OpenPartner()
Select Case DataGrid1.Col
       Case 3: RcPartner.DBOpen "SELECT NoAccount as [Code], AccountName AS Description FROM         GLAccount", CNN, lckLockBatch
       Case 4: RcPartner.DBOpen "SELECT LocID, DescProp AS Decription FROM         SetupLoc", CNN, lckLockBatch
       Case 5: RcPartner.DBOpen "SELECT LocFisID, [Desc] AS Description FROM         SetupLocFisik", CNN, lckLockBatch
       Case 6: RcPartner.DBOpen "SELECT     StrucID, [Desc] AS Description FROM         SetupStructure", CNN, lckLockBatch
End Select
If RcPartner.Recordcount <> 0 Then
   Select Case DataGrid1.Col
          Case 3:
            mCall.FromTagActive = "GENERAL ACCOUNT"
          Case 4:
            mCall.FromTagActive = "LOCATION"
          Case 5:
            mCall.FromTagActive = "PHYSICAL LOCATION"
          Case 6:
            mCall.FromTagActive = "STRUCTUR"
   End Select
   Set mCall.FormData = RcPartner.DBRecordset
   mCall.LookUp Me
Else
   MessageBox "Data Belum Ada Atau Data Masih Kosong.", "Peringatan", msgOkOnly
   If MyDDE.ChildRecordset.Recordcount <> 0 Then
      MyDDE.ChildRecordset.CancelBatch adAffectCurrent
      If MyDDE.ChildRecordset.Recordcount <> 0 Then MyDDE.ChildRecordset.MoveLast
   End If

End If
Exit Sub
Hell:
    Err.Clear
End Sub

