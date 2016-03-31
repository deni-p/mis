VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{82968C93-C596-4A47-8A14-646737648F29}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmRetirementMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retirement"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
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
   ScaleHeight     =   5865
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Tag             =   "Asset Retirement"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   45
      ScaleHeight     =   4755
      ScaleWidth      =   9780
      TabIndex        =   12
      Tag             =   "Retirement Maintenance"
      Top             =   60
      Width           =   9810
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAAF6F&
         ForeColor       =   &H80000008&
         Height          =   4110
         Left            =   90
         ScaleHeight     =   4080
         ScaleWidth      =   9600
         TabIndex        =   13
         Top             =   420
         Width           =   9630
         Begin VB.PictureBox Picture3 
            Height          =   2355
            Left            =   75
            ScaleHeight     =   2295
            ScaleWidth      =   9420
            TabIndex        =   14
            Top             =   1620
            Width           =   9480
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   2280
               Left            =   30
               TabIndex        =   10
               Top             =   15
               Width           =   9360
               _ExtentX        =   16510
               _ExtentY        =   4022
               _Version        =   393216
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   3
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
               ColumnCount     =   6
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
                  DataField       =   "CashProceed"
                  Caption         =   "Cash Proceed"
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
               BeginProperty Column04 
                  DataField       =   "NonCashProceed"
                  Caption         =   "Non-Cash Proceed"
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
               BeginProperty Column05 
                  DataField       =   "SalesExpenses"
                  Caption         =   "Expenses Of Sale"
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
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "RetireType"
            Height          =   315
            ItemData        =   "FrmRetirementMaintenance.frx":0000
            Left            =   1380
            List            =   "FrmRetirementMaintenance.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Tag             =   "ASM"
            Top             =   900
            Width           =   3660
         End
         Begin VB.TextBox txtRetirement 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Id Retire"
            Height          =   285
            Index           =   1
            Left            =   1380
            TabIndex        =   7
            Tag             =   "ASM"
            Top             =   1245
            Width           =   1965
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "RetireDate"
            Height          =   315
            Left            =   1380
            TabIndex        =   3
            Tag             =   "ASM"
            Top             =   570
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57475073
            CurrentDate     =   38612
         End
         Begin MSDataListLib.DataCombo CboUang 
            DataField       =   "CurrID"
            Height          =   330
            Left            =   6975
            TabIndex        =   9
            Tag             =   "ASM"
            Top             =   1222
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Currency Name"
            BoundColumn     =   "CurrID"
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
         Begin MSDataListLib.DataCombo cboAssets 
            DataField       =   "Assets ID"
            Height          =   330
            Left            =   1380
            TabIndex        =   1
            Tag             =   "ASM"
            Top             =   225
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
         Begin VB.Label lblAssets 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Left            =   5115
            TabIndex        =   15
            Top             =   315
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   5580
            X2              =   7080
            Y1              =   1515
            Y2              =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   4
            Left            =   5595
            TabIndex        =   8
            Top             =   1290
            Width           =   870
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   105
            X2              =   1605
            Y1              =   1515
            Y2              =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retirement Code"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   6
            Top             =   1290
            Width           =   1215
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   105
            X2              =   1605
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retirement Type"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   4
            Top             =   960
            Width           =   1200
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   105
            X2              =   1605
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retirement Date"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   2
            Top             =   630
            Width           =   1185
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   105
            X2              =   1605
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Asset ID"
            ForeColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   0
            Top             =   300
            Width           =   615
         End
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   5295
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      InitControlSet  =   1
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmRetirementMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RcInsC As New DBQuick
Dim RcAst As New DBQuick

Private Sub cboAssets_Click(Area As Integer)
lblAssets = cboAssets.BoundText
End Sub

Private Sub cboAssets_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub CboUang_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Columns(ColIndex) = "" Then DataGrid1.Columns(ColIndex) = 0
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

RcInsC.DBOpen "SELECT     CurrID, [Currency Name] FROM         [Currency Setup] ORDER BY CurrID", CNN, lckLockReadOnly
Set CboUang.RowSource = RcInsC.DBRecordset

RcAst.DBOpen "SELECT     [No Aktiva] AS [Assets ID], [Nama Aktiva] AS [Description] FROM         [Tabel Aktiva Tetap] ORDER BY [No Aktiva]", CNN, lckLockReadOnly
Set cboAssets.RowSource = RcAst.DBRecordset

With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmRetirementMaintenance
    .BindFormTAG = "ASM"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT [No Aktiva] AS [Assets ID], RetireDate, RetireType, CurrID, [Id Retire] FROM RetireMaintenance"
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
Set FrmRetirementMaintenance = Nothing
End Sub

Private Sub MyDDE_ExecuteOrder(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
On Error Resume Next
MyDDE.PrepareAppend = " INSERT INTO RetireMaintenance" & _
                      " ([Id Retire], [No Aktiva], RetireDate, RetireType, CurrID)" & _
                      " VALUES (N'" & txtRetirement(1) & "', N'" & lblAssets & "', CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), N'" & Combo1.Text & "', N'" & CboUang.BoundText & "')"
                      
MyDDE.PrepareUpdate = " UPDATE RetireMaintenance " & _
                      " set [No Aktiva] =  N'" & lblAssets & "' , RetireDate = CONVERT(DATETIME, '" & Format(DTPicker1.Value, "dd/mm/yy") & "', 3), RetireType=N'" & Combo1.Text & "', CurrID=N'" & CboUang.BoundText & "'" & _
                      " WHERE ([Id Retire]=N'" & txtRetirement(1) & "')"
MyDDE.PrepareDelete = " DELETE FROM RetireMaintenance WHERE  ([No Aktiva] = N'" & lblAssets & "') and ([Id Retire]=N'" & txtRetirement(1) & "')"
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
                       If SendDataToServer("DELETE FROM [Detail RetireMaintenance] WHERE     ([Id Retire] = N'" & txtRetirement(1) & "')") = True Then
                            .MoveFirst
                            Do
                              If .EOF Then Exit Do
                              SendDataToServer (" INSERT INTO [Detail RetireMaintenance]" & _
                                                " ([Id Retire], Quantity, Cost, [Percent], CashProceed, NonCashProceed, SalesExpenses)" & _
                                                " VALUES (N'" & .Fields("Id Retire") & "', " & .Fields("Quantity") & ", " & .Fields("Cost") & ", " & .Fields("Percent") & ", " & .Fields("CashProceed") & ", " & .Fields("NonCashProceed") & ", " & .Fields("SalesExpenses") & ")")
                              .MoveNext
                            Loop
                            .MoveLast
                       End If
                    End If
               End With
       Case tmbDetail:
            If MyDDE.CheckEmptyControl = False Then
               With MyDDE.ChildRecordset
                    .Fields("Id Retire") = txtRetirement(1)
                    .Fields("Quantity") = 0
                    .Fields("Cost") = 0
                    .Fields("Percent") = 0
                    .Fields("CashProceed") = 0
                    .Fields("NonCashProceed") = 0
                    .Fields("SalesExpenses") = 0
               End With
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
End Sub

Private Sub MyDDE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
OpenDetail MyDDE.GetFieldByName("Id Retire")
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
Rc.DBOpen " SELECT Quantity, Cost, [Percent], CashProceed, NonCashProceed, SalesExpenses, [Id Retire], [IdX Retire] FROM         [Detail RetireMaintenance] WHERE     ([Id Retire] = N'" & Param & "')", CNN, lckLockBatch
Set MyDDE.ChildRecordset = Rc.DBRecordset.Clone(adLockBatchOptimistic)
Set DataGrid1.DataSource = MyDDE.ChildRecordset
Rc.CloseDB
Set Rc = Nothing
End Sub

Private Sub txtRetirement_GotFocus(Index As Integer)
Block txtRetirement(Index)
End Sub

Private Sub txtRetirement_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub
