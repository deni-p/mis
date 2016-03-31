VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00F8C996-2DE8-46A8-BC86-FC76BF56E773}#1.0#0"; "SemeruDC.ocx"
Begin VB.Form FrmRouting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRouting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Tag             =   "Stages - Control Point"
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EAAF6F&
      BorderStyle     =   0  'None
      Height          =   4845
      Left            =   0
      ScaleHeight     =   4845
      ScaleWidth      =   8880
      TabIndex        =   5
      Top             =   0
      Width           =   8880
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Stage ID"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   1395
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "Partner"
         Top             =   135
         Width           =   1965
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Keterangan"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Partner"
         Top             =   465
         Width           =   3045
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         DataField       =   "Catatan"
         DataSource      =   "Adodc1"
         Height          =   855
         Index           =   2
         Left            =   1395
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Partner"
         Top             =   1140
         Width           =   5535
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "ResourcesID"
         Height          =   330
         Left            =   1395
         TabIndex        =   3
         Tag             =   "Partner"
         Top             =   795
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Description"
         BoundColumn     =   "ResourcesID"
         Text            =   "DataCombo1"
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmRouting.frx":6852
         Height          =   2625
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Tag             =   "Partner"
         Top             =   2100
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   4630
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16577005
         ForeColor       =   7159830
         HeadLines       =   1
         RowHeight       =   16
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Stage ID"
            Caption         =   "Stage ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ResourcesID"
            Caption         =   "Resouces ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Keterangan"
            Caption         =   "Keterangan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   1
         X1              =   4365
         X2              =   165
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3270
         X2              =   165
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stage ID"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   0
         Left            =   195
         TabIndex        =   10
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   1
         Left            =   195
         TabIndex        =   9
         Top             =   510
         Width           =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   2
         X1              =   4365
         X2              =   165
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resources"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   2
         Left            =   195
         TabIndex        =   8
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         ForeColor       =   &H80000005&
         Height          =   210
         Index           =   3
         Left            =   195
         TabIndex        =   7
         Top             =   1710
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         Index           =   3
         X1              =   4365
         X2              =   165
         Y1              =   1980
         Y2              =   1980
      End
   End
   Begin SemeruDC.SemeruOleDC MyDDE 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4845
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      BindFormTAG     =   "Partner"
      ActiveLanguage  =   1
   End
End
Attribute VB_Name = "FrmRouting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RcSource As New DBQuick

Private Sub DataCombo1_Click(Area As Integer)
'OpenSumberDaya
'Text1 = DataCombo1.BoundText
End Sub

Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ScanKey KeyCode, Shift, MyDDE
End Sub

Private Sub Form_Load()
'HiasForm Picture1, Me
HiasFormManTell Picture2, Me
GridLayout
OpenSumberDaya
With MyDDE
    .EditModeReplace = False
    Set .BindForm = FrmRouting
    .BindFormTAG = "Partner"
    Set .ActiveConnection = CNN
    .PrepareQuery = "SELECT     [Manufacture Stage].StageID AS [Stage ID], [Resources Type].TypeID AS [ResourcesID], [Manufacture Stage].Description AS Keterangan,  [Manufacture Stage].Note AS Catatan FROM         [Resources Type] INNER JOIN  [Manufacture Stage] ON [Resources Type].TypeID = [Manufacture Stage].TypeID ORDER BY [Manufacture Stage].StageID"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MyDDE.CheckRecordPendinged = True Then
   ScanKey vbKeyF5, 0, MyDDE
   If MyDDE.IsSucces = True Then
      Cancel = False
      MyDDE.ClearRecordset
      Set FrmRouting = Nothing
   Else
      Cancel = True
   End If
Else
   MyDDE.ClearRecordset
   Set FrmRouting = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmRouting = Nothing
End Sub

Private Sub MYDDE_AfterPrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Select Case AdReasonActiveDb
       Case tmbAddNew:
            'mVarDataDc = True
            txtBox(0).SetFocus
       Case tmbEdit:
            txtBox(0).Enabled = False
            'mVarDataDc = True
            txtBox(1).SetFocus
       Case tmbPrint:
            CallRPTReport "Stage Table.rpt"
       Case Else: 'mVarDataDc = False
End Select
End Sub

Private Sub MyDDE_PrepareActiveDB(ByVal AdReasonActiveDb As SemeruDC.TypeButtonData)
Dim mDel As New clsDelete
Select Case AdReasonActiveDb
       Case tmbDelete:
            If MyDDE.CheckEmptyControl = False Then
               If mDel.CekDelete(txtBox(0), reDelMasterCurency) = False Then
                  MyDDE.IsChildMemberReady = True
                  PrepareQuery
               Else
                  MyDDE.CancelTrans = True
                  MessageBox "Record (" & txtBox(0) & ") Sedang Dipakai Transaksi Lain." & vbCrLf & "Record Tidak Bisa DiHapus.", "Peringatan", msgOkOnly
                  MyDDE.IsChildMemberReady = False
               End If
            Else
               MyDDE.IsChildMemberReady = False
            End If
       Case tmbSave:
            If MyDDE.CheckEmptyControl = False Then
               MyDDE.IsChildMemberReady = True
               PrepareQuery
            Else
               MyDDE.IsChildMemberReady = False
            End If
End Select
Set mDel = Nothing
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveForm Picture1.Parent.hwnd
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
Block txtBox(Index)
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn) And (Index <> 2) Then KeyEnter KeyCode
End Sub

Private Sub PrepareQuery()
With MyDDE
    .PrepareAppend = " INSERT INTO [Manufacture Stage] ([StageID], [Description],[TypeID],[Note]) " & _
                     " VALUES (N'" & (txtBox(0)) & "', N'" & txtBox(1) & "',N'" & DataCombo1.BoundText & "',N'" & txtBox(2) & "')"
                     
    .PrepareUpdate = " UPDATE [Manufacture Stage] Set [Description] = N'" & txtBox(1) & "',TypeID=N'" & DataCombo1.BoundText & "',Note=N'" & txtBox(2) & "' WHERE     ([StageID] = N'" & txtBox(0) & "')"
    
    .PrepareDelete = " DELETE FROM [Manufacture Stage] WHERE   ([StageID] = N'" & txtBox(0) & "') "
End With
End Sub

Private Sub GridLayout()
DataGrid1(0).Height = 2625
DataGrid1(0).Width = 8160
DataGrid1(0).Columns(0).Width = 2355.024
DataGrid1(0).Columns(1).Width = 2399.811
DataGrid1(0).Columns(2).Width = 2835.213
End Sub

Private Sub OpenSumberDaya()
RcSource.DBOpen "SELECT     TypeID as [ResourcesID], Description  FROM         [Resources Type] ORDER BY TypeID", CNN, lckLockReadOnly
DataCombo1.BoundColumn = "ResourcesID"
DataCombo1.ListField = "Description"
Set DataCombo1.RowSource = RcSource.DBRecordset
End Sub







